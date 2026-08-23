"""
KICPA(한국공인회계사회) 구인(수습CPA) 게시판 신규 공고 감시 → 텔레그램 알림

동작 방식
---------
1. 목록 페이지(list.face)를 요청해서 게시글 표를 파싱한다.
2. 각 행에서 "번호"(게시글 순번)를 뽑아, state.json에 저장된 마지막으로 본
   번호보다 큰 행만 "신규"로 판단한다. (robots.txt 상 크롤링 자체는 허용된
   상태 - Allow: / - 이지만, 서버 부하를 줄이기 위해 10분에 1회, 목록 페이지
   1회 요청만 하는 것을 기본으로 한다.)
3. 신규 글이 있으면, 가능하면 상세페이지(ijIdNum)까지 찾아 이메일/마감일 등
   추가 정보를 붙이고, 텔레그램으로 전송한다.
4. 처리한 최신 번호를 state.json에 저장한다. (GitHub Actions에서는 워크플로우가
   이 파일을 커밋해서 다음 실행 때 이어서 비교한다.)

주의
----
- 목록 페이지의 "상세보기" 링크는 순수 <a href="..."> 가 아니라 자바스크립트로
  동작하는 것으로 보인다(예: onclick 핸들러 안에 ijIdNum이 들어있는 방식).
  이 스크립트는 정규식으로 그 숫자 ID를 최대한 찾아보고, 못 찾으면 상세 링크
  없이 목록 페이지 링크로 대체한다. 배포 후 첫 실행 로그에서 상세 링크가
  잘 잡히는지 꼭 확인해서 필요하면 ID_PATTERN 정규식을 조정할 것.
"""

import json
import os
import re
import smtplib
import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import requests
from bs4 import BeautifulSoup

LIST_URL = "https://www.kicpa.or.kr/home/jobOffrSrchNewGnrl/list.face"
DETAIL_URL = "https://www.kicpa.or.kr/home/jobOffrSrchNewGnrl/detail.face?ijIdNum={id}"

STATE_PATH = Path(__file__).parent / "state.json"
RESUME_PATH = Path(__file__).parent / "입사지원서.docx"

# 목록 페이지 HTML 안에서 게시글 고유 ID(ijIdNum, 13자리 숫자 형태 - 상세페이지
# detail.face?ijIdNum=1786323784665 에서 확인됨)를 찾기 위한 패턴.
# 정확한 자바스크립트 함수명(onclick="fn_view('...')" 등)을 모르는 상태라,
# 우선 "ijIdNum=" 형태를 먼저 찾고, 없으면 10~14자리 숫자를 폭넓게 찾는다.
# 실제 배포 후 첫 실행 로그(각 row의 id 값)를 보고 오탐이 있으면 좁혀서 조정할 것.
ID_PATTERN_STRICT = re.compile(r"ijIdNum['\"]?\s*[:=]\s*['\"]?(\d{10,})")
ID_PATTERN_LOOSE = re.compile(r"\b(\d{10,14})\b")

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

GMAIL_EMAIL = os.environ.get("GMAIL_EMAIL", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
# TEST_MODE이 "true"(기본값)인 동안은 실제 회사 담당자가 아니라 본인(GMAIL_EMAIL)
# 에게만 지원메일이 갑니다. 여러 번 받아보고 제목/본문/첨부가 정상인 걸
# 확인한 뒤에만 GitHub Secrets에서 TEST_MODE 값을 "false"로 바꾸세요.
TEST_MODE = os.environ.get("TEST_MODE", "true").strip().lower() != "false"

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465

EMAIL_SUBJECT_TEMPLATE = "{company} 수습회계사 지원 - 김경식"
EMAIL_BODY_TEMPLATE = (
    "안녕하십니까. 제60회 공인회계사 시험에 합격한 김경식입니다.\n\n"
    "{company}의 수습회계사 채용 공고를 보고 지원하게 되었습니다.\n\n"
    "감사합니다.\n"
    "김경식 드림"
)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; kicpa-watcher/1.0; +personal use)"
}


def fingerprint(row: dict) -> str:
    """행의 고유 식별자. ijIdNum이 잡히면 그걸 쓰고, 못 잡았으면 제목+회사+
    등록일 조합으로 대체 식별한다 (완벽하진 않지만 안전한 폴백)."""
    if row.get("id"):
        return f"id:{row['id']}"
    return f"fp:{row['title']}|{row['company']}|{row['posted_at']}"


def load_state() -> dict:
    if STATE_PATH.exists():
        data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        # 예전 형식(last_seen_no 기반)에서 새 형식(seen_ids 기반)으로 마이그레이션
        if "seen_ids" not in data:
            data["seen_ids"] = []
            data["_migrated_from_last_seen_no"] = True
        return data
    return {"seen_ids": []}


def save_state(state: dict) -> None:
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def fetch_list_html() -> str:
    resp = requests.get(LIST_URL, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_rows(html: str) -> list[dict]:
    """게시판 표를 파싱해서 행 리스트를 반환. 각 행: no, title, company, region,
    status, employment_type, posted_at, id(있으면)."""
    soup = BeautifulSoup(html, "html.parser")
    rows = []

    # id 후보들을 문서 순서대로 미리 뽑아둔다 (행과의 매칭은 best-effort).
    ids_in_order = ID_PATTERN_STRICT.findall(html)

    table = None
    for t in soup.find_all("table"):
        header_text = t.get_text()
        if "제목" in header_text and "회사명" in header_text:
            table = t
            break

    if table is None:
        return rows

    id_cursor = 0
    for tr in table.find_all("tr"):
        cells = tr.find_all("td")
        if len(cells) < 7:
            continue
        texts = [c.get_text(strip=True) for c in cells]
        no_text = texts[0]
        if not no_text.isdigit():
            continue

        row_id = None
        # 이 행의 <a> 태그 onclick/href 속성에서 직접 ID를 찾아본다.
        # 먼저 "ijIdNum=" 명시 패턴, 없으면 10~14자리 숫자 아무거나.
        link_tag = tr.find("a")
        if link_tag is not None:
            attr_text = " ".join(
                str(link_tag.get(attr, "")) for attr in ("onclick", "href", "data-id", "data-idnum")
            )
            m = ID_PATTERN_STRICT.search(attr_text)
            if not m:
                m = ID_PATTERN_LOOSE.search(attr_text)
            if m:
                row_id = m.group(1)

        if row_id is None and id_cursor < len(ids_in_order):
            row_id = ids_in_order[id_cursor]
            id_cursor += 1
        if row_id is None:
            # 전체 페이지에서 헐겁게 찾은 숫자들도 최후 수단으로 시도한다.
            loose_ids = ID_PATTERN_LOOSE.findall(html)
            if id_cursor < len(loose_ids):
                row_id = loose_ids[id_cursor]

        rows.append(
            {
                "no": int(no_text),
                "title": texts[1],
                "company": texts[2],
                "region": texts[3],
                "status": texts[4],
                "employment_type": texts[5],
                "posted_at": texts[6],
                "id": row_id,
            }
        )

    return rows


def fetch_detail(row_id: str) -> dict:
    """상세페이지에서 이메일/마감일 등을 best-effort로 뽑는다. 실패해도 예외를
    던지지 않고 빈 dict를 반환한다."""
    try:
        resp = requests.get(DETAIL_URL.format(id=row_id), headers=HEADERS, timeout=15)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")
        text = soup.get_text("\n", strip=True)

        detail = {}
        email_match = re.search(r"[\w\.-]+@[\w\.-]+\.\w+", text)
        if email_match:
            detail["email"] = email_match.group(0)
        deadline_match = re.search(r"마감일[^\d]*(\d{4}\.\d{2}\.\d{2})", text)
        if deadline_match:
            detail["deadline"] = deadline_match.group(1)
        return detail
    except requests.RequestException:
        return {}


def format_message(row: dict, detail: dict) -> str:
    lines = [
        "🚨 KICPA 신규 채용공고",
        "",
        f"[{row['company']}] {row['title']}",
        f"지역: {row['region']} / 고용형태: {row['employment_type']}",
        f"등록일: {row['posted_at']}",
    ]
    if detail.get("deadline"):
        lines.append(f"마감일: {detail['deadline']}")
    if detail.get("email"):
        lines.append(f"담당 이메일: {detail['email']}")
    if row.get("id"):
        lines.append("")
        lines.append(DETAIL_URL.format(id=row["id"]))
    else:
        lines.append("")
        lines.append(LIST_URL)
    return "\n".join(lines)


def send_telegram(text: str) -> None:
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("[WARN] TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID 가 설정되지 않아 전송을 건너뜁니다.")
        print(text)
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    resp = requests.post(
        url,
        json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "disable_web_page_preview": False},
        timeout=15,
    )
    if resp.status_code != 200:
        print(f"[ERROR] 텔레그램 전송 실패: {resp.status_code} {resp.text}", file=sys.stderr)


def send_application_email(row: dict, detail: dict) -> None:
    """detail에 이메일이 파싱되어 있으면 지원메일을 발송한다. TEST_MODE일 때는
    실제 회사가 아니라 본인 메일로만 보낸다."""
    if not GMAIL_EMAIL or not GMAIL_APP_PASSWORD:
        print("[WARN] GMAIL_EMAIL / GMAIL_APP_PASSWORD 가 설정되지 않아 이메일 발송을 건너뜁니다.")
        return

    recipient = detail.get("email")
    if not recipient:
        print(f"[WARN] #{row['no']} {row['title']} - 담당 이메일을 찾지 못해 자동 지원메일을 보내지 않았습니다.")
        return

    if not RESUME_PATH.exists():
        print(f"[WARN] 이력서 파일({RESUME_PATH.name})을 찾을 수 없어 이메일 발송을 건너뜁니다.")
        return

    subject = EMAIL_SUBJECT_TEMPLATE.format(company=row["company"])
    body = EMAIL_BODY_TEMPLATE.format(company=row["company"])

    actual_recipient = recipient
    if TEST_MODE:
        subject = f"[TEST] {subject} (실제 수신처였을 주소: {recipient})"
        actual_recipient = GMAIL_EMAIL

    msg = MIMEMultipart()
    msg["From"] = GMAIL_EMAIL
    msg["To"] = actual_recipient
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    with open(RESUME_PATH, "rb") as f:
        part = MIMEApplication(f.read(), Name=RESUME_PATH.name)
    part["Content-Disposition"] = f'attachment; filename="{RESUME_PATH.name}"'
    msg.attach(part)

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=15) as server:
            server.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_EMAIL, actual_recipient, msg.as_string())
        mode_note = "TEST_MODE" if TEST_MODE else "실제발송"
        print(f"[INFO] 지원메일 발송({mode_note}): #{row['no']} {row['title']} -> {actual_recipient}")
    except smtplib.SMTPException as e:
        print(f"[ERROR] 지원메일 발송 실패: {e}", file=sys.stderr)


def main() -> None:
    state = load_state()
    seen_ids = set(state.get("seen_ids", []))
    is_migration_bootstrap = state.pop("_migrated_from_last_seen_no", False)

    html = fetch_list_html()
    rows = parse_rows(html)

    if not rows:
        print("[WARN] 게시글 파싱 결과가 비어 있습니다. 페이지 구조가 바뀌었을 수 있습니다.")
        return

    for row in rows:
        row["_fp"] = fingerprint(row)

    if is_migration_bootstrap:
        # 예전(번호 비교) 방식에서 막 넘어온 첫 실행: 지금 보이는 글들을
        # 전부 "이미 확인함"으로만 기록하고, 알림은 보내지 않는다. (과거
        # 글을 전부 신규로 오인해서 한꺼번에 스팸 보내는 걸 방지)
        seen_ids.update(row["_fp"] for row in rows)
        state["seen_ids"] = sorted(seen_ids)[-500:]
        save_state(state)
        print(f"[INFO] state.json을 새 형식으로 마이그레이션했습니다. 이번 실행은 알림을 생략합니다. (등록: {len(rows)}건)")
        return

    new_rows = [r for r in rows if r["_fp"] not in seen_ids]
    # 목록은 보통 최신글이 위(번호 큰 순)로 오므로, 오래된 것부터 순서대로
    # 알림을 보내도록 번호 오름차순 정렬
    new_rows.sort(key=lambda r: r["no"])

    if not new_rows:
        print(f"[INFO] 신규 공고 없음. (확인된 글 수: {len(seen_ids)})")
        return

    for row in new_rows:
        detail = fetch_detail(row["id"]) if row.get("id") else {}
        message = format_message(row, detail)
        send_telegram(message)
        print(f"[INFO] 알림 전송: #{row['no']} {row['title']}")
        send_application_email(row, detail)
        seen_ids.add(row["_fp"])

    # seen_ids가 무한정 커지지 않도록 최근 500개만 유지
    state["seen_ids"] = sorted(seen_ids)[-500:]
    save_state(state)


if __name__ == "__main__":
    main()
