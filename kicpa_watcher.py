"""
KICPA(한국공인회계사회) 구인게시판 신규 공고 감시 → 텔레그램 알림 + 자동 지원메일

감시 대상 2개 게시판
--------------------
1. 구인(수습CPA) - jobOffrSrchNewGnrl : 이미 수습/합격자 대상으로 스코프가 좁혀진
   게시판이라 별도 필터 없이 전부 감시한다.
2. 구인(CPA) - jobOffrSrchGnrl : 경력직 등 전체 채용공고가 섞여 올라오는 일반
   게시판. 제목에 "신입" 또는 "수습"이 들어간 것만 걸러서 감시한다.

두 게시판은 표 컬럼 구성이 다르다(수습CPA는 "고용형태" 컬럼이 있고, 일반게시판은
없는 대신 "채용구분" 컬럼이 있음). 그래서 컬럼 위치를 고정하지 않고, 헤더 행의
텍스트를 읽어서 이름 기반으로 매핑한다.

동작 방식
---------
1. 각 게시판의 목록 페이지를 요청해서 표를 파싱한다 (헤더 기반 동적 매핑).
2. 일반게시판은 제목에 신입/수습 키워드가 없는 행을 걸러낸다.
3. 각 행의 게시글 ID(ijIdNum)를 뽑아 "이미 확인한 글 목록"(state.json의
   seen_ids)에 있는지로 신규 여부를 판단한다 (번호 크기 비교 방식이 아님 -
   자동삭제로 번호가 줄어들 수 있어서 번호 비교는 신뢰할 수 없다).
4. 새 게시판을 처음 추가한 시점에는, 그 게시판에 이미 있던 글들을 한꺼번에
   "신규"로 오인해서 스팸 알림을 보내지 않도록, 게시판별로 "첫 실행 1회"는
   조용히 seen_ids만 채우고 알림은 생략한다 (bootstrapped_boards로 추적).
5. 신규 글이 있으면 상세페이지에서 이메일/마감일 등을 찾아 텔레그램 알림 +
   (TEST_MODE 아니면 실제) 지원메일을 보낸다.
6. 처리 결과를 state.json에 저장하고, GitHub Actions가 이 파일을 커밋해서
   다음 실행 때 이어서 비교한다.

주의
----
- 목록 페이지의 "상세보기" 링크는 자바스크립트로 동작하는 것으로 보여서,
  정규식으로 ijIdNum 숫자를 최대한 추측해서 찾는다. 실패하면 상세 정보 없이
  목록 링크로 대체한다.
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

STATE_PATH = Path(__file__).parent / "state.json"
RESUME_PATH = Path(__file__).parent / "입사지원서.docx"

# 감시할 게시판 목록. "intern" 게시판은 기존부터 감시해오던 곳이라 seen_ids
# 식별자를 예전 형식(prefix 없음) 그대로 유지해서 이미 저장된 state.json과
# 호환되게 한다. 새로 추가하는 게시판은 board key를 prefix로 붙여 구분한다.
BOARDS = [
    {
        "key": "intern",
        "label": "KICPA 신규 채용공고",
        "list_url": "https://www.kicpa.or.kr/home/jobOffrSrchNewGnrl/list.face",
        "detail_url": "https://www.kicpa.or.kr/home/jobOffrSrchNewGnrl/detail.face?ijIdNum={id}",
        "title_keywords": None,  # 필터 없음 (이미 수습CPA 전용 게시판)
        "id_prefix": "",  # 예전 seen_ids 형식과 호환을 위해 접두어 없음
    },
    {
        "key": "general",
        "label": "KICPA 일반구인 (신입/수습)",
        "list_url": "https://www.kicpa.or.kr/home/jobOffrSrchGnrl/list.face",
        "detail_url": "https://www.kicpa.or.kr/home/jobOffrSrchGnrl/detail.face?ijIdNum={id}",
        "title_keywords": ["신입", "수습"],  # 제목에 이 중 하나라도 있어야 통과
        "id_prefix": "gnrl:",
    },
]

# 목록 페이지 HTML 안에서 게시글 고유 ID(ijIdNum, 13자리 숫자 형태 - 상세페이지
# detail.face?ijIdNum=1786323784665 에서 확인됨)를 찾기 위한 패턴.
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

EMAIL_SUBJECT_TEMPLATE = "{company} {position} 지원 - 김경식"
EMAIL_BODY_TEMPLATE = (
    "안녕하십니까. 제60회 공인회계사 시험에 합격한 김경식입니다.\n\n"
    "{company}의 {position} 채용 공고를 보고 지원하게 되었습니다.\n\n"
    "감사합니다.\n"
    "김경식 드림"
)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; kicpa-watcher/1.0; +personal use)"
}


def position_word(title: str) -> str:
    """공고 제목에서 지원메일에 쓸 직무 명칭을 뽑는다."""
    if "수습" in title:
        return "수습회계사"
    if "신입" in title:
        return "신입회계사"
    return "채용"


def fingerprint(board: dict, row: dict) -> str:
    """행의 고유 식별자. ijIdNum이 잡히면 그걸 쓰고, 못 잡았으면 제목+회사+
    등록일 조합으로 대체 식별한다."""
    prefix = board["id_prefix"]
    if row.get("id"):
        return f"{prefix}id:{row['id']}"
    return f"{prefix}fp:{row.get('title','')}|{row.get('company','')}|{row.get('posted_at','')}"


def load_state() -> dict:
    if STATE_PATH.exists():
        data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        migrated = False
        if "seen_ids" not in data:
            data["seen_ids"] = []
            migrated = True
        if "bootstrapped_boards" not in data:
            # 이 필드가 아예 없다는 건 "intern" 게시판만 감시하던 예전 버전이라는
            # 뜻이라, intern은 이미 정상 운영중이었던 걸로 간주해 부트스트랩을
            # 건너뛴다. general처럼 새로 추가되는 게시판만 부트스트랩 대상.
            data["bootstrapped_boards"] = ["intern"]
            migrated = True
        data["_migrated"] = migrated
        return data
    return {"seen_ids": [], "bootstrapped_boards": []}


def save_state(state: dict) -> None:
    state.pop("_migrated", None)
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def fetch_list_html(list_url: str) -> str:
    resp = requests.get(list_url, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_rows(html: str) -> list[dict]:
    """게시판 표를 헤더 이름 기반으로 파싱해서 행 리스트를 반환한다.
    각 행: no, title, company, region, status, employment_type, posted_at, id."""
    soup = BeautifulSoup(html, "html.parser")
    rows = []

    ids_in_order = ID_PATTERN_STRICT.findall(html)

    table = None
    for t in soup.find_all("table"):
        header_text = t.get_text()
        if "제목" in header_text and "회사명" in header_text:
            table = t
            break

    if table is None:
        return rows

    trs = table.find_all("tr")
    if not trs:
        return rows

    # 첫 번째 tr을 헤더로 간주하고 라벨을 뽑는다.
    header_cells = trs[0].find_all(["th", "td"])
    header_labels = [c.get_text(strip=True) for c in header_cells]

    id_cursor = 0
    for tr in trs[1:]:
        cells = tr.find_all("td")
        if len(cells) < 2:
            continue
        texts = [c.get_text(strip=True) for c in cells]
        row_map = dict(zip(header_labels, texts))

        no_text = row_map.get("번호", texts[0] if texts else "")
        if not no_text.isdigit():
            continue

        row_id = None
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
            loose_ids = ID_PATTERN_LOOSE.findall(html)
            if id_cursor < len(loose_ids):
                row_id = loose_ids[id_cursor]

        rows.append(
            {
                "no": int(no_text),
                "title": row_map.get("제목", ""),
                "company": row_map.get("회사명", ""),
                "region": row_map.get("지역", ""),
                "status": row_map.get("구직완료 구분") or row_map.get("채용구분") or "",
                "employment_type": row_map.get("고용형태", ""),
                "posted_at": row_map.get("등록일자", ""),
                "id": row_id,
            }
        )

    return rows


def fetch_detail(detail_url_tmpl: str, row_id: str) -> dict:
    """상세페이지에서 이메일/마감일 등을 best-effort로 뽑는다."""
    try:
        resp = requests.get(detail_url_tmpl.format(id=row_id), headers=HEADERS, timeout=15)
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


def format_message(board: dict, row: dict, detail: dict) -> str:
    lines = [
        f"🚨 {board['label']}",
        "",
        f"[{row['company']}] {row['title']}",
    ]
    extra = []
    if row.get("region"):
        extra.append(f"지역: {row['region']}")
    if row.get("employment_type"):
        extra.append(f"고용형태: {row['employment_type']}")
    elif row.get("status"):
        extra.append(f"채용구분: {row['status']}")
    if extra:
        lines.append(" / ".join(extra))
    if row.get("posted_at"):
        lines.append(f"등록일: {row['posted_at']}")
    if detail.get("deadline"):
        lines.append(f"마감일: {detail['deadline']}")
    if detail.get("email"):
        lines.append(f"담당 이메일: {detail['email']}")
    lines.append("")
    if row.get("id"):
        lines.append(board["detail_url"].format(id=row["id"]))
    else:
        lines.append(board["list_url"])
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

    position = position_word(row["title"])
    subject = EMAIL_SUBJECT_TEMPLATE.format(company=row["company"], position=position)
    body = EMAIL_BODY_TEMPLATE.format(company=row["company"], position=position)

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


def process_board(board: dict, state: dict) -> None:
    seen_ids = set(state["seen_ids"])
    is_bootstrap = board["key"] not in state["bootstrapped_boards"]

    html = fetch_list_html(board["list_url"])
    rows = parse_rows(html)

    if not rows:
        print(f"[WARN] [{board['key']}] 게시글 파싱 결과가 비어 있습니다. 페이지 구조가 바뀌었을 수 있습니다.")
        return

    keywords = board.get("title_keywords")
    if keywords:
        rows = [r for r in rows if any(kw in r["title"] for kw in keywords)]

    for row in rows:
        row["_fp"] = fingerprint(board, row)

    if is_bootstrap:
        seen_ids.update(row["_fp"] for row in rows)
        state["seen_ids"] = sorted(seen_ids)[-1000:]
        state["bootstrapped_boards"].append(board["key"])
        print(f"[INFO] [{board['key']}] 최초 감시 시작 - 현재 글 {len(rows)}건을 조용히 등록만 했습니다 (알림 생략).")
        return

    new_rows = [r for r in rows if r["_fp"] not in seen_ids]
    new_rows.sort(key=lambda r: r["no"])

    if not new_rows:
        print(f"[INFO] [{board['key']}] 신규 공고 없음. (확인된 글 수: {len(seen_ids)})")
        return

    for row in new_rows:
        detail = fetch_detail(board["detail_url"], row["id"]) if row.get("id") else {}
        message = format_message(board, row, detail)
        send_telegram(message)
        print(f"[INFO] [{board['key']}] 알림 전송: #{row['no']} {row['title']}")
        send_application_email(row, detail)
        seen_ids.add(row["_fp"])

    state["seen_ids"] = sorted(seen_ids)[-1000:]


def main() -> None:
    state = load_state()
    for board in BOARDS:
        process_board(board, state)
    save_state(state)


if __name__ == "__main__":
    main()
