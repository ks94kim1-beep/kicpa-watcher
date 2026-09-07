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
   seen_ids)에 있는지로 신규 여부를 판단한다. 번호(no) 크기나 목록에 실린
   글 개수를 기준으로 판단하지 않는다 - 자동삭제로 번호가 줄어들 수 있어서
   번호/개수 비교는 신뢰할 수 없다. 아래 로그에 찍히는 숫자는 어디까지나
   "지금까지 누적 관리 중인 글 개수" 참고용 정보일 뿐, 그 숫자로 신규 여부를
   판단하는 게 아니다.
4. ID는 이미 본 적이 있지만 등록일자(posted_at)가 이전에 기록해둔 값과
   달라진 경우는 "재등록(끌올)"로 보고, 신규 글과 동일하게 다시 알림 +
   지원메일 처리를 한다 (seen_post_dates에 글별 마지막 등록일자를 저장해서
   비교한다).
5. 새 게시판을 처음 추가한 시점에는, 그 게시판에 이미 있던 글들을 한꺼번에
   "신규"로 오인해서 스팸 알림을 보내지 않도록, 게시판별로 "첫 실행 1회"는
   조용히 seen_ids/seen_post_dates만 채우고 알림은 생략한다
   (bootstrapped_boards로 추적).
6. 신규/재등록 글이 있으면 상세페이지에서 이메일/마감일 등을 찾아 텔레그램
   알림 + (TEST_MODE 아니면 실제) 지원메일을 보낸다.
7. 처리 결과를 state.json에 저장하고, GitHub Actions가 이 파일을 커밋해서
   다음 실행 때 이어서 비교한다.

주의
----
- 목록 페이지의 "상세보기" 링크는 자바스크립트로 동작하는 것으로 보여서,
  정규식으로 ijIdNum 숫자를 최대한 추측해서 찾는다. 실패하면 상세 정보 없이
  목록 링크로 대체한다.
- 이 파일을 처음 배포한 시점 기준으로, 그 이전부터 state.json에 있던 글들은
  seen_post_dates에 기준 등록일자가 없다. 그런 글들은 배포 후 첫 실행에서
  현재 등록일자를 조용히 기준값으로 채워 넣기만 하고, 재등록 알림은 그 다음
  변화가 감지될 때부터 정상 동작한다.
"""

import imaplib
import json
import os
import re
import smtplib
import sys
from datetime import datetime, timedelta, timezone
from email import message_from_bytes
from email.header import decode_header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

# 상세페이지에 이런 확장자의 첨부파일이 있으면 "회사가 자체 지원서 양식을
# 요구하는 것"으로 보고, 범용 이력서를 자동으로 보내지 않는다. docx/hwp만
# 대상으로 한다 - pdf/xlsx/pptx/zip 등은 회사소개 자료나 공고 원문 PDF인
# 경우가 많아서 여기 포함하면 오탐(정상 발송 건너뜀)이 너무 잦아진다.
ATTACHMENT_EXT_PATTERN = re.compile(r"\.(docx?|hwpx?)$", re.IGNORECASE)

STATE_PATH = Path(__file__).parent / "state.json"
RESUME_PATH = Path(__file__).parent / "입사지원서.docx"
RESUME_PDF_PATH = Path(__file__).parent / "입사지원서.pdf"

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
IMAP_HOST = "imap.gmail.com"
IMAP_PORT = 993
# 답장 확인 시 이 기간(일) 이내에 받은 메일만 훑는다. 너무 오래된 지원 건에
# 대한 답장까지 매번 다 뒤질 필요는 없어서 범위를 제한한다.
REPLY_CHECK_LOOKBACK_DAYS = 45

EMAIL_SUBJECT_TEMPLATE = "{company} {position} 지원 - 김경식"
EMAIL_BODY_TEMPLATE = (
    "안녕하십니까. 제60회 공인회계사 시험에 합격한 김경식입니다.\n\n"
    "{company}의 {position} 공고를 보고 지원하게 되었습니다.\n\n"
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
    등록일 조합으로 대체 식별한다. (등록일이 fp에 포함되므로, ID가 안 잡히는
    글은 등록일이 바뀌는 순간 자동으로 "새 글"처럼 처리된다 - 재등록 감지가
    이미 내장되어 있는 셈. ID가 잡히는 일반적인 경우의 재등록 감지는
    seen_post_dates로 별도 처리한다.)"""
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
        if "seen_post_dates" not in data:
            # 재등록(끌올) 감지를 위한 필드. 이 필드가 없던 예전 state.json에는
            # 기준 등록일자가 없으므로, 이번 실행에서 지금 보이는 값들로
            # 조용히 채워 넣고 다음 변화부터 재등록으로 인식한다.
            data["seen_post_dates"] = {}
            migrated = True
        data.setdefault("sent_applications", [])
        data.setdefault("notified_reply_ids", [])
        data["_migrated"] = migrated
        return data
    return {
        "seen_ids": [],
        "bootstrapped_boards": [],
        "seen_post_dates": {},
        "sent_applications": [],
        "notified_reply_ids": [],
    }


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
    """상세페이지에서 이메일/마감일/첨부파일 등을 best-effort로 뽑는다."""
    try:
        detail_url = detail_url_tmpl.format(id=row_id)
        resp = requests.get(detail_url, headers=HEADERS, timeout=15)
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

        # 첨부파일(자체 지원서 양식 등) 탐지: <a> 태그의 href나 텍스트가
        # docx/hwp/xlsx/pdf/zip 같은 파일 확장자로 끝나면 첨부파일로 간주한다.
        attachments = []
        seen_urls = set()
        for a in soup.find_all("a"):
            href = (a.get("href") or "").strip()
            link_text = a.get_text(strip=True)
            candidate = link_text if ATTACHMENT_EXT_PATTERN.search(link_text) else href
            if not ATTACHMENT_EXT_PATTERN.search(candidate):
                continue
            abs_url = urljoin(detail_url, href) if href and href != "#" else ""
            dedup_key = abs_url or link_text
            if dedup_key in seen_urls:
                continue
            seen_urls.add(dedup_key)
            attachments.append({"name": link_text or candidate, "url": abs_url})

        if attachments:
            detail["attachments"] = attachments

        return detail
    except requests.RequestException:
        return {}


def format_message(board: dict, row: dict, detail: dict, is_bump: bool = False) -> str:
    tag = "♻️ 재등록(끌올)" if is_bump else "🚨"
    lines = [
        f"{tag} {board['label']}",
        "",
        f"[{row['company']}] {row['title']}",
    ]
    if is_bump:
        lines.append("(이전에 이미 지원 처리했던 글이 등록일자를 바꿔 다시 올라왔습니다)")
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
    if detail.get("attachments"):
        lines.append("")
        lines.append("⚠️ 자체 양식 첨부됨 - 자동지원 보류, 직접 확인 후 지원해주세요")
        for att in detail["attachments"]:
            if att["url"]:
                lines.append(f"- {att['name']} : {att['url']}")
            else:
                lines.append(f"- {att['name']}")
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


def send_application_email(row: dict, detail: dict) -> dict | None:
    """detail에 이메일이 파싱되어 있으면 지원메일을 발송한다. TEST_MODE일 때는
    실제 회사가 아니라 본인 메일로만 보낸다. 실제(TEST_MODE 아닌) 발송이
    성공하면 나중에 답장 확인용으로 쓸 기록(dict)을 반환하고, 그 외에는
    None을 반환한다."""
    if not GMAIL_EMAIL or not GMAIL_APP_PASSWORD:
        print("[WARN] GMAIL_EMAIL / GMAIL_APP_PASSWORD 가 설정되지 않아 이메일 발송을 건너뜁니다.")
        return None

    recipient = detail.get("email")
    if not recipient:
        print(f"[WARN] #{row['no']} {row['title']} - 담당 이메일을 찾지 못해 자동 지원메일을 보내지 않았습니다.")
        return None

    if not RESUME_PATH.exists():
        print(f"[WARN] 이력서 파일({RESUME_PATH.name})을 찾을 수 없어 이메일 발송을 건너뜁니다.")
        return None

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

    if RESUME_PDF_PATH.exists():
        with open(RESUME_PDF_PATH, "rb") as f:
            pdf_part = MIMEApplication(f.read(), Name=RESUME_PDF_PATH.name, _subtype="pdf")
        pdf_part["Content-Disposition"] = f'attachment; filename="{RESUME_PDF_PATH.name}"'
        msg.attach(pdf_part)
    else:
        print("[WARN] 입사지원서.pdf 파일이 없어 PDF는 첨부하지 못했습니다 (docx만 첨부됨).")

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=15) as server:
            server.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_EMAIL, actual_recipient, msg.as_string())
        mode_note = "TEST_MODE" if TEST_MODE else "실제발송"
        print(f"[INFO] 지원메일 발송({mode_note}): #{row['no']} {row['title']} -> {actual_recipient}")
        if not TEST_MODE:
            return {
                "email": recipient.lower(),
                "company": row["company"],
                "title": row["title"],
                "sent_at": datetime.now(timezone.utc).isoformat(),
            }
    except smtplib.SMTPException as e:
        print(f"[ERROR] 지원메일 발송 실패: {e}", file=sys.stderr)
    return None


def process_board(board: dict, state: dict) -> None:
    seen_ids = set(state["seen_ids"])
    seen_post_dates = state.setdefault("seen_post_dates", {})
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
        for row in rows:
            seen_ids.add(row["_fp"])
            seen_post_dates[row["_fp"]] = row.get("posted_at", "")
        state["seen_ids"] = sorted(seen_ids)[-1000:]
        state["bootstrapped_boards"].append(board["key"])
        print(f"[INFO] [{board['key']}] 최초 감시 시작 - 현재 글 {len(rows)}건을 조용히 등록만 했습니다 (알림 생략).")
        return

    # 판단 기준: (1) fp가 seen_ids에 아예 없으면 신규, (2) fp는 이미 있지만
    # 저장해둔 등록일자와 지금 등록일자가 다르면 재등록(끌올). 목록에 실린
    # 글 개수나 번호(no) 크기는 이 판단에 전혀 쓰지 않는다.
    new_rows = []
    bumped_rows = []
    for row in rows:
        fp = row["_fp"]
        cur_date = row.get("posted_at", "")
        if fp not in seen_ids:
            new_rows.append(row)
            continue
        prev_date = seen_post_dates.get(fp)
        if prev_date and cur_date and cur_date != prev_date:
            bumped_rows.append(row)

    to_process = new_rows + bumped_rows
    to_process.sort(key=lambda r: r["no"])

    if not to_process:
        print(
            f"[INFO] [{board['key']}] 신규/재등록 공고 없음 "
            f"(판단 기준: 글 ID + 등록일자 변경 / 참고: 누적 관리 중인 글 {len(seen_ids)}건)"
        )
    else:
        bumped_fps = {r["_fp"] for r in bumped_rows}
        for row in to_process:
            is_bump = row["_fp"] in bumped_fps
            detail = fetch_detail(board["detail_url"], row["id"]) if row.get("id") else {}
            message = format_message(board, row, detail, is_bump=is_bump)
            send_telegram(message)
            tag = "재등록" if is_bump else "신규"
            print(f"[INFO] [{board['key']}] {tag} 알림 전송: #{row['no']} {row['title']}")
            if detail.get("attachments"):
                print(f"[INFO] [{board['key']}] #{row['no']} {row['title']} - 첨부파일(자체 양식) 감지, 자동 지원메일 건너뜀.")
            else:
                record = send_application_email(row, detail)
                if record:
                    state.setdefault("sent_applications", []).append(record)

    # 이번 크롤링에서 보인 모든 글에 대해 seen_ids/등록일자 기준값을 최신화한다.
    # (신규/재등록 여부와 무관하게 항상 최신화해서, 다음 실행에서 정확히
    # 비교할 수 있게 한다. 이 부분도 개수가 아니라 fp 단위로 개별 갱신한다.)
    for row in rows:
        seen_ids.add(row["_fp"])
        seen_post_dates[row["_fp"]] = row.get("posted_at", "")

    state["seen_ids"] = sorted(seen_ids)[-1000:]


def decode_mime_words(raw: str) -> str:
    """메일 제목 등에 쓰이는 MIME 인코딩(=?UTF-8?B?...?=)을 사람이 읽을 수
    있는 문자열로 풀어준다."""
    if not raw:
        return ""
    parts = decode_header(raw)
    out = []
    for text, enc in parts:
        if isinstance(text, bytes):
            out.append(text.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(text)
    return "".join(out)


def check_email_replies(state: dict) -> None:
    """Gmail 받은편지함을 확인해서, 우리가 실제로 지원메일을 보냈던 회사
    주소로부터 온 메일이 있으면 텔레그램으로 알려준다. 같은 메일에 대해
    중복 알림이 가지 않도록 Message-ID 기준으로 기록해둔다."""
    if not GMAIL_EMAIL or not GMAIL_APP_PASSWORD:
        return

    sent_apps = state.get("sent_applications", [])
    if not sent_apps:
        return

    sender_to_company = {}
    for app in sent_apps:
        sender_to_company.setdefault(app["email"].lower(), app["company"])

    notified = set(state.get("notified_reply_ids", []))

    try:
        imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=20)
        imap.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
        imap.select("INBOX")

        since_date = (datetime.now(timezone.utc) - timedelta(days=REPLY_CHECK_LOOKBACK_DAYS)).strftime("%d-%b-%Y")
        status, data = imap.search(None, f"(SINCE {since_date})")
        if status != "OK" or not data or not data[0]:
            imap.logout()
            return

        msg_ids = data[0].split()
        for msg_id in msg_ids:
            status, msg_data = imap.fetch(msg_id, "(BODY.PEEK[HEADER])")
            if status != "OK" or not msg_data or msg_data[0] is None:
                continue
            raw_header = msg_data[0][1]
            msg = message_from_bytes(raw_header)

            _, from_addr = parseaddr(msg.get("From", ""))
            from_addr_l = from_addr.lower()
            if from_addr_l not in sender_to_company:
                continue

            message_id = (msg.get("Message-ID") or "").strip() or f"uid:{msg_id.decode()}"
            if message_id in notified:
                continue

            subject = decode_mime_words(msg.get("Subject", "(제목 없음)"))
            company = sender_to_company[from_addr_l]

            text = (
                "📩 지원메일에 답장이 왔습니다!\n\n"
                f"회사: {company}\n"
                f"보낸사람: {from_addr}\n"
                f"제목: {subject}\n\n"
                "Gmail에서 내용을 확인해주세요."
            )
            send_telegram(text)
            print(f"[INFO] 답장 알림 전송: {company} <{from_addr}>")
            notified.add(message_id)

        imap.logout()
    except (imaplib.IMAP4.error, OSError) as e:
        print(f"[WARN] 메일 답장 확인 중 오류(다음 실행에 재시도): {e}", file=sys.stderr)

    state["notified_reply_ids"] = sorted(notified)[-1000:]


def main() -> None:
    state = load_state()
    for board in BOARDS:
        process_board(board, state)

    if "sent_applications" in state:
        state["sent_applications"] = state["sent_applications"][-500:]

    check_email_replies(state)

    save_state(state)


if __name__ == "__main__":
    main()
