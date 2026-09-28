"""
KICPA Watcher 일일 리포트 (매일 00:00 KST 실행)

1. state.json의 processed_log에서 "리포트 대상 날짜(KST)"에 처리된 공고를 모은다.
2. Gmail 보낸편지함(IMAP)을 읽어 자동 지원메일이 실제로 나갔는지 확인한다.
3. 지금 게시판에 아직 올라와 있는데 지원 흔적이 없는 공고(자체 양식 보류,
   이메일 없음, 발송 실패 등)를 모아 "아직 지원 안 한 공고"로 보여준다.
   자체 양식 건을 같은 Gmail로 직접 지원했다면 보낸편지함에 잡히므로 자동으로
   목록에서 빠진다.
4. 결과를 텔레그램으로 보낸다.

state.json은 읽기만 하고 수정/커밋하지 않는다 (10분 워처와 충돌 방지).
코드 레포가 public이라 Actions 로그에는 회사명/이메일을 찍지 않고 건수만 남긴다.

리포트 대상 날짜: REPORT_DATE(YYYY-MM-DD) 환경변수가 있으면 그 날짜,
없으면 "현재 KST 시각 - 12시간"의 날짜. 자정 실행이 조금 늦어져도 전날이
대상이 되고, 낮에 수동 실행하면 그날 지금까지의 현황이 나온다.
"""

import imaplib
import json
import os
import re
import sys
from datetime import date, datetime, timedelta, timezone
from email import message_from_bytes
from email.utils import getaddresses, parsedate_to_datetime

import requests

from kicpa_watcher import (
    BOARDS,
    GMAIL_APP_PASSWORD,
    GMAIL_EMAIL,
    IMAP_HOST,
    IMAP_PORT,
    STATE_PATH,
    fetch_list_html,
    fingerprint,
    parse_rows,
    send_telegram,
)

KST = timezone(timedelta(hours=9))
SENT_LOOKBACK_DAYS = 45
WEEKDAYS = "월화수목금토일"

STATUS_LABELS = {
    "sent": "자동 지원",
    "test": "TEST_MODE(본인에게 발송)",
    "skipped_form": "자체 양식",
    "skipped_merge": "서류 병합 요구",
    "no_email": "담당 이메일 못 찾음",
    "detail_failed": "상세페이지 조회 실패",
    "send_failed": "메일 발송 실패",
    "no_config": "Gmail 설정 없음",
    "no_resume": "이력서 파일 없음",
}
AUTO_STATUSES = {"sent", "test"}
HELD_STATUSES = {"skipped_form", "skipped_merge"}


def target_date() -> date:
    override = os.environ.get("REPORT_DATE", "").strip()
    if override:
        return datetime.strptime(override, "%Y-%m-%d").date()
    return (datetime.now(KST) - timedelta(hours=12)).date()


def to_kst(iso: str) -> datetime:
    dt = datetime.fromisoformat(iso)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(KST)


def load_sent_mail() -> dict | None:
    """보낸편지함에서 최근 SENT_LOOKBACK_DAYS일 동안 보낸 메일의
    {수신주소(소문자): [발송시각, ...]} 매핑을 만든다. 실패하면 None."""
    if not GMAIL_EMAIL or not GMAIL_APP_PASSWORD:
        return None
    try:
        imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=20)
        imap.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)

        # Gmail 언어 설정에 따라 보낸편지함 폴더 이름이 달라서, 이름 대신
        # \Sent 플래그가 붙은 폴더를 찾는다.
        sent_box = None
        status, folders = imap.list()
        for raw in folders or []:
            line = raw.decode(errors="replace") if isinstance(raw, bytes) else str(raw)
            if "\\Sent" in line:
                m = re.search(r'\) (?:"[^"]*"|NIL) (.+)$', line)
                if m:
                    sent_box = m.group(1).strip()
                    break
        if not sent_box:
            print("[WARN] 보낸편지함 폴더를 찾지 못했습니다.", file=sys.stderr)
            imap.logout()
            return None

        status, _ = imap.select(sent_box, readonly=True)
        if status != "OK":
            print("[WARN] 보낸편지함을 열지 못했습니다.", file=sys.stderr)
            imap.logout()
            return None

        since = (datetime.now(timezone.utc) - timedelta(days=SENT_LOOKBACK_DAYS)).strftime("%d-%b-%Y")
        status, data = imap.search(None, f"(SINCE {since})")
        sent: dict[str, list[datetime]] = {}
        if status == "OK" and data and data[0]:
            for msg_id in data[0].split():
                status, msg_data = imap.fetch(msg_id, "(BODY.PEEK[HEADER.FIELDS (TO CC DATE)])")
                if status != "OK" or not msg_data or msg_data[0] is None:
                    continue
                msg = message_from_bytes(msg_data[0][1])
                try:
                    dt = parsedate_to_datetime(msg.get("Date", ""))
                except (TypeError, ValueError):
                    continue
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=timezone.utc)
                for _, addr in getaddresses(msg.get_all("To", []) + msg.get_all("Cc", [])):
                    if addr:
                        sent.setdefault(addr.lower(), []).append(dt)
        imap.logout()
        return sent
    except (imaplib.IMAP4.error, OSError) as e:
        print(f"[WARN] 보낸편지함 확인 실패: {e}", file=sys.stderr)
        return None


def sent_after(sent_map: dict, addr: str, after_iso: str) -> bool:
    """addr로 after_iso 시각(5분 여유) 이후에 보낸 메일이 있는지."""
    if not addr or addr not in sent_map:
        return False
    after = to_kst(after_iso) - timedelta(minutes=5)
    return any(dt >= after for dt in sent_map[addr])


def fetch_current_fps() -> set | None:
    """지금 게시판에 올라와 있는(감시 대상) 공고들의 식별자 집합."""
    fps = set()
    for board in BOARDS:
        try:
            rows = parse_rows(fetch_list_html(board["list_url"]))
        except requests.RequestException as e:
            print(f"[WARN] 게시판 조회 실패({board['key']}): {e}", file=sys.stderr)
            return None
        keywords = board.get("title_keywords")
        if keywords:
            rows = [r for r in rows if any(kw in r["title"] for kw in keywords)]
        for r in rows:
            fps.add(fingerprint(board, r))
    return fps


def build_report(state: dict, day: date, sent_map: dict | None, current_fps: set | None) -> str:
    log = state.get("processed_log", [])
    todays = [e for e in log if to_kst(e["processed_at"]).date() == day]

    lines = [f"📋 {day:%m/%d}({WEEKDAYS[day.weekday()]}) KICPA 일일 리포트", ""]

    if not todays:
        lines.append("새로 올라온 공고 없음 (감시는 정상 작동 중)")
    else:
        bumps = sum(1 for e in todays if e.get("is_bump"))
        head = f"새 공고 {len(todays)}건"
        if bumps:
            head += f" (재등록 {bumps}건 포함)"
        lines.append(head)

        auto = [e for e in todays if e["status"] in AUTO_STATUSES]
        held = [e for e in todays if e["status"] in HELD_STATUSES]
        failed = [e for e in todays if e["status"] not in AUTO_STATUSES | HELD_STATUSES]

        if auto:
            lines += ["", f"✅ 지원메일 발송 {len(auto)}건"]
            for e in auto:
                mark = ""
                if e["status"] == "test":
                    mark = " (TEST)"
                elif sent_map is not None and not sent_after(sent_map, e["recipient"], e["processed_at"]):
                    mark = " ⚠️ 보낸편지함에서 확인 안 됨"
                lines.append(f"- {e['company']}{mark}")

        if held:
            lines += ["", f"⏭ 자동지원 보류 {len(held)}건"]
            for e in held:
                lines.append(f"- {e['company']}: {STATUS_LABELS[e['status']]}")

        if failed:
            lines += ["", f"⚠️ 발송 안 됨 {len(failed)}건"]
            for e in failed:
                lines.append(f"- {e['company']}: {STATUS_LABELS.get(e['status'], e['status'])}")

    # 아직 게시 중인데 지원 흔적이 없는 공고 (날짜 무관, 누적)
    lines.append("")
    if current_fps is None:
        lines.append("※ 게시판 조회에 실패해서 미지원 공고 점검은 건너뛰었습니다.")
    else:
        latest: dict[str, dict] = {}
        for e in log:
            latest[e["fp"]] = e  # 같은 공고가 여러 번 처리됐으면 마지막 기록 기준
        pending = []
        for fp, e in latest.items():
            if fp not in current_fps or e["status"] in AUTO_STATUSES:
                continue
            if sent_map is not None and sent_after(sent_map, e["recipient"], e["processed_at"]):
                continue  # 직접 지원한 흔적이 보낸편지함에 있음
            pending.append(e)
        if pending:
            lines.append(f"📌 아직 지원 안 한 공고 (게시 중) {len(pending)}건")
            for e in pending:
                note = STATUS_LABELS.get(e["status"], e["status"])
                if not e["recipient"]:
                    note += ", 지원 여부 직접 확인 필요"
                lines.append(f"- {e['company']}: {note}")
        else:
            lines.append("📌 게시 중인 공고 중 미지원 건 없음")

    if sent_map is None:
        lines += ["", "※ 보낸편지함 확인에 실패해서 실제 발송 대조는 건너뛰었습니다."]

    return "\n".join(lines)


def main() -> None:
    state = json.loads(STATE_PATH.read_text(encoding="utf-8")) if STATE_PATH.exists() else {}
    day = target_date()
    sent_map = load_sent_mail()
    current_fps = fetch_current_fps()
    report = build_report(state, day, sent_map, current_fps)
    send_telegram(report)
    print(f"[INFO] {day} 일일 리포트 전송 완료 ({len(report.splitlines())}줄)")


if __name__ == "__main__":
    main()
