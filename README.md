# KICPA 구인게시판 감시 → 텔레그램 알림

한국공인회계사회(KICPA) 구인(수습CPA) 게시판(`jobOffrSrchNewGnrl`)에 새 글이
올라오면 10분 이내에 텔레그램으로 알려주는 봇입니다. GitHub Actions로 돌아가서
개인 서버나 항상 켜진 컴퓨터가 필요 없습니다.

robots.txt 확인 결과 `www.kicpa.or.kr`는 `Allow: /` 로 크롤링 자체는
허용되어 있습니다. 다만 이용약관에 별도의 자동수집 제한이 있을 수 있으니
한 번 확인해보시길 권해드립니다. 이 스크립트는 10분에 목록 페이지 1회
요청(하루 144회)만 하도록 만들어 서버 부하를 최소화했습니다.

---

## 1단계 — 텔레그램 봇 만들기

1. 텔레그램에서 **@BotFather** 검색 후 대화 시작
2. `/newbot` 입력
3. 봇 이름, 봇 아이디(반드시 `bot`으로 끝나야 함) 입력
4. 완료되면 **토큰**(`123456789:AAxxxxxxxxxxxxxxxxxxxxxxxxxx` 형태)이 나옵니다.
   이걸 `TELEGRAM_BOT_TOKEN`으로 사용합니다.
5. 방금 만든 봇과 아무 메시지나 한 번 주고받으세요 (예: `/start`).
   봇은 먼저 말을 걸 수 없어서, 사용자가 먼저 대화를 시작해야 메시지를
   받을 수 있습니다.

## 2단계 — 내 chat_id 알아내기

1. 브라우저에서 아래 주소로 접속 (토큰 부분 교체):
   `https://api.telegram.org/bot<위에서 받은 토큰>/getUpdates`
2. 방금 봇에게 보낸 메시지가 JSON으로 보일 텐데, 그 안에서
   `"chat":{"id":123456789, ...}` 부분의 숫자가 `TELEGRAM_CHAT_ID`입니다.

## 3단계 — GitHub 저장소 만들기

1. GitHub에서 새 저장소 생성 (Private 추천 — 어차피 코드에 개인정보는 없지만
   습관적으로 private 권장)
2. 아래 파일 구조를 그대로 업로드:

```
kicpa-watcher/
├── kicpa_watcher.py
├── requirements.txt
├── state.json
└── .github/
    └── workflows/
        └── watch.yml
```

## 4단계 — 저장소 Secrets 등록

저장소 → Settings → Secrets and variables → Actions → New repository secret

- `TELEGRAM_BOT_TOKEN` : 1단계에서 받은 토큰
- `TELEGRAM_CHAT_ID` : 2단계에서 확인한 숫자

## 5단계 — 첫 실행 전 반드시 확인할 것

`state.json`의 `last_seen_no`가 **현재 게시판에 있는 가장 큰 번호**로
맞춰져 있어야 합니다. 그렇지 않으면 첫 실행 때 기존 글 전체가 "신규"로
인식되어 한꺼번에 알림이 옵니다.

- 현재(2026.08.16 확인 기준) 최신 번호는 **13**이라, `state.json`은
  `{"last_seen_no": 13}`으로 이미 맞춰 두었습니다.
- 며칠 뒤에 처음 배포한다면, 배포 직전에 게시판(
  https://www.kicpa.or.kr/home/jobOffrSrchNewGnrl/list.face )에 들어가서
  가장 위 글의 "번호"를 확인하고 그 값으로 `state.json`을 수정하세요.

## 6단계 — 수동으로 한 번 테스트 실행

저장소 → Actions 탭 → "KICPA Watcher" 워크플로우 → **Run workflow** 버튼으로
수동 실행해보세요. 로그에서:

- `[INFO] 신규 공고 없음` 이 뜨면 정상 (state.json의 번호가 최신이라는 뜻)
- `[INFO] 알림 전송: #14 ...` 처럼 뜨면 텔레그램으로 실제 메시지가 갔는지 확인
- `[WARN] 게시글 파싱 결과가 비어 있습니다` 가 뜨면 사이트 구조가 예상과
  달라 표를 못 찾은 것 — 이 경우 알려주시면 파싱 로직을 다시 맞춰드릴게요.

이후로는 `.github/workflows/watch.yml`에 설정된 대로 **10분마다 자동 실행**되고,
새 글을 확인할 때마다 `state.json`을 자동으로 커밋해서 다음 실행에 이어집니다.

---

## 알아두면 좋은 점

- **상세페이지 링크(ijIdNum)**: 목록 페이지의 링크가 자바스크립트로 동작하는
  것으로 보여서, 정규식으로 ID를 최대한 추측해서 찾습니다(`kicpa_watcher.py`
  상단 주석 참고). 실제 배포 후 로그에서 `id` 값이 이상하게 잡히면
  `ID_PATTERN_STRICT` / `ID_PATTERN_LOOSE` 부분을 조정해야 할 수 있습니다.
- **이메일/마감일 추출**: 상세페이지 텍스트에서 정규식으로 뽑기 때문에
  100% 정확하지 않을 수 있어요. 최종 지원 전에는 항상 원문 링크를 열어
  직접 확인하세요.
- **1개월 자동삭제**: 게시판 공지에 "등록 1개월 경과 글은 자동 삭제"라고
  되어 있어서, 번호(no)가 항상 순증가만 하지 않을 수도 있습니다(삭제된
  글의 번호가 재사용되는 경우는 못 봤지만, 혹시 이상 동작이 보이면 알려주세요).
