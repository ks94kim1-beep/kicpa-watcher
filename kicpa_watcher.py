name: KICPA Watcher

on:
  schedule:
    # UTC 기준. 10분마다 실행 (KST = UTC+9)
    - cron: "*/10 * * * *"
  workflow_dispatch: {}  # 수동 실행 버튼도 사용 가능

permissions:
  contents: write  # 이 저장소(kicpa-watcher) 자체에 대한 쓰기 권한 (지금은 커밋할 게 없지만 유지)

jobs:
  check:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout kicpa-watcher (code)
        uses: actions/checkout@v4

      - name: Checkout kicpa-watcher-data (private, 이력서/state.json 보관용)
        uses: actions/checkout@v4
        with:
          repository: ks94kim1-beep/kicpa-watcher-data
          token: ${{ secrets.DATA_REPO_TOKEN }}
          path: data-repo

      - name: Copy resume & state.json into working directory
        run: |
          cp data-repo/입사지원서.docx ./입사지원서.docx
          cp data-repo/입사지원서.pdf ./입사지원서.pdf
          cp data-repo/state.json ./state.json

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Run watcher
        env:
          TELEGRAM_BOT_TOKEN: ${{ secrets.TELEGRAM_BOT_TOKEN }}
          TELEGRAM_CHAT_ID: ${{ secrets.TELEGRAM_CHAT_ID }}
          GMAIL_EMAIL: ${{ secrets.GMAIL_EMAIL }}
          GMAIL_APP_PASSWORD: ${{ secrets.GMAIL_APP_PASSWORD }}
          APPLICANT_NAME: ${{ secrets.APPLICANT_NAME }}
          APPLICANT_BIRTH_DATE: ${{ secrets.APPLICANT_BIRTH_DATE }}
          APPLICANT_BIRTH_YEAR: ${{ secrets.APPLICANT_BIRTH_YEAR }}
          APPLICANT_PASS_YEAR: ${{ secrets.APPLICANT_PASS_YEAR }}
          APPLICANT_PASS_ROUND: ${{ secrets.APPLICANT_PASS_ROUND }}
          TEST_MODE: ${{ secrets.TEST_MODE }}
        run: python kicpa_watcher.py

      - name: Commit updated state.json back to kicpa-watcher-data
        run: |
          cp ./state.json data-repo/state.json
          cd data-repo
          git config user.name "kicpa-watcher-bot"
          git config user.email "actions@github.com"
          git add state.json
          git diff --quiet --cached || git commit -m "update state.json [skip ci]"
          git push
