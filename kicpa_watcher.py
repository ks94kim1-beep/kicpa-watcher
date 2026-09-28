name: KICPA Daily Report

on:
  schedule:
    # UTC 15:00 = 한국시간 자정(00:00). 예약 실행은 몇 분~수십 분 늦을 수 있음
    - cron: "0 15 * * *"
  workflow_dispatch:
    inputs:
      report_date:
        description: "리포트 날짜 (YYYY-MM-DD, 비우면 자동)"
        required: false
        default: ""

permissions:
  contents: read  # state.json을 읽기만 하고 커밋하지 않음

jobs:
  report:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout kicpa-watcher (code)
        uses: actions/checkout@v4

      - name: Checkout kicpa-watcher-data (private, state.json 보관용)
        uses: actions/checkout@v4
        with:
          repository: ks94kim1-beep/kicpa-watcher-data
          token: ${{ secrets.DATA_REPO_TOKEN }}
          path: data-repo

      - name: Copy state.json into working directory
        run: cp data-repo/state.json ./state.json

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Run daily report
        env:
          TELEGRAM_BOT_TOKEN: ${{ secrets.TELEGRAM_BOT_TOKEN }}
          TELEGRAM_CHAT_ID: ${{ secrets.TELEGRAM_CHAT_ID }}
          GMAIL_EMAIL: ${{ secrets.GMAIL_EMAIL }}
          GMAIL_APP_PASSWORD: ${{ secrets.GMAIL_APP_PASSWORD }}
          REPORT_DATE: ${{ github.event.inputs.report_date }}
        run: python daily_report.py
