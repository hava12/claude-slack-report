#!/bin/sh
# 처음 실행할 때 그래프·엑셀 라이브러리를 전용 가상 환경에 설치하고, 그다음부터는 바로 slack_report.py 를 실행한다.
V="${SLACK_REPORT_VENV:-$HOME/.cache/slack-report/venv}"
if [ ! -x "$V/bin/python" ]; then
  python3 -m venv "$V" && "$V/bin/pip" install -q --disable-pip-version-check matplotlib openpyxl || exit 1
fi
exec "$V/bin/python" "$(dirname "$0")/slack_report.py" "$@"
