# claude-slack-report

DB 조회 결과를 Slack 채널에 **요약 글 + 사용한 쿼리 + 그래프(PNG) + 엑셀**로 보내는 Claude Code 스킬입니다.
"○○ 확인해서 슬랙으로 보내 줘"라고 하면 Claude가 조회 → 미리 보기 → 승인 → 보내기 순서로 처리합니다.

- **조회 전용**: `SELECT`·`SHOW`·`DESCRIBE`·`EXPLAIN` 한 문장만 실행하고, 추가·수정·삭제 쿼리는 실행하지 않습니다. 보내기 도구도 조회가 아닌 쿼리면 멈춥니다. 다만 진짜 보장은 **읽기 전용 DB 계정**이니 꼭 함께 쓰세요.
- **보내기 전 승인**: 보낼 글·쿼리·그래프를 먼저 보여 주고, 승인해야 보냅니다.

## 준비

1. **조회 도구(MCP)**: Claude Code에 DB를 조회할 수 있는 MCP가 연결돼 있어야 합니다. 예를 들어 메타베이스 0.60 이상은 MCP가 들어 있습니다(관리자 설정 → AI → MCP에서 켜기).
   ```
   claude mcp add --scope user --transport http metabase https://<메타베이스 주소>/api/metabase-mcp
   ```
   Claude Code는 `https`가 아닌 주소로는 로그인하지 않습니다(`localhost`만 예외).
2. **Slack 앱**: api.slack.com/apps에서 앱을 만들고, Bot Token Scopes에 `chat:write`, `files:write`를 넣어 워크스페이스에 설치합니다. 결과를 받을 채널에 봇을 초대합니다(`/invite @봇이름`).
3. **Python 3.8 이상**: 처음 실행할 때 `~/.cache/slack-report/venv`에 matplotlib·openpyxl을 알아서 설치합니다. 따로 설치 명령은 없습니다.

## 설치

Claude Code에서:
```
/plugin marketplace add hava12/claude-slack-report
/plugin install slack-report@claude-slack-report
```

그다음 터미널에서 설정 파일을 만들고, `SLACK_BOT_TOKEN`(`xoxb-…`)과 기본 `SLACK_CHANNEL`(채널 ID)을 채웁니다.
```
mkdir -p ~/.config/slack-report
cp ~/.claude/plugins/marketplaces/claude-slack-report/skills/slack-report/.env.template ~/.config/slack-report/.env
chmod 600 ~/.config/slack-report/.env
open -e ~/.config/slack-report/.env
```
설정이 없는 채로 보내려고 하면, 스킬이 이 명령을 알려 줍니다. 환경 변수 `SLACK_BOT_TOKEN`·`SLACK_CHANNEL`을 써도 됩니다(파일보다 먼저).

업데이트는 `/plugin marketplace update claude-slack-report`입니다.

## 쓰기

```
판매중인 상품이 10개 이상인 센터를 조회해서 그래프는 상위 10곳, 엑셀은 전체로 슬랙에 보내 줘
```

## 확인

```
~/.cache/slack-report/venv/bin/python skills/slack-report/test_slack_report.py
```
(Slack에는 아무것도 보내지 않습니다. `run.sh`를 한 번 실행해 가상 환경이 생긴 뒤에 돌리세요.)

## 라이선스

MIT
