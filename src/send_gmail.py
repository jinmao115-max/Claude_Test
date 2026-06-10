import base64
import json
import logging
import os
import tempfile
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
SENDER = os.environ.get("GMAIL_SENDER", "jinmao115@gmail.com")


def send_email(to_email: str, search_results: list) -> None:
    """
    search_results: [{"condition": dict, "results": list, "cheapest": dict|None}, ...]
    """
    creds   = _get_credentials()
    subject = f"【ホテル最安値】{date.today().isoformat()} の価格レポート"
    html    = _build_html(search_results)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = SENDER
    msg["To"]      = to_email
    msg.attach(MIMEText(html, "html", "utf-8"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service = build("gmail", "v1", credentials=creds)
    service.users().messages().send(userId="me", body={"raw": raw}).execute()

    logger.info("Gmail sent → %s", to_email)


# ── credentials ──────────────────────────────────────────────

def _get_credentials() -> Credentials:
    token_json = os.environ.get("GMAIL_TOKEN")
    creds_json = os.environ.get("GMAIL_CREDENTIALS")

    if token_json:
        info = json.loads(token_json)
        creds = Credentials.from_authorized_user_info(info, SCOPES)
        # refresh_token が含まれていればそのまま使用可能
        return creds

    # ローカル実行: token.json ファイルにフォールバック
    if creds_json:
        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False) as f:
            f.write(creds_json)
            creds_file = f.name
    else:
        creds_file = "credentials.json"

    return Credentials.from_authorized_user_file("token.json", SCOPES)


# ── HTML builder ─────────────────────────────────────────────

def _build_html(search_results: list) -> str:
    today = date.today().isoformat()
    sections = "".join(_section_html(r) for r in search_results)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head><meta charset="UTF-8">
<style>
  body {{ font-family: sans-serif; background:#f5f7fa; padding:20px; color:#1e293b; }}
  h1   {{ font-size:20px; margin-bottom:4px; }}
  .sub {{ color:#64748b; font-size:13px; margin-bottom:24px; }}
  .block {{ background:#fff; border:1px solid #e2e8f0; border-radius:10px;
            padding:16px 20px; margin-bottom:20px; }}
  .block h2 {{ font-size:15px; margin-bottom:4px; color:#2563eb; }}
  .meta {{ font-size:12px; color:#64748b; margin-bottom:12px; }}
  .cheapest {{ background:#f0fdf4; border:1px solid #bbf7d0; border-radius:8px;
               padding:10px 14px; margin-bottom:12px; font-size:14px; }}
  table {{ width:100%; border-collapse:collapse; font-size:13px; }}
  th,td {{ padding:7px 10px; border-bottom:1px solid #e2e8f0; text-align:left; }}
  th    {{ background:#f8fafc; font-weight:600; }}
  a     {{ color:#2563eb; text-decoration:none; }}
</style>
</head>
<body>
<h1>🏨 ホテル最安値レポート</h1>
<p class="sub">{today} 実行</p>
{sections}
</body></html>"""


def _section_html(r: dict) -> str:
    cond     = r.get("condition", {})
    results  = r.get("results", [])
    cheapest = r.get("cheapest")
    name     = cond.get("hotel_name") or cond.get("location", "検索条件")
    checkin  = cond.get("checkin", "")
    checkout = cond.get("checkout", "")
    guests   = cond.get("guests", 2)

    cheapest_html = ""
    if cheapest:
        cheapest_html = (
            f'<div class="cheapest">🏆 最安値: '
            f'<a href="{cheapest["url"]}">{cheapest["hotel_name"]}</a> '
            f'— <strong>¥{cheapest["price"]:,}</strong> ({cheapest["site"]})</div>'
        )

    rows = "".join(
        f'<tr><td>{i}</td>'
        f'<td><a href="{h["url"]}">{h["hotel_name"]}</a></td>'
        f'<td>{h["site"]}</td>'
        f'<td>¥{h["price"]:,}</td></tr>'
        for i, h in enumerate(results[:10], 1)
    ) or "<tr><td colspan='4'>結果なし</td></tr>"

    return f"""
<div class="block">
  <h2>📍 {name}</h2>
  <p class="meta">{checkin} → {checkout} ／ {guests}名</p>
  {cheapest_html}
  <table>
    <tr><th>#</th><th>ホテル名</th><th>サイト</th><th>価格</th></tr>
    {rows}
  </table>
</div>"""
