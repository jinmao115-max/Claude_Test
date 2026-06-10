"""
Gmail sender using the Gmail API (OAuth2 credentials).
Reads credentials from environment variables GMAIL_SENDER and GMAIL_TOKEN_JSON,
or falls back to credentials.json / token.json in the project root.
"""

import base64
import json
import logging
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

RECIPIENT = os.environ.get("NOTIFY_EMAIL", "jinmao115@gmail.com")
SENDER = os.environ.get("GMAIL_SENDER", "jinmao115@gmail.com")
TOKEN_JSON = os.environ.get("GMAIL_TOKEN_JSON", "")
TOKEN_FILE = "token.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


def _get_credentials() -> Credentials:
    if TOKEN_JSON:
        info = json.loads(TOKEN_JSON)
        return Credentials.from_authorized_user_info(info, SCOPES)
    return Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)


def send_gmail(report: dict) -> None:
    subject = f"【ホテル最安値】{report['date']} の価格レポート"
    html_body = report["html_body"]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = SENDER
    msg["To"] = RECIPIENT
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()

    creds = _get_credentials()
    service = build("gmail", "v1", credentials=creds)
    service.users().messages().send(userId="me", body={"raw": raw}).execute()

    logger.info("Gmail sent to %s (subject: %s)", RECIPIENT, subject)
