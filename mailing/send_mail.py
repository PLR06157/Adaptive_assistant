"""
Utility for sending personalized HTML emails with attachments via Microsoft 365.

The script reads recipient data from a CSV file, renders an HTML template using
row values, and delivers the messages through the Microsoft Graph API.
"""

from __future__ import annotations

import argparse
import base64
import csv
import json
import logging
import mimetypes
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional
from premailer import transform

import requests
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"


class ConfigurationError(RuntimeError):
    """Raised when required configuration is missing."""


def _read_env(name: str, *, required: bool = True, default: Optional[str] = None) -> str:
    value = os.getenv(name, default)
    if required and not value:
        raise ConfigurationError(
            f"Missing required configuration for {name}. "
            "Set it in your environment or .env file."
        )
    return value or ""


def _load_html_template(path: Path) -> str:
    try:
        with open(path, 'r') as f:
            html = f.read()

        inline_html = transform(html)

        with open(path, 'w') as f:
            f.write(inline_html)
        
        return path.read_text(encoding="utf-8")
    except FileNotFoundError as exc:
        raise ConfigurationError(f"HTML template not found: {path}") from exc


def _guess_mime_type(path: Path) -> str:
    mime, _ = mimetypes.guess_type(path.name)
    return mime or "application/octet-stream"


def _build_attachment(path: Path) -> Dict[str, str]:
    if not path.exists():
        raise ConfigurationError(f"Attachment file does not exist: {path}")
    file_bytes = path.read_bytes()
    content_bytes = base64.b64encode(file_bytes).decode("ascii")
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": path.name,
        "contentType": _guess_mime_type(path),
        "contentBytes": content_bytes,
    }


@dataclass
class Recipient:
    email: str
    subject: str
    context: Dict[str, str]


class GraphMailer:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str, sender: str) -> None:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        self._sender = sender
        self._client = ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority,
        )

    def _get_token(self) -> str:
        token = self._client.acquire_token_silent(GRAPH_SCOPE, account=None)
        if not token:
            token = self._client.acquire_token_for_client(scopes=GRAPH_SCOPE)
        if "access_token" not in token:
            raise RuntimeError(f"Unable to acquire access token: {json.dumps(token, indent=2)}")
        return token["access_token"]

    def send(
        self,
        recipients: Iterable[Recipient],
        html_template: str,
        *,
        attachment: Optional[Dict[str, str]] = None,
        dry_run: bool = False,
    ) -> None:
        token = None if dry_run else self._get_token()
        total = 0
        for recipient in recipients:
            total += 1
            rendered_html = html_template.format(**recipient.context)
            if dry_run:
                logging.info(
                    "[DRY-RUN] Would send to %s with subject '%s'",
                    recipient.email,
                    recipient.subject,
                )
                continue
            payload = {
                "message": {
                    "subject": recipient.subject,
                    "body": {"contentType": "HTML", "content": rendered_html},
                    "toRecipients": [{"emailAddress": {"address": recipient.email}}],
                },
                "saveToSentItems": True,
            }
            if attachment:
                payload["message"]["attachments"] = [attachment]
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            }
            response = requests.post(
                f"{GRAPH_ENDPOINT}/users/{self._sender}/sendMail",
                headers=headers,
                json=payload,
                timeout=30,
            )
            if response.status_code >= 300:
                raise RuntimeError(
                    f"Failed to send mail to {recipient.email}: "
                    f"{response.status_code} {response.text}"
                )
            logging.info("Sent mail to %s", recipient.email)
        logging.info("Processed %d recipient(s).", total)


def _parse_recipients(
    csv_path: Path,
    *,
    email_column: str,
    subject_column: Optional[str],
    default_subject: Optional[str],
) -> List[Recipient]:
    if not csv_path.exists():
        raise ConfigurationError(f"CSV file not found: {csv_path}")
    recipients: List[Recipient] = []
    with csv_path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        for idx, row in enumerate(reader, start=2):
            email = row.get(email_column, "").strip()
            if not email:
                logging.warning("Row %d missing email in column '%s'; skipping.", idx, email_column)
                continue
            subject = ""
            if subject_column:
                subject = row.get(subject_column, "").strip()
            subject = subject or (default_subject or "").strip()
            if not subject:
                raise ConfigurationError(
                    f"Row {idx} missing subject and no default subject provided."
                )
            recipients.append(
                Recipient(
                    email=email,
                    subject=subject,
                    context={key: value for key, value in row.items()},
                )
            )
    if not recipients:
        raise ConfigurationError("No valid recipients found in CSV.")
    return recipients


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Send personalized HTML emails with optional attachment via Microsoft 365."
    )
    parser.add_argument(
        "--csv",
        dest="csv_path",
        default=os.getenv("CSV_PATH", "mailing/recipients.csv"),
        help="Path to the recipient CSV file (default: %(default)s).",
    )
    parser.add_argument(
        "--template",
        dest="template_path",
        default=os.getenv("HTML_TEMPLATE_PATH", "mailing/email_template.html"),
        help="Path to the HTML template file (default: %(default)s).",
    )
    parser.add_argument(
        "--attachment",
        default=os.getenv("ATTACHMENT_PATH"),
        help="Path to the file attachment (optional).",
    )
    parser.add_argument(
        "--email-column",
        default=os.getenv("EMAIL_COLUMN", "email"),
        help="Name of the CSV column containing recipient email addresses (default: %(default)s).",
    )
    parser.add_argument(
        "--subject-column",
        default=os.getenv("SUBJECT_COLUMN"),
        help="Name of the CSV column containing the subject line.",
    )
    parser.add_argument(
        "--default-subject",
        default=os.getenv("DEFAULT_SUBJECT"),
        help="Subject to use when the subject column is absent or empty.",
    )
    parser.add_argument(
        "--log-level",
        default=os.getenv("LOG_LEVEL", "INFO"),
        help="Logging level (default: %(default)s).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Render emails without calling the Graph API.",
    )
    return parser


def main() -> None:
    load_dotenv()
    parser = build_parser()
    args = parser.parse_args()

    logging.basicConfig(level=args.log_level.upper(), format="%(levelname)s %(message)s")

    try:
        tenant_id = _read_env("TENANT_ID")
        client_id = _read_env("CLIENT_ID")
        client_secret = _read_env("CLIENT_SECRET")
        sender_address = _read_env("SENDER_EMAIL")

        html_template = _load_html_template(Path(args.template_path))

        recipients = _parse_recipients(
            Path(args.csv_path),
            email_column=args.email_column,
            subject_column=args.subject_column,
            default_subject=args.default_subject,
        )

        attachment = None
        if args.attachment:
            attachment = _build_attachment(Path(args.attachment))

        mailer = GraphMailer(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            sender=sender_address,
        )
        mailer.send(
            recipients,
            html_template,
            attachment=attachment,
            dry_run=args.dry_run,
        )
    except ConfigurationError as exc:
        logging.error("%s", exc)
    except Exception as exc:  # pylint: disable=broad-except
        logging.exception("Unexpected error: %s", exc)


if __name__ == "__main__":
    main()
