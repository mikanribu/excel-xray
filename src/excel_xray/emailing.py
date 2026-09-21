"""Optional email delivery for generated Excel reports."""

from __future__ import annotations

import os
import smtplib
from collections.abc import Iterable
from email.message import EmailMessage
from email.utils import getaddresses
from pathlib import Path


DEFAULT_RECIPIENTS = (
    "jacobweglarz@gmail.com",
    "clementine.pages@gmail.com",
)
DEFAULT_RECIPIENT = ", ".join(DEFAULT_RECIPIENTS)
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465
# A 17 MiB source file stays below Gmail's message-size limit after MIME
# base64 encoding and message headers are added.
MAX_ATTACHMENT_BYTES = 17 * 1024 * 1024


class EmailDeliveryError(RuntimeError):
    """A report could not be sent using the configured Gmail account."""


def send_excel_report(
    report_path: str | Path,
    *,
    recipient: str | Iterable[str] | None = None,
    sender: str | None = None,
    app_password: str | None = None,
) -> None:
    """Send one generated Excel report through authenticated Gmail SMTP."""
    path = Path(report_path)
    if path.suffix.lower() != ".xlsx":
        raise EmailDeliveryError("Email delivery supports generated .xlsx reports only")
    if not path.is_file():
        raise EmailDeliveryError(f"Excel report does not exist: {path}")

    try:
        size = path.stat().st_size
    except OSError as exc:
        raise EmailDeliveryError(f"Could not read the Excel report: {exc}") from exc
    if size > MAX_ATTACHMENT_BYTES:
        limit_mb = MAX_ATTACHMENT_BYTES // (1024 * 1024)
        raise EmailDeliveryError(
            f"Excel report is {size / (1024 * 1024):.1f} MiB; Gmail delivery is limited "
            f"to {limit_mb} MiB per attachment. Use the local report or dashboard download."
        )

    sender = sender or os.environ.get("GMAIL_ADDRESS")
    app_password = app_password or os.environ.get("GMAIL_APP_PASSWORD")
    if not sender or not app_password:
        raise EmailDeliveryError(
            "Email delivery needs GMAIL_ADDRESS and GMAIL_APP_PASSWORD environment variables. "
            "Use a Google App Password; do not use your regular account password."
        )

    if recipient is None:
        recipient_values = list(DEFAULT_RECIPIENTS)
    elif isinstance(recipient, str):
        recipient_values = [recipient]
    else:
        recipient_values = list(recipient)
    clean_recipients = tuple(
        dict.fromkeys(address for _, address in getaddresses(recipient_values))
    )
    if not clean_recipients or any(
        not address or "@" not in address or "\n" in address or "\r" in address
        for address in clean_recipients
    ):
        raise EmailDeliveryError("Invalid email recipient list")

    message = EmailMessage()
    message["Subject"] = f"Excel X-ray report: {path.name}"
    message["From"] = sender
    message["To"] = ", ".join(clean_recipients)
    message.set_content(
        "The generated Excel X-ray report is attached.\n\n"
        "Review the workbook before sharing it further; it may contain sensitive "
        "workbook analysis."
    )
    try:
        attachment = path.read_bytes()
    except OSError as exc:
        raise EmailDeliveryError(f"Could not read the Excel report: {exc}") from exc
    message.add_attachment(
        attachment,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=path.name,
    )

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as smtp:
            smtp.login(sender, app_password)
            smtp.send_message(message, to_addrs=list(clean_recipients))
    except (OSError, smtplib.SMTPException) as exc:
        raise EmailDeliveryError(f"Gmail could not send the Excel report: {exc}") from exc
