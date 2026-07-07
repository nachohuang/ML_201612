"""Thin wrapper around extract-msg, used to prefill tracking sheet fields
(sender/date) from the archived .msg without requiring Outlook/win32com."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import extract_msg


@dataclass
class EmailMeta:
    subject: str | None
    sender: str | None
    received_date: str | None


def read_email_meta(path: Path) -> EmailMeta:
    msg = extract_msg.Message(str(path))
    try:
        return EmailMeta(
            subject=msg.subject,
            sender=msg.sender,
            received_date=str(msg.date) if msg.date else None,
        )
    finally:
        msg.close()
