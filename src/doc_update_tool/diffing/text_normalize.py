"""Normalize text before comparing zh content across versions, so that harmless
differences (full-width vs half-width punctuation, stray whitespace) don't show up
as false changes in the diff report."""

from __future__ import annotations

import unicodedata


def normalize(text: str | None) -> str:
    if text is None:
        return ""
    text = unicodedata.normalize("NFKC", text)
    return " ".join(text.split())
