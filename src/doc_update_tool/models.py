"""Shared data structures used across adapters, diffing, translation and tracking."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Literal


class ChangeType(str, Enum):
    ADDED = "新增"
    MODIFIED = "修改"
    DELETED = "刪除"
    UNCHANGED = "不變"


@dataclass
class Segment:
    """One zh/en pair extracted from a document (an Excel row-field or a Word paragraph pair)."""

    location_id: str
    """Stable position identifier of the zh side (e.g. 'Sheet1!C12' or 'paragraph:4'), used
    when no row_key is configured, and shown in diff reports."""

    en_location_id: str
    """Where the corresponding translation is written back to (e.g. 'Sheet1!D12' or 'paragraph:5')."""

    row_key: str | None
    """Content-based key used to align rows/segments across versions (e.g. a requirement ID column)."""

    field_name: str | None
    """Human-readable column/field name, shown in diff reports. None for Word paragraphs."""

    zh_text: str
    en_text: str


@dataclass
class ChangeRecord:
    location_id: str
    field_name: str | None
    change_type: ChangeType
    old_zh: str | None = None
    new_zh: str | None = None
    old_en: str | None = None
    new_en: str | None = None
    new_en_location_id: str | None = None
    """Where to write a fresh translation in the *new* document. Set for ADDED/MODIFIED
    records; None for DELETED (no corresponding location in the new document)."""


@dataclass
class TranslationItem:
    change_id: str
    zh_text: str
    context: str | None = None


@dataclass
class TranslationResult:
    change_id: str
    en_text: str
    source: Literal["manual", "llm", "carried_forward"] = "manual"


@dataclass
class SegmentUpdate:
    """Instruction to write en_text back into a document at en_location_id."""

    en_location_id: str
    en_text: str


@dataclass
class IncomingBundle:
    document_path: Path
    email_path: Path | None


@dataclass
class VersionInfo:
    version_no: int
    is_first_version: bool
    archive_dir: Path
