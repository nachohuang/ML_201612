"""Load and validate per-client YAML configuration.

Real customer document layouts are not yet known, so column/paragraph mapping
is entirely config-driven — nothing about zh/en layout is hardcoded here.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml


class ConfigError(ValueError):
    """Raised when a client config file is missing required fields or malformed."""


@dataclass
class Folders:
    incoming: Path
    archive: Path
    output: Path


@dataclass
class TrackingSheetConfig:
    path: Path
    sheet_name: str
    header_row: int
    columns: dict[str, str]  # logical field name -> column letter


@dataclass
class ExcelMapping:
    sheet_name: str = "auto"
    header_row: int = 1
    key_columns: list[str] = field(default_factory=list)
    zh_columns: list[str] = field(default_factory=list)
    en_columns: list[str] = field(default_factory=list)
    on_delete: str = "mark"  # "mark" | "remove"


@dataclass
class WordMapping:
    mode: str = "alternating_paragraphs"  # alternating_paragraphs | same_paragraph_split | table_based
    split_delimiter: str | None = None
    zh_columns: list[str] = field(default_factory=list)  # used when mode == table_based
    en_columns: list[str] = field(default_factory=list)


@dataclass
class ClientConfig:
    client_id: str
    folders: Folders
    tracking_sheet: TrackingSheetConfig
    document_format: str  # "excel" | "word"
    excel_mapping: ExcelMapping
    word_mapping: WordMapping


_REQUIRED_TRACKING_FIELDS = [
    "version_no",
    "received_date",
    "sender",
    "email_path",
    "bilingual_file_path",
    "diff_report_path",
    "update_summary",
    "is_first_version",
    "overseas_confirm_status",
    "overseas_confirm_date",
    "overseas_confirm_method",
    "customer_confirm_status",
    "customer_confirm_date",
    "customer_confirm_method",
    "status",
    "notes",
]


def _require(mapping: dict, key: str, context: str) -> object:
    if key not in mapping:
        raise ConfigError(f"{context}: missing required field '{key}'")
    return mapping[key]


def load_client_config(path: Path) -> ClientConfig:
    path = Path(path)
    if not path.exists():
        raise ConfigError(f"Client config not found: {path}")

    with path.open("r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    client = _require(raw, "client", str(path))
    client_id = _require(client, "id", "client")

    base_dir = path.parent

    folders_raw = _require(client, "folders", "client")
    folders = Folders(
        incoming=_resolve(base_dir, _require(folders_raw, "incoming", "client.folders")),
        archive=_resolve(base_dir, _require(folders_raw, "archive", "client.folders")),
        output=_resolve(base_dir, _require(folders_raw, "output", "client.folders")),
    )

    tracking_raw = _require(client, "tracking_sheet", "client")
    columns = _require(tracking_raw, "columns", "client.tracking_sheet")
    missing = [f for f in _REQUIRED_TRACKING_FIELDS if f not in columns]
    if missing:
        raise ConfigError(
            f"client.tracking_sheet.columns is missing required field(s): {', '.join(missing)}"
        )
    tracking_sheet = TrackingSheetConfig(
        path=_resolve(base_dir, _require(tracking_raw, "path", "client.tracking_sheet")),
        sheet_name=tracking_raw.get("sheet_name", "Sheet1"),
        header_row=int(tracking_raw.get("header_row", 1)),
        columns=dict(columns),
    )

    document_format = client.get("document_format", "excel")
    if document_format not in ("excel", "word"):
        raise ConfigError(
            f"client.document_format must be 'excel' or 'word', got: {document_format!r}"
        )

    excel_raw = client.get("excel_mapping", {}) or {}
    excel_mapping = ExcelMapping(
        sheet_name=excel_raw.get("sheet_name", "auto"),
        header_row=int(excel_raw.get("header_row", 1)),
        key_columns=list(excel_raw.get("key_columns", [])),
        zh_columns=list(excel_raw.get("zh_columns", [])),
        en_columns=list(excel_raw.get("en_columns", [])),
        on_delete=excel_raw.get("on_delete", "mark"),
    )
    if document_format == "excel":
        if len(excel_mapping.zh_columns) != len(excel_mapping.en_columns):
            raise ConfigError(
                "client.excel_mapping.zh_columns and en_columns must have the same length "
                f"(got {len(excel_mapping.zh_columns)} vs {len(excel_mapping.en_columns)})"
            )
        if not excel_mapping.zh_columns:
            raise ConfigError(
                "client.excel_mapping must define at least one zh_columns/en_columns pair "
                "when document_format is 'excel'"
            )

    word_raw = client.get("word_mapping", {}) or {}
    word_mapping = WordMapping(
        mode=word_raw.get("mode", "alternating_paragraphs"),
        split_delimiter=word_raw.get("split_delimiter"),
        zh_columns=list(word_raw.get("zh_columns", [])),
        en_columns=list(word_raw.get("en_columns", [])),
    )
    if document_format == "word" and word_mapping.mode not in (
        "alternating_paragraphs",
        "same_paragraph_split",
        "table_based",
    ):
        raise ConfigError(f"client.word_mapping.mode is invalid: {word_mapping.mode!r}")

    return ClientConfig(
        client_id=client_id,
        folders=folders,
        tracking_sheet=tracking_sheet,
        document_format=document_format,
        excel_mapping=excel_mapping,
        word_mapping=word_mapping,
    )


def _resolve(base_dir: Path, value: str) -> Path:
    p = Path(value)
    return p if p.is_absolute() else (base_dir / p).resolve()
