"""Shared helpers for the process / apply-translations two-step pipeline.

A CLI invocation of `process` and the later `apply-translations` invocation are
separate OS processes (the user edits the worklist file in between), so state has to
be persisted to disk rather than kept in memory. `.working/` under the client's output
folder holds that in-flight state; it's an implementation detail, not something the
user needs to touch.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path

from doc_update_tool.config import ClientConfig
from doc_update_tool.file_types.excel_adapter import ExcelAdapter
from doc_update_tool.file_types.word_adapter import WordAdapter


def get_adapter(config: ClientConfig):
    if config.document_format == "excel":
        return ExcelAdapter(config.excel_mapping)
    return WordAdapter(config.word_mapping)


def working_dir(config: ClientConfig) -> Path:
    d = config.folders.output.parent / ".working"
    d.mkdir(parents=True, exist_ok=True)
    return d


@dataclass
class PendingState:
    version_no: int
    is_first_version: bool
    document_filename: str
    sender: str | None
    received_date: str | None
    email_filename: str | None
    diff_report_present: bool


def save_state(config: ClientConfig, state: PendingState) -> None:
    path = working_dir(config) / "state.json"
    path.write_text(json.dumps(asdict(state), ensure_ascii=False, indent=2), encoding="utf-8")


def load_state(config: ClientConfig) -> PendingState:
    path = working_dir(config) / "state.json"
    if not path.exists():
        raise RuntimeError(
            "找不到待處理狀態，請先執行 `docmgr process`，並在填完待翻譯清單後再執行 apply-translations。"
        )
    data = json.loads(path.read_text(encoding="utf-8"))
    return PendingState(**data)


def clear_state(config: ClientConfig) -> None:
    (working_dir(config) / "state.json").unlink(missing_ok=True)
