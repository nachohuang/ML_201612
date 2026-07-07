"""v1: the customer's first document has no previous version to diff against, so the
whole document is translated (no diff report step)."""

from __future__ import annotations

import shutil
from pathlib import Path

from doc_update_tool.archive.manager import ArchiveManager
from doc_update_tool.archive.msg_reader import read_email_meta
from doc_update_tool.config import ClientConfig
from doc_update_tool.models import TranslationItem
from doc_update_tool.pipeline.common import PendingState, get_adapter, save_state, working_dir
from doc_update_tool.translate.manual_file_translator import ManualFileTranslator


def process_first_version(config: ClientConfig, archive: ArchiveManager) -> Path:
    bundle = archive.discover_incoming()
    adapter = get_adapter(config)
    segments = adapter.extract_segments(adapter.load(bundle.document_path))

    items = [
        TranslationItem(change_id=s.en_location_id, zh_text=s.zh_text, context=s.field_name)
        for s in segments
        if s.zh_text.strip()
    ]

    wdir = working_dir(config)
    work_doc_path = wdir / f"new_document{bundle.document_path.suffix.lower()}"
    shutil.copy2(bundle.document_path, work_doc_path)

    worklist_path = wdir / "待翻譯清單.xlsx"
    ManualFileTranslator().write_worklist(items, worklist_path)

    email_meta = read_email_meta(bundle.email_path) if bundle.email_path else None
    if bundle.email_path:
        shutil.copy2(bundle.email_path, wdir / "email.msg")

    save_state(
        config,
        PendingState(
            version_no=archive.next_version_no(),
            is_first_version=True,
            document_filename=bundle.document_path.name,
            sender=email_meta.sender if email_meta else None,
            received_date=email_meta.received_date if email_meta else None,
            email_filename=bundle.email_path.name if bundle.email_path else None,
            diff_report_present=False,
        ),
    )

    return worklist_path
