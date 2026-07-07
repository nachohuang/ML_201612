"""v2+: the customer edited the bilingual file we sent them (old English + new Chinese
mixed in one file), so we diff the Chinese side against the archived previous version
and only need fresh translations for what changed."""

from __future__ import annotations

import shutil
from pathlib import Path

from doc_update_tool.archive.manager import ArchiveManager
from doc_update_tool.archive.msg_reader import read_email_meta
from doc_update_tool.config import ClientConfig
from doc_update_tool.diffing.align import align_and_diff
from doc_update_tool.diffing.diff_report import write_diff_report
from doc_update_tool.models import ChangeType, TranslationItem
from doc_update_tool.pipeline.common import PendingState, get_adapter, save_state, working_dir
from doc_update_tool.translate.manual_file_translator import ManualFileTranslator


def process_update_version(config: ClientConfig, archive: ArchiveManager) -> Path:
    bundle = archive.discover_incoming()
    adapter = get_adapter(config)

    new_segments = adapter.extract_segments(adapter.load(bundle.document_path))

    old_doc_path = archive.latest_bilingual_document()
    if old_doc_path is None:
        raise RuntimeError("archive 內找不到前一版雙語檔，無法比對差異")
    old_segments = adapter.extract_segments(adapter.load(old_doc_path))

    records = align_and_diff(old_segments, new_segments)

    wdir = working_dir(config)
    diff_report_path = wdir / "diff_report.xlsx"
    write_diff_report(records, diff_report_path)

    items = [
        TranslationItem(change_id=r.new_en_location_id, zh_text=r.new_zh or "", context=r.field_name)
        for r in records
        if r.change_type in (ChangeType.ADDED, ChangeType.MODIFIED) and r.new_en_location_id
    ]

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
            is_first_version=False,
            document_filename=bundle.document_path.name,
            sender=email_meta.sender if email_meta else None,
            received_date=email_meta.received_date if email_meta else None,
            email_filename=bundle.email_path.name if bundle.email_path else None,
            diff_report_present=True,
        ),
    )

    return worklist_path
