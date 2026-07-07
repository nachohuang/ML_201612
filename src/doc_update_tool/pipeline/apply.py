"""Second half of the pipeline: merge the filled-in worklist back into the working
document, archive the result, and record it in the tracking sheet."""

from __future__ import annotations

from pathlib import Path

from doc_update_tool.archive.manager import ArchiveManager
from doc_update_tool.config import ClientConfig
from doc_update_tool.models import SegmentUpdate
from doc_update_tool.pipeline.common import clear_state, get_adapter, load_state, working_dir
from doc_update_tool.translate.manual_file_translator import ManualFileTranslator
from doc_update_tool.tracking.tracking_sheet import upsert_row


def apply_translations(
    config: ClientConfig, archive: ArchiveManager, worklist_path: Path | None = None
) -> Path:
    state = load_state(config)
    wdir = working_dir(config)

    worklist_path = worklist_path or (wdir / "待翻譯清單.xlsx")
    results = ManualFileTranslator().read_results(worklist_path)

    work_doc_path = wdir / f"new_document{Path(state.document_filename).suffix.lower()}"
    adapter = get_adapter(config)
    doc = adapter.load(work_doc_path)

    updates = [SegmentUpdate(en_location_id=r.change_id, en_text=r.en_text) for r in results]
    adapter.apply_translations(doc, updates)
    adapter.save(doc, work_doc_path)

    diff_report_path = wdir / "diff_report.xlsx"
    diff_report_path = diff_report_path if diff_report_path.exists() else None

    email_path = wdir / "email.msg"
    email_path = email_path if email_path.exists() else None

    version_dir = archive.commit_version(
        version_no=state.version_no,
        document_path=work_doc_path,
        diff_report_path=diff_report_path,
        email_path=email_path,
    )

    bilingual_path = version_dir / f"bilingual{work_doc_path.suffix.lower()}"
    diff_report_archived = version_dir / "diff_report.xlsx" if diff_report_path else None
    email_archived = version_dir / "email.msg" if email_path else None

    upsert_row(
        config.tracking_sheet,
        version_no=state.version_no,
        row_data={
            "version_no": state.version_no,
            "received_date": state.received_date or "",
            "sender": state.sender or "",
            "email_path": str(email_archived) if email_archived else "",
            "bilingual_file_path": str(bilingual_path),
            "diff_report_path": str(diff_report_archived) if diff_report_archived else "",
            "update_summary": "首次提供文件" if state.is_first_version else "",
            "is_first_version": "是" if state.is_first_version else "否",
            "overseas_confirm_status": "未確認",
            "customer_confirm_status": "未確認",
            "status": "待確認",
        },
    )

    clear_state(config)
    return version_dir
