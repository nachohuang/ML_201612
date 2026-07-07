"""Render a ChangeRecord list as the "新舊對照列表" (old-vs-new comparison table) xlsx —
the same report format regardless of whether the source was Excel or Word."""

from __future__ import annotations

from pathlib import Path

import openpyxl

from doc_update_tool.models import ChangeRecord, ChangeType

_HEADERS = ["位置", "欄位名稱", "舊內容", "新內容", "變動類型"]


def write_diff_report(records: list[ChangeRecord], path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "差異報告"
    ws.append(_HEADERS)

    for record in records:
        if record.change_type == ChangeType.UNCHANGED:
            continue  # internal bookkeeping only, not shown to the user
        ws.append(
            [
                record.location_id,
                record.field_name or "",
                record.old_zh or "",
                record.new_zh or "",
                record.change_type.value,
            ]
        )

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
