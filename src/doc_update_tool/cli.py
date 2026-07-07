from __future__ import annotations

from pathlib import Path

import click

from doc_update_tool.archive.manager import ArchiveManager
from doc_update_tool.config import load_client_config
from doc_update_tool.file_types import excel_adapter, word_adapter
from doc_update_tool.pipeline import first_version, update_version
from doc_update_tool.pipeline.apply import apply_translations as run_apply_translations
from doc_update_tool.tracking.tracking_sheet import read_row, upsert_row


def _load(config_path: str):
    config = load_client_config(Path(config_path))
    archive = ArchiveManager(config.folders.incoming, config.folders.archive, config.folders.output)
    return config, archive


@click.group()
def main():
    """客戶需求文件更新處理工具"""


@main.command()
@click.option("--config", "config_path", required=True, help="客戶設定檔路徑 (yaml)")
def process(config_path: str):
    """比對 incoming/ 裡的新文件並產生待翻譯清單"""
    config, archive = _load(config_path)
    if archive.is_first_version():
        worklist_path = first_version.process_first_version(config, archive)
        click.echo(f"首次文件，無前版可比對，已產生全文待翻譯清單：{worklist_path}")
    else:
        worklist_path = update_version.process_update_version(config, archive)
        click.echo(f"已產生差異報告與待翻譯清單：{worklist_path}")
    click.echo("請填完英文翻譯後執行：docmgr apply-translations --config <設定檔>")


@main.command(name="apply-translations")
@click.option("--config", "config_path", required=True)
@click.option("--worklist", "worklist_path", required=False, type=click.Path(exists=True))
def apply_translations_cmd(config_path: str, worklist_path: str | None):
    """讀回填好的待翻譯清單，合併回文件並歸檔、更新追蹤表"""
    config, archive = _load(config_path)
    version_dir = run_apply_translations(
        config, archive, Path(worklist_path) if worklist_path else None
    )
    click.echo(f"已歸檔至：{version_dir}")
    click.echo("請把歸檔內容上傳到 Teams/SharePoint。")


@main.command(name="inspect-excel")
@click.argument("file_path", type=click.Path(exists=True))
@click.option("--sheet", default="auto")
@click.option("--header-row", default=1, type=int)
def inspect_excel_cmd(file_path: str, sheet: str, header_row: int):
    """列出 Excel 各欄位的表頭與範例內容，協助設定 zh_columns/en_columns"""
    for col in excel_adapter.inspect(Path(file_path), sheet_name=sheet, header_row=header_row):
        click.echo(col)


@main.command(name="inspect-word")
@click.argument("file_path", type=click.Path(exists=True))
def inspect_word_cmd(file_path: str):
    """列出 Word 各段落的中英文判斷，協助設定 word_mapping"""
    for para in word_adapter.inspect(Path(file_path)):
        click.echo(para)


@main.command()
@click.option("--config", "config_path", required=True)
@click.option("--version", "version_no", required=True, type=int)
@click.option("--party", required=True, type=click.Choice(["overseas", "customer"]))
@click.option("--date", "confirm_date", required=True)
@click.option("--method", required=True)
def confirm(config_path: str, version_no: int, party: str, confirm_date: str, method: str):
    """登記海外團隊/客戶的確認狀態，雙方皆確認後自動將狀態標記為已完成"""
    config, _archive = _load(config_path)
    current = read_row(config.tracking_sheet, version_no) or {}
    updates = {
        f"{party}_confirm_status": "已確認",
        f"{party}_confirm_date": confirm_date,
        f"{party}_confirm_method": method,
    }
    other = "customer" if party == "overseas" else "overseas"
    if current.get(f"{other}_confirm_status") == "已確認":
        updates["status"] = "已完成"
    upsert_row(config.tracking_sheet, version_no, updates)
    click.echo(f"已更新版本 {version_no} 的 {party} 確認狀態。")


if __name__ == "__main__":
    main()
