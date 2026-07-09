/** Mirrors diffing/diff_report.py — renders a ChangeRecord list as the "新舊對照列表"
 * xlsx, built fresh each time so there's no round-trip fidelity risk here. */

import type ExcelJSNS from "exceljs";

import { ChangeType } from "./models.ts";
import type { ChangeRecord } from "./models.ts";
import { toArrayBuffer } from "./xlsx-buffer.ts";

const HEADERS = ["位置", "欄位名稱", "舊內容", "新內容", "變動類型"];

export async function buildDiffReportWorkbook(
  ExcelJS: typeof ExcelJSNS,
  records: ChangeRecord[]
): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("差異報告");
  ws.addRow(HEADERS);

  for (const record of records) {
    if (record.changeType === ChangeType.UNCHANGED) continue; // internal bookkeeping only
    ws.addRow([record.locationId, record.fieldName ?? "", record.oldZh ?? "", record.newZh ?? "", record.changeType]);
  }

  return toArrayBuffer(await wb.xlsx.writeBuffer());
}

/** One combined report row across a whole batch of documents — unlike ChangeRecord,
 * this also carries which document it came from plus the user's on-screen note and
 * translation, since a batch run produces one merged report rather than one per file. */
export interface ReportRow {
  docName: string;
  locationId: string;
  fieldName: string | null;
  oldZh?: string;
  newZh?: string;
  changeType: ChangeType;
  note: string;
  translation: string;
}

const COMBINED_HEADERS = ["文件名稱", "位置", "欄位名稱", "舊內容", "新內容", "變動類型", "備註", "翻譯"];

export async function buildCombinedReportWorkbook(ExcelJS: typeof ExcelJSNS, rows: ReportRow[]): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("差異報告");
  ws.addRow(COMBINED_HEADERS);

  for (const row of rows) {
    if (row.changeType === ChangeType.UNCHANGED) continue;
    ws.addRow([row.docName, row.locationId, row.fieldName ?? "", row.oldZh ?? "", row.newZh ?? "", row.changeType, row.note, row.translation]);
  }

  return toArrayBuffer(await wb.xlsx.writeBuffer());
}
