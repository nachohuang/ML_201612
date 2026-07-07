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
