/** Mirrors tracking/tracking_sheet.py's safe upsert — but adapted for the browser:
 * there's no real filesystem path to hyperlink to (Path B has no folder access), so
 * path-like fields are written as plain descriptive text (e.g. the name the file will
 * have once the user unzips the output bundle) rather than a clickable hyperlink.
 *
 * There's also no on-disk backup/restore step: the caller (app.ts) keeps the original
 * uploaded bytes until upsertRow resolves successfully, so a mid-way failure here just
 * means the original tracking sheet file is never touched/downloaded.
 */

import type ExcelJSNS from "exceljs";

import type { TrackingSheetMapping } from "./config.ts";
import { toArrayBuffer } from "./xlsx-buffer.ts";

export type TrackingRowData = Partial<Record<keyof TrackingSheetMapping["columns"], string | number>>;

export class TrackingSheetError extends Error {}

export async function upsertRow(
  ExcelJS: typeof ExcelJSNS,
  bytes: ArrayBuffer,
  mapping: TrackingSheetMapping,
  versionNo: number,
  rowData: TrackingRowData
): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const ws = wb.getWorksheet(mapping.sheetName);
  if (!ws) throw new TrackingSheetError(`追蹤表內找不到工作表: ${mapping.sheetName}`);

  const versionCol = mapping.columns.versionNo;
  let targetRow = findRowByVersion(ws, versionCol, mapping.headerRow, versionNo);

  if (targetRow === null) {
    targetRow = ws.rowCount + 1;
    if (ws.rowCount > mapping.headerRow) copyRowStyle(ws, ws.rowCount, targetRow);
  }

  for (const [field, value] of Object.entries(rowData)) {
    const colLetter = mapping.columns[field as keyof TrackingSheetMapping["columns"]];
    if (!colLetter) continue; // unrecognized field name, ignore rather than fail the whole upsert
    ws.getCell(`${colLetter}${targetRow}`).value = value as string | number;
  }

  return toArrayBuffer(await wb.xlsx.writeBuffer());
}

export async function readRow(
  ExcelJS: typeof ExcelJSNS,
  bytes: ArrayBuffer,
  mapping: TrackingSheetMapping,
  versionNo: number
): Promise<Record<string, unknown> | null> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const ws = wb.getWorksheet(mapping.sheetName);
  if (!ws) return null;

  const row = findRowByVersion(ws, mapping.columns.versionNo, mapping.headerRow, versionNo);
  if (row === null) return null;

  const result: Record<string, unknown> = {};
  for (const [field, col] of Object.entries(mapping.columns)) {
    result[field] = ws.getCell(`${col}${row}`).value;
  }
  return result;
}

function findRowByVersion(ws: ExcelJSNS.Worksheet, versionCol: string, headerRow: number, versionNo: number): number | null {
  for (let row = headerRow + 1; row <= ws.rowCount; row++) {
    const value = ws.getCell(`${versionCol}${row}`).value;
    if (value != null && String(value).trim() === String(versionNo)) return row;
  }
  return null;
}

function copyRowStyle(ws: ExcelJSNS.Worksheet, srcRow: number, dstRow: number): void {
  for (let col = 1; col <= ws.columnCount; col++) {
    const srcCell = ws.getRow(srcRow).getCell(col);
    const dstCell = ws.getRow(dstRow).getCell(col);
    // Deep-clone before assigning: after a save/reload round trip, ExcelJS cells with
    // identical style values share the same in-memory style object, and a later plain
    // `cell.font = {...}` mutates that shared object in place rather than detaching
    // first — so any code that mutates an *existing* row's style after creation (not
    // done anywhere today) must clone first or it will silently reach into other rows
    // with the same style.
    dstCell.style = JSON.parse(JSON.stringify(srcCell.style));
  }
}
