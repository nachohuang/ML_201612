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

import { TRACKING_FIELD_LABELS } from "./config.ts";
import type { TrackingColumns, TrackingSheetMapping } from "./config.ts";
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

  let targetRow = findRowByVersion(ws, mapping, versionNo, rowData.sourceFileName as string | undefined);

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

  const row = findRowByVersion(ws, mapping, versionNo);
  if (row === null) return null;

  const result: Record<string, unknown> = {};
  for (const [field, col] of Object.entries(mapping.columns)) {
    result[field] = ws.getCell(`${col}${row}`).value;
  }
  return result;
}

export interface TrackingEntry {
  versionNo: number;
  sourceFileName: string;
}

/** Lists (versionNo, sourceFileName) for every existing row, used to match a newly
 * uploaded file against tracking history by filename (see plan Context: customer
 * filenames often change slightly between versions, e.g. a date suffix, so this is a
 * hint for the user to confirm/pick from rather than something to silently trust). */
export async function listEntries(
  ExcelJS: typeof ExcelJSNS,
  bytes: ArrayBuffer,
  mapping: TrackingSheetMapping
): Promise<TrackingEntry[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const ws = wb.getWorksheet(mapping.sheetName);
  if (!ws) return [];

  const entries: TrackingEntry[] = [];
  for (let row = mapping.headerRow + 1; row <= ws.rowCount; row++) {
    const versionValue = ws.getCell(`${mapping.columns.versionNo}${row}`).value;
    if (versionValue == null || String(versionValue).trim() === "") continue;
    const sourceFileName = String(ws.getCell(`${mapping.columns.sourceFileName}${row}`).value ?? "");
    entries.push({ versionNo: Number(versionValue), sourceFileName });
  }
  return entries;
}

/** Matches on versionNo alone by default, but if `sourceFileName` is given, also
 * requires that column to match — a shared tracking sheet can hold several unrelated
 * documents' histories side by side (e.g. an Excel spec and a Word spec attached to the
 * same batch of files), and each one's version numbering starts independently at 1, so
 * versionNo alone isn't a safe row key once more than one document family is involved:
 * without this, upserting the second document's "version 1" would silently overwrite
 * the first document's "version 1" row instead of adding its own. */
function findRowByVersion(
  ws: ExcelJSNS.Worksheet,
  mapping: TrackingSheetMapping,
  versionNo: number,
  sourceFileName?: string
): number | null {
  for (let row = mapping.headerRow + 1; row <= ws.rowCount; row++) {
    const value = ws.getCell(`${mapping.columns.versionNo}${row}`).value;
    if (value == null || String(value).trim() !== String(versionNo)) continue;
    if (sourceFileName !== undefined) {
      const rowFileName = String(ws.getCell(`${mapping.columns.sourceFileName}${row}`).value ?? "");
      if (rowFileName !== sourceFileName) continue;
    }
    return row;
  }
  return null;
}

export type TrackingRow = Record<keyof TrackingColumns, string>;

const TRACKING_FIELD_KEYS = Object.keys(TRACKING_FIELD_LABELS) as Array<keyof TrackingColumns>;

/** Reads every data row's every column (unlike listEntries, which only reads
 * versionNo/sourceFileName for filename-matching) — powers the tracking-management
 * tab's editable table, where the user can see and change any field on screen. */
export async function listAllRows(
  ExcelJS: typeof ExcelJSNS,
  bytes: ArrayBuffer,
  mapping: TrackingSheetMapping
): Promise<TrackingRow[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const ws = wb.getWorksheet(mapping.sheetName);
  if (!ws) return [];

  const rows: TrackingRow[] = [];
  for (let row = mapping.headerRow + 1; row <= ws.rowCount; row++) {
    const versionValue = ws.getCell(`${mapping.columns.versionNo}${row}`).value;
    const rowValues = TRACKING_FIELD_KEYS.map(
      (field) => ws.getCell(`${mapping.columns[field]}${row}`).value
    );
    const rowIsBlank = versionValue == null && rowValues.every((v) => v == null || String(v).trim() === "");
    if (rowIsBlank) continue;

    const result = {} as TrackingRow;
    for (let i = 0; i < TRACKING_FIELD_KEYS.length; i++) {
      result[TRACKING_FIELD_KEYS[i]] = rowValues[i] == null ? "" : String(rowValues[i]);
    }
    rows.push(result);
  }
  return rows;
}

/** Rebuilds the tracking sheet's data rows from the given in-memory rows (the state of
 * the on-screen editable table) — used both to save edits back onto an uploaded tracking
 * file and to create a brand-new one from scratch when the user has none yet. Existing
 * data rows are cleared and replaced wholesale rather than diffed cell-by-cell: this is
 * a "what you see is what gets saved" table editor, not a merge. */
export async function buildTrackingWorkbook(
  ExcelJS: typeof ExcelJSNS,
  mapping: TrackingSheetMapping,
  rows: TrackingRow[],
  existingBytes?: ArrayBuffer
): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  let ws: ExcelJSNS.Worksheet;
  let templateRow: number | null = null;

  if (existingBytes) {
    await wb.xlsx.load(existingBytes);
    const existing = wb.getWorksheet(mapping.sheetName);
    if (existing) {
      ws = existing;
      if (ws.rowCount > mapping.headerRow) templateRow = mapping.headerRow + 1;
      const clearThrough = Math.max(ws.rowCount, mapping.headerRow + rows.length);
      for (let row = mapping.headerRow + 1; row <= clearThrough; row++) {
        for (const field of TRACKING_FIELD_KEYS) {
          ws.getCell(`${mapping.columns[field]}${row}`).value = null;
        }
      }
    } else {
      ws = wb.addWorksheet(mapping.sheetName);
      writeHeaderRow(ws, mapping);
    }
  } else {
    ws = wb.addWorksheet(mapping.sheetName);
    writeHeaderRow(ws, mapping);
  }

  rows.forEach((rowData, i) => {
    const targetRow = mapping.headerRow + 1 + i;
    if (templateRow !== null && targetRow !== templateRow) copyRowStyle(ws, templateRow, targetRow);
    for (const field of TRACKING_FIELD_KEYS) {
      const value = rowData[field];
      ws.getCell(`${mapping.columns[field]}${targetRow}`).value = value === "" ? null : value;
    }
  });

  return toArrayBuffer(await wb.xlsx.writeBuffer());
}

function writeHeaderRow(ws: ExcelJSNS.Worksheet, mapping: TrackingSheetMapping): void {
  for (const field of TRACKING_FIELD_KEYS) {
    ws.getCell(`${mapping.columns[field]}${mapping.headerRow}`).value = TRACKING_FIELD_LABELS[field];
  }
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
