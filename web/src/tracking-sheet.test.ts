import assert from "node:assert/strict";
import { test } from "node:test";

import ExcelJS from "exceljs";

import { DEFAULT_SETTINGS, TRACKING_FIELD_LABELS } from "./config.ts";
import { buildTrackingWorkbook, listAllRows, listEntries, readRow, upsertRow } from "./tracking-sheet.ts";
import type { TrackingRow } from "./tracking-sheet.ts";
import { toArrayBuffer } from "./xlsx-buffer.ts";

const MAPPING = DEFAULT_SETTINGS.trackingSheetMapping;

async function buildTrackingSheet(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(MAPPING.sheetName);
  for (const [field, col] of Object.entries(MAPPING.columns)) {
    ws.getCell(`${col}1`).value = field;
    ws.getCell(`${col}1`).font = { bold: true };
  }
  ws.getCell(`${MAPPING.columns.versionNo}2`).value = 1;
  ws.getCell(`${MAPPING.columns.updateSummary}2`).value = "首次提供文件";
  ws.getCell(`${MAPPING.columns.sourceFileName}2`).value = "G6-需求文件0707.xlsx";
  for (const col of Object.values(MAPPING.columns)) {
    ws.getCell(`${col}2`).font = { italic: true };
  }
  return toArrayBuffer(await wb.xlsx.writeBuffer());
}

async function readWorkbook(bytes: ArrayBuffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  return wb.getWorksheet(MAPPING.sheetName)!;
}

test("upsert appends a new row and copies style", async () => {
  const original = await buildTrackingSheet();
  const updated = await upsertRow(ExcelJS, original, MAPPING, 2, {
    versionNo: 2,
    updateSummary: "第二次更新",
    bilingualFilePath: "v02_20260707/bilingual.xlsx",
  });

  const ws = await readWorkbook(updated);
  assert.equal(ws.rowCount, 3);
  assert.equal(ws.getCell("A3").value, 2);
  assert.equal(ws.getCell("G3").value, "第二次更新");
  assert.equal(ws.getCell("A3").font?.italic, true); // style copied from row 2
  assert.equal(ws.getCell("E3").value, "v02_20260707/bilingual.xlsx");
});

test("upsert for the same version updates the row instead of duplicating", async () => {
  const original = await buildTrackingSheet();
  const afterFirst = await upsertRow(ExcelJS, original, MAPPING, 2, { versionNo: 2, status: "待確認" });
  const afterSecond = await upsertRow(ExcelJS, afterFirst, MAPPING, 2, { status: "已完成" });

  const ws = await readWorkbook(afterSecond);
  assert.equal(ws.rowCount, 3); // still just header + v1 + v2, no duplicate v2 row
  assert.equal(ws.getCell("O3").value, "已完成");
});

test("sequential appends each copy style from their own preceding row", async () => {
  const original = await buildTrackingSheet();
  const afterV2 = await upsertRow(ExcelJS, original, MAPPING, 2, { versionNo: 2 });
  const afterV3 = await upsertRow(ExcelJS, afterV2, MAPPING, 3, { versionNo: 3 });

  const ws = await readWorkbook(afterV3);
  assert.equal(ws.rowCount, 4);
  assert.equal(ws.getCell("A3").font?.italic, true);
  assert.equal(ws.getCell("A4").font?.italic, true);
});

test("upsert keeps two different documents' 'version 1' rows separate when sourceFileName is given", async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet(MAPPING.sheetName);
  for (const [field, col] of Object.entries(MAPPING.columns)) {
    wb.getWorksheet(MAPPING.sheetName)!.getCell(`${col}1`).value = field;
  }
  const empty = toArrayBuffer(await wb.xlsx.writeBuffer());

  // A batch of two unrelated documents both resolve independently to "version 1"
  // against the same (empty) tracking sheet — without matching on sourceFileName too,
  // the second upsert would silently overwrite the first document's row.
  const afterFirst = await upsertRow(ExcelJS, empty, MAPPING, 1, { versionNo: 1, sourceFileName: "spec.xlsx" });
  const afterSecond = await upsertRow(ExcelJS, afterFirst, MAPPING, 1, { versionNo: 1, sourceFileName: "notes.docx" });

  const rows = await listAllRows(ExcelJS, afterSecond, MAPPING);
  assert.deepEqual(
    rows.map((r) => r.sourceFileName),
    ["spec.xlsx", "notes.docx"]
  );
});

test("upsert for the same document+version updates in place instead of duplicating", async () => {
  const original = await buildTrackingSheet(); // version 1, sourceFileName G6-需求文件0707.xlsx
  const updated = await upsertRow(ExcelJS, original, MAPPING, 1, {
    versionNo: 1,
    sourceFileName: "G6-需求文件0707.xlsx",
    status: "已完成",
  });

  const rows = await listAllRows(ExcelJS, updated, MAPPING);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].status, "已完成");
});

test("read row returns null when version missing", async () => {
  const original = await buildTrackingSheet();
  const row = await readRow(ExcelJS, original, MAPPING, 99);
  assert.equal(row, null);
});

test("read row returns existing values", async () => {
  const original = await buildTrackingSheet();
  const row = await readRow(ExcelJS, original, MAPPING, 1);
  assert.equal(row?.updateSummary, "首次提供文件");
});

test("listEntries returns version + source filename for every row", async () => {
  const original = await buildTrackingSheet();
  const withV2 = await upsertRow(ExcelJS, original, MAPPING, 2, {
    versionNo: 2,
    sourceFileName: "G6-需求文件0714.xlsx",
  });

  const entries = await listEntries(ExcelJS, withV2, MAPPING);
  assert.deepEqual(entries, [
    { versionNo: 1, sourceFileName: "G6-需求文件0707.xlsx" },
    { versionNo: 2, sourceFileName: "G6-需求文件0714.xlsx" },
  ]);
});

test("listEntries returns an empty list when the tracking sheet has no data rows yet", async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(MAPPING.sheetName);
  for (const col of Object.values(MAPPING.columns)) ws.getCell(`${col}1`).value = "header";
  const bytes = toArrayBuffer(await wb.xlsx.writeBuffer());

  const entries = await listEntries(ExcelJS, bytes, MAPPING);
  assert.deepEqual(entries, []);
});

test("listAllRows reads every column of every data row for the tracking-management tab", async () => {
  const original = await buildTrackingSheet();
  const withV2 = await upsertRow(ExcelJS, original, MAPPING, 2, {
    versionNo: 2,
    sourceFileName: "G6-需求文件0714.xlsx",
    status: "已完成",
  });

  const rows = await listAllRows(ExcelJS, withV2, MAPPING);
  assert.equal(rows.length, 2);
  assert.equal(rows[0].versionNo, "1");
  assert.equal(rows[0].updateSummary, "首次提供文件");
  assert.equal(rows[1].versionNo, "2");
  assert.equal(rows[1].sourceFileName, "G6-需求文件0714.xlsx");
  assert.equal(rows[1].status, "已完成");
});

test("listAllRows returns an empty list for a tracking sheet with no data rows", async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(MAPPING.sheetName);
  for (const col of Object.values(MAPPING.columns)) ws.getCell(`${col}1`).value = "header";
  const bytes = toArrayBuffer(await wb.xlsx.writeBuffer());

  assert.deepEqual(await listAllRows(ExcelJS, bytes, MAPPING), []);
});

test("buildTrackingWorkbook creates a brand-new sheet with header labels when there's no existing file", async () => {
  const row: TrackingRow = { ...blankRow(), versionNo: "1", sourceFileName: "需求文件.docx", status: "待確認" };
  const bytes = await buildTrackingWorkbook(ExcelJS, MAPPING, [row]);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const ws = wb.getWorksheet(MAPPING.sheetName)!;
  assert.equal(ws.getCell(`${MAPPING.columns.versionNo}1`).value, TRACKING_FIELD_LABELS.versionNo);
  assert.equal(ws.getCell(`${MAPPING.columns.versionNo}2`).value, "1");
  assert.equal(ws.getCell(`${MAPPING.columns.sourceFileName}2`).value, "需求文件.docx");
  assert.equal(ws.getCell(`${MAPPING.columns.status}2`).value, "待確認");
});

test("buildTrackingWorkbook rewrites an existing sheet's data rows wholesale (on-screen edits are what gets saved)", async () => {
  const original = await buildTrackingSheet(); // one data row: version 1
  const editedRows: TrackingRow[] = [
    { ...blankRow(), versionNo: "1", sourceFileName: "改過的檔名.xlsx", status: "已完成" },
    { ...blankRow(), versionNo: "2", sourceFileName: "第二版.xlsx", status: "待確認" },
  ];

  const bytes = await buildTrackingWorkbook(ExcelJS, MAPPING, editedRows, original);
  const rows = await listAllRows(ExcelJS, bytes, MAPPING);
  assert.equal(rows.length, 2);
  assert.equal(rows[0].sourceFileName, "改過的檔名.xlsx");
  assert.equal(rows[0].status, "已完成");
  assert.equal(rows[1].sourceFileName, "第二版.xlsx");
});

test("buildTrackingWorkbook clears stale trailing rows when the edited table has fewer rows than before", async () => {
  const original = await buildTrackingSheet();
  const withV2 = await upsertRow(ExcelJS, original, MAPPING, 2, { versionNo: 2, sourceFileName: "v2.xlsx" });

  const editedRows: TrackingRow[] = [{ ...blankRow(), versionNo: "1", sourceFileName: "only-one-left.xlsx" }];
  const bytes = await buildTrackingWorkbook(ExcelJS, MAPPING, editedRows, withV2);

  const rows = await listAllRows(ExcelJS, bytes, MAPPING);
  assert.deepEqual(
    rows.map((r) => r.sourceFileName),
    ["only-one-left.xlsx"]
  );
});

function blankRow(): TrackingRow {
  const row = {} as TrackingRow;
  for (const field of Object.keys(TRACKING_FIELD_LABELS) as Array<keyof typeof TRACKING_FIELD_LABELS>) row[field] = "";
  return row;
}
