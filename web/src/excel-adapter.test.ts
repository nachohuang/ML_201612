import assert from "node:assert/strict";
import { test } from "node:test";

import ExcelJS from "exceljs";

import type { ExcelMapping } from "./config.ts";
import { ExcelAdapter, inspectExcel } from "./excel-adapter.ts";
import type { SegmentUpdate } from "./models.ts";
import { toArrayBuffer } from "./xlsx-buffer.ts";

interface BuildOpts {
  zhCol?: string;
  enCol?: string;
  headerRow?: number;
  addFormula?: boolean;
  mergeRange?: string;
}

async function buildWorkbook(opts: BuildOpts = {}): Promise<ArrayBuffer> {
  const { zhCol = "C", enCol = "D", headerRow = 1, addFormula = false, mergeRange } = opts;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.getCell(`A${headerRow}`).value = "編號";
  ws.getCell(`${zhCol}${headerRow}`).value = "需求說明";
  ws.getCell(`${enCol}${headerRow}`).value = "English";
  ws.getCell(`A${headerRow}`).font = { bold: true };

  const row1 = headerRow + 1;
  ws.getCell(`A${row1}`).value = 1;
  ws.getCell(`${zhCol}${row1}`).value = "第一項需求";
  ws.getCell(`${enCol}${row1}`).value = "";

  const row2 = row1 + 1;
  ws.getCell(`A${row2}`).value = 2;
  ws.getCell(`${zhCol}${row2}`).value = "第二項需求";
  ws.getCell(`${enCol}${row2}`).value = "";

  if (addFormula) ws.getCell("Z1").value = { formula: "1+1" };
  if (mergeRange) ws.mergeCells(mergeRange);

  return toArrayBuffer(await wb.xlsx.writeBuffer());
}

function mapping(overrides: Partial<ExcelMapping> = {}): ExcelMapping {
  return { sheetName: "auto", headerRow: 1, keyColumns: ["A"], zhColumns: ["C"], enColumns: ["D"], ...overrides };
}

test("extract segments — adjacent columns", async () => {
  const bytes = await buildWorkbook({ zhCol: "C", enCol: "D" });
  const adapter = new ExcelAdapter(ExcelJS, mapping());
  const segments = adapter.extractSegments(await adapter.load(bytes));

  assert.equal(segments.length, 2);
  assert.equal(segments[0].zhText, "第一項需求");
  assert.equal(segments[0].rowKey, "1");
  assert.equal(segments[0].fieldName, "需求說明");
  assert.ok(segments[0].locationId.endsWith("!C2"));
  assert.ok(segments[0].enLocationId.endsWith("!D2"));
});

test("extract segments — non-adjacent columns are config-driven, not hardcoded", async () => {
  const bytes = await buildWorkbook({ zhCol: "B", enCol: "F" });
  const adapter = new ExcelAdapter(ExcelJS, mapping({ zhColumns: ["B"], enColumns: ["F"] }));
  const segments = adapter.extractSegments(await adapter.load(bytes));

  assert.equal(segments.length, 2);
  assert.equal(segments[0].zhText, "第一項需求");
  assert.ok(segments[0].enLocationId.endsWith("!F2"));
});

test("extract segments — header not on first row", async () => {
  const bytes = await buildWorkbook({ headerRow: 3 });
  const adapter = new ExcelAdapter(ExcelJS, mapping({ headerRow: 3 }));
  const segments = adapter.extractSegments(await adapter.load(bytes));

  assert.equal(segments.length, 2);
  assert.equal(segments[0].fieldName, "需求說明");
  assert.ok(segments[0].locationId.endsWith("!C4"));
});

test("apply translations round trip preserves the rest of the file", async () => {
  const bytes = await buildWorkbook({ addFormula: true, mergeRange: "F1:G1" });
  const adapter = new ExcelAdapter(ExcelJS, mapping());

  const wb = await adapter.load(bytes);
  const segments = adapter.extractSegments(wb);
  const updates: SegmentUpdate[] = [
    { enLocationId: segments[0].enLocationId, enText: "Item one" },
    { enLocationId: segments[1].enLocationId, enText: "Item two" },
  ];
  adapter.applyTranslations(wb, updates);
  const outBytes = await adapter.save(wb);

  const reloaded = new ExcelJS.Workbook();
  await reloaded.xlsx.load(outBytes);
  const ws = reloaded.worksheets[0];

  assert.equal(ws.getCell("D2").value, "Item one");
  assert.equal(ws.getCell("D3").value, "Item two");
  assert.equal(ws.getCell("C2").value, "第一項需求"); // zh column untouched
  assert.equal((ws.getCell("Z1").value as { formula: string }).formula, "1+1"); // unrelated formula untouched
  assert.equal(ws.getCell("A1").font?.bold, true); // header style preserved
  assert.ok(ws.getCell("F1").isMerged); // merged cells preserved
});

test("extract segments skips fully blank rows", async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.getCell("A1").value = "編號";
  ws.getCell("C1").value = "需求說明";
  ws.getCell("D1").value = "English";
  ws.getCell("A2").value = 1;
  ws.getCell("C2").value = "內容";
  // rows 3-5 intentionally blank
  const bytes = toArrayBuffer(await wb.xlsx.writeBuffer());

  const adapter = new ExcelAdapter(ExcelJS, mapping());
  const segments = adapter.extractSegments(await adapter.load(bytes));
  assert.equal(segments.length, 1);
});

test("inspect flags formula columns", async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.getCell("A1").value = "計算欄";
  ws.getCell("A2").value = { formula: "1+1" };
  const bytes = toArrayBuffer(await wb.xlsx.writeBuffer());

  const cols = await inspectExcel(ExcelJS, bytes);
  const calcCol = cols.find((c) => c.header === "計算欄")!;
  assert.equal(calcCol.hasFormula, true);
});
