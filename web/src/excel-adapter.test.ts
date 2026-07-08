import assert from "node:assert/strict";
import { test } from "node:test";

import ExcelJS from "exceljs";

import type { SheetMapping } from "./config.ts";
import { ExcelAdapter, inspectWorkbook } from "./excel-adapter.ts";
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

function mapping(overrides: Partial<SheetMapping> = {}): Record<string, SheetMapping> {
  const m: SheetMapping = {
    sheetName: "Sheet1",
    mode: "table",
    headerRow: 1,
    keyColumns: ["A"],
    zhColumns: ["C"],
    enColumns: ["D"],
    ...overrides,
  };
  return { [m.sheetName]: m };
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

test("multiple sheets with entirely different layouts are all extracted", async () => {
  const wb = new ExcelJS.Workbook();
  const reqWs = wb.addWorksheet("需求列表");
  reqWs.getCell("A1").value = "編號";
  reqWs.getCell("C1").value = "需求說明";
  reqWs.getCell("D1").value = "English";
  reqWs.getCell("A2").value = 1;
  reqWs.getCell("C2").value = "第一項需求";

  const logWs = wb.addWorksheet("維護紀錄");
  logWs.getCell("A1").value = "日期";
  logWs.getCell("B1").value = "維護人員";
  logWs.getCell("C1").value = "Maintainer (EN)";
  logWs.getCell("A2").value = "2026-07-07";
  logWs.getCell("B2").value = "陳宜筠";

  const bytes = toArrayBuffer(await wb.xlsx.writeBuffer());

  const sheetMappings: Record<string, SheetMapping> = {
    需求列表: { sheetName: "需求列表", mode: "table", headerRow: 1, keyColumns: ["A"], zhColumns: ["C"], enColumns: ["D"] },
    維護紀錄: { sheetName: "維護紀錄", mode: "table", headerRow: 1, keyColumns: ["A"], zhColumns: ["B"], enColumns: ["C"] },
  };
  const adapter = new ExcelAdapter(ExcelJS, sheetMappings);
  const segments = adapter.extractSegments(await adapter.load(bytes));

  assert.equal(segments.length, 2);
  assert.ok(segments.some((s) => s.locationId === "需求列表!C2" && s.zhText === "第一項需求"));
  assert.ok(segments.some((s) => s.locationId === "維護紀錄!B2" && s.zhText === "陳宜筠"));
});

test("a renamed sheet in the reference workbook is still found by tab position", async () => {
  // Real customer files rename tabs between versions (e.g. dropping an English
  // suffix) while keeping the same left-to-right order — extractSegments must not
  // silently lose that whole sheet's diff just because the name changed.
  const newWb = new ExcelJS.Workbook();
  newWb.addWorksheet("其他頁籤"); // position 0, irrelevant filler tab
  const newTargetWs = newWb.addWorksheet("需求列表"); // position 1 — new, shorter name
  newTargetWs.getCell("A1").value = "編號";
  newTargetWs.getCell("C1").value = "需求說明";
  newTargetWs.getCell("D1").value = "English";
  newTargetWs.getCell("A2").value = 1;
  newTargetWs.getCell("C2").value = "第一項需求（已修改）";
  const newBytes = toArrayBuffer(await newWb.xlsx.writeBuffer());

  const oldWb = new ExcelJS.Workbook();
  oldWb.addWorksheet("其他頁籤Other"); // position 0, also renamed, irrelevant
  const oldTargetWs = oldWb.addWorksheet("需求列表RequirementList"); // position 1 — old, longer name
  oldTargetWs.getCell("A1").value = "編號";
  oldTargetWs.getCell("C1").value = "需求說明";
  oldTargetWs.getCell("D1").value = "English";
  oldTargetWs.getCell("A2").value = 1;
  oldTargetWs.getCell("C2").value = "第一項需求";
  const oldBytes = toArrayBuffer(await oldWb.xlsx.writeBuffer());

  const sheetMappings: Record<string, SheetMapping> = {
    需求列表: { sheetName: "需求列表", mode: "table", headerRow: 1, keyColumns: ["A"], zhColumns: ["C"], enColumns: ["D"] },
  };
  const adapter = new ExcelAdapter(ExcelJS, sheetMappings);

  const newWbLoaded = await adapter.load(newBytes);
  const oldWbLoaded = await adapter.load(oldBytes);

  // Without a reference workbook, the renamed old sheet can't be found at all.
  assert.equal(adapter.extractSegments(oldWbLoaded).length, 0);

  // With the new workbook as a position reference, it resolves via tab order instead.
  const oldSegments = adapter.extractSegments(oldWbLoaded, newWbLoaded);
  assert.equal(oldSegments.length, 1);
  assert.equal(oldSegments[0].zhText, "第一項需求");
  assert.equal(oldSegments[0].locationId, "需求列表RequirementList!C2"); // uses the sheet's real (old) name
});

test("rich-text cells (mixed fonts within one cell) are read as their concatenated text", async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.getCell("A1").value = "編號";
  ws.getCell("C1").value = "需求說明";
  ws.getCell("D1").value = "English";
  ws.getCell("A2").value = 1;
  ws.getCell("C2").value = {
    richText: [{ text: "第一段" }, { font: { bold: true }, text: "→第二段" }],
  };
  const bytes = toArrayBuffer(await wb.xlsx.writeBuffer());

  const adapter = new ExcelAdapter(ExcelJS, mapping());
  const segments = adapter.extractSegments(await adapter.load(bytes));
  assert.equal(segments[0].zhText, "第一段→第二段");
});

test("a configured sheet that doesn't exist in this workbook is skipped, not an error", async () => {
  const bytes = await buildWorkbook();
  const sheetMappings: Record<string, SheetMapping> = {
    ...mapping(),
    不存在的頁籤: { sheetName: "不存在的頁籤", mode: "table", headerRow: 1, keyColumns: ["A"], zhColumns: ["C"], enColumns: ["D"] },
  };
  const adapter = new ExcelAdapter(ExcelJS, sheetMappings);
  const segments = adapter.extractSegments(await adapter.load(bytes));
  assert.equal(segments.length, 2); // only Sheet1's segments, no error thrown
});

test("inspectWorkbook flags formula columns across all sheets", async () => {
  const wb = new ExcelJS.Workbook();
  const ws1 = wb.addWorksheet("Sheet1");
  ws1.getCell("A1").value = "計算欄";
  ws1.getCell("A2").value = { formula: "1+1" };
  const ws2 = wb.addWorksheet("Sheet2");
  ws2.getCell("A1").value = "純文字欄";
  ws2.getCell("A2").value = "hello";
  const bytes = toArrayBuffer(await wb.xlsx.writeBuffer());

  const sheets = await inspectWorkbook(ExcelJS, bytes);
  assert.equal(sheets.length, 2);
  const sheet1 = sheets.find((s) => s.sheetName === "Sheet1")!;
  const calcCol = sheet1.columns.find((c) => c.header === "計算欄")!;
  assert.equal(calcCol.hasFormula, true);
  const sheet2 = sheets.find((s) => s.sheetName === "Sheet2")!;
  assert.equal(sheet2.columns.find((c) => c.header === "純文字欄")!.hasFormula, false);
});
