import assert from "node:assert/strict";
import { test } from "node:test";

import ExcelJS from "exceljs";

import { buildCombinedReportWorkbook, buildDiffReportWorkbook } from "./diff-report.ts";
import type { ReportRow } from "./diff-report.ts";
import { ChangeType } from "./models.ts";
import type { ChangeRecord } from "./models.ts";

test("diff report lists only added/modified/deleted, not unchanged", async () => {
  const records: ChangeRecord[] = [
    { locationId: "C2", fieldName: "需求說明", changeType: ChangeType.UNCHANGED, oldZh: "A", newZh: "A" },
    { locationId: "C3", fieldName: "需求說明", changeType: ChangeType.MODIFIED, oldZh: "B", newZh: "B2" },
    { locationId: "C4", fieldName: "需求說明", changeType: ChangeType.ADDED, newZh: "C" },
  ];

  const bytes = await buildDiffReportWorkbook(ExcelJS, records);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const ws = wb.worksheets[0];

  assert.equal(ws.rowCount, 3); // header + modified + added, unchanged excluded
  assert.equal(ws.getCell("A2").value, "C3");
  assert.equal(ws.getCell("E2").value, "修改");
  assert.equal(ws.getCell("A3").value, "C4");
  assert.equal(ws.getCell("E3").value, "新增");
});

test("combined report carries the source document name plus note/translation per row, across a batch of documents", async () => {
  const rows: ReportRow[] = [
    {
      docName: "需求文件A.xlsx",
      locationId: "Sheet1!C3",
      fieldName: "需求說明",
      oldZh: "B",
      newZh: "B2",
      changeType: ChangeType.MODIFIED,
      note: "已跟客戶確認過",
      translation: "Updated requirement",
    },
    {
      docName: "需求文件B.docx",
      locationId: "para:4",
      fieldName: null,
      newZh: "新段落",
      changeType: ChangeType.ADDED,
      note: "",
      translation: "New paragraph",
    },
    { docName: "需求文件A.xlsx", locationId: "Sheet1!C2", fieldName: null, changeType: ChangeType.UNCHANGED, oldZh: "A", newZh: "A", note: "", translation: "" },
  ];

  const bytes = await buildCombinedReportWorkbook(ExcelJS, rows);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const ws = wb.worksheets[0];

  assert.equal(ws.rowCount, 3); // header + modified + added, unchanged excluded
  assert.equal(ws.getCell("A2").value, "需求文件A.xlsx");
  assert.equal(ws.getCell("G2").value, "已跟客戶確認過");
  assert.equal(ws.getCell("H2").value, "Updated requirement");
  assert.equal(ws.getCell("A3").value, "需求文件B.docx");
  assert.equal(ws.getCell("F3").value, "新增");
  assert.equal(ws.getCell("H3").value, "New paragraph");
});
