import assert from "node:assert/strict";
import { test } from "node:test";

import ExcelJS from "exceljs";

import { buildDiffReportWorkbook } from "./diff-report.ts";
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
