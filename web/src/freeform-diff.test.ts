import assert from "node:assert/strict";
import { test } from "node:test";

import ExcelJS from "exceljs";

import { diffFreeformSheet } from "./freeform-diff.ts";
import { ChangeType } from "./models.ts";

function buildSheet(wb: ExcelJS.Workbook, name: string, cells: Record<string, string>): ExcelJS.Worksheet {
  const ws = wb.addWorksheet(name);
  for (const [addr, value] of Object.entries(cells)) ws.getCell(addr).value = value;
  return ws;
}

test("detects modified, added, and deleted cells across a free-form layout", () => {
  const wb = new ExcelJS.Workbook();
  const oldWs = buildSheet(wb, "old", {
    A1: "匯出匯款作業設定",
    B1: "查詢條件",
    A5: "將被刪除的內容",
  });
  const newWs = buildSheet(wb, "new", {
    A1: "匯出匯款作業設定（修改）",
    B1: "查詢條件",
    C3: "新增的內容",
  });

  const records = diffFreeformSheet(oldWs, newWs);
  const byAddress = Object.fromEntries(records.map((r) => [r.locationId.split("!")[1], r]));

  assert.equal(byAddress["A1"].changeType, ChangeType.MODIFIED);
  assert.equal(byAddress["A1"].oldZh, "匯出匯款作業設定");
  assert.equal(byAddress["A1"].newZh, "匯出匯款作業設定（修改）");
  assert.equal(byAddress["A1"].newEnLocationId, undefined); // no translation slot for freeform cells

  assert.equal(byAddress["C3"].changeType, ChangeType.ADDED);
  assert.equal(byAddress["A5"].changeType, ChangeType.DELETED);
  assert.equal(byAddress["B1"], undefined); // unchanged cell isn't reported at all
});

test("a brand-new sheet with no old counterpart reports every cell as added", () => {
  const wb = new ExcelJS.Workbook();
  const newWs = buildSheet(wb, "new", { A1: "全新頁籤", B2: "內容" });

  const records = diffFreeformSheet(undefined, newWs);
  assert.equal(records.length, 2);
  assert.ok(records.every((r) => r.changeType === ChangeType.ADDED));
});
