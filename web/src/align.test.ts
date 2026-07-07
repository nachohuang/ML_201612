import assert from "node:assert/strict";
import { test } from "node:test";

import { alignAndDiff } from "./align.ts";
import { ChangeType } from "./models.ts";
import type { Segment } from "./models.ts";

function seg(rowKey: string | null, zh: string, en = "", field = "需求說明", loc?: string): Segment {
  const locationId = loc ?? `C${rowKey}`;
  return {
    locationId,
    enLocationId: `D${rowKey}`,
    rowKey,
    fieldName: field,
    zhText: zh,
    enText: en,
  };
}

test("modified and added with row key", () => {
  const oldSegs = [seg("1", "A"), seg("2", "B")];
  const newSegs = [seg("1", "A"), seg("2", "B2"), seg("3", "C")];

  const records = alignAndDiff(oldSegs, newSegs);
  const byKey = Object.fromEntries(records.map((r) => [r.locationId, r]));

  assert.equal(byKey["C1"].changeType, ChangeType.UNCHANGED);
  assert.equal(byKey["C2"].changeType, ChangeType.MODIFIED);
  assert.equal(byKey["C2"].oldZh, "B");
  assert.equal(byKey["C2"].newZh, "B2");
  assert.equal(byKey["C2"].newEnLocationId, "D2");
  assert.equal(byKey["C3"].changeType, ChangeType.ADDED);
  assert.equal(byKey["C3"].newZh, "C");
});

test("deleted row", () => {
  const oldSegs = [seg("1", "A"), seg("2", "B"), seg("3", "C")];
  const newSegs = [seg("1", "A"), seg("3", "C")];

  const records = alignAndDiff(oldSegs, newSegs);
  const deletedRecords = records.filter((r) => r.changeType === ChangeType.DELETED);
  assert.equal(deletedRecords.length, 1);
  assert.equal(deletedRecords[0].oldZh, "B");
  assert.equal(deletedRecords[0].newEnLocationId, undefined);
});

test("unchanged carries forward old english", () => {
  const oldSegs = [seg("1", "A", "Translated A")];
  const newSegs = [seg("1", "A")];

  const records = alignAndDiff(oldSegs, newSegs);
  assert.equal(records[0].changeType, ChangeType.UNCHANGED);
  assert.equal(records[0].newEn, "Translated A");
});

test("normalization ignores full-width and whitespace differences", () => {
  const oldSegs = [seg("1", "第一項")];
  const newSegs = [seg("1", "第一項 ")]; // trailing space only

  const records = alignAndDiff(oldSegs, newSegs);
  assert.equal(records[0].changeType, ChangeType.UNCHANGED);
});

test("position fallback without row key", () => {
  const oldSegs = [seg(null, "A", "", "需求說明", "C2"), seg(null, "B", "", "需求說明", "C3")];
  const newSegs = [seg(null, "A2", "", "需求說明", "C2"), seg(null, "B", "", "需求說明", "C3")];

  const records = alignAndDiff(oldSegs, newSegs);
  assert.ok(records.some((r) => r.changeType === ChangeType.MODIFIED));
});
