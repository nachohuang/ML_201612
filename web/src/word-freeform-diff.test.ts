import assert from "node:assert/strict";
import { test } from "node:test";

import { DOMParser } from "@xmldom/xmldom";

import { ChangeType } from "./models.ts";
import { diffFreeformWordDocument } from "./word-freeform-diff.ts";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function paragraph(text: string): string {
  return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

function parseBody(bodyXml: string) {
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W_NS}"><w:body>${bodyXml}</w:body></w:document>`;
  return new DOMParser().parseFromString(xml, "application/xml");
}

test("detects a modified paragraph", () => {
  const oldDom = parseBody(paragraph("計算Calculate") + paragraph("不變的段落"));
  const newDom = parseBody(paragraph("計算Calculate（修改）") + paragraph("不變的段落"));

  const records = diffFreeformWordDocument(oldDom, newDom);
  assert.equal(records.length, 1);
  assert.equal(records[0].changeType, ChangeType.MODIFIED);
  assert.equal(records[0].newZh, "計算Calculate（修改）");
  assert.equal(records[0].newEnLocationId, undefined); // no translation slot for freeform content
});

test("an inserted paragraph in the middle doesn't get misread as every later paragraph changing", () => {
  // This is exactly the real-world failure mode: a naive index-i-vs-index-i compare
  // would flag every paragraph after the insertion point as "modified" once shifted.
  const oldDom = parseBody(paragraph("段落A") + paragraph("段落B") + paragraph("段落C"));
  const newDom = parseBody(paragraph("段落A") + paragraph("新插入的段落") + paragraph("段落B") + paragraph("段落C"));

  const records = diffFreeformWordDocument(oldDom, newDom);
  assert.equal(records.length, 1, "only the genuinely new paragraph should be reported");
  assert.equal(records[0].changeType, ChangeType.ADDED);
  assert.equal(records[0].newZh, "新插入的段落");
});

test("detects a deleted paragraph without disturbing paragraphs around it", () => {
  const oldDom = parseBody(paragraph("段落A") + paragraph("將被刪除") + paragraph("段落B"));
  const newDom = parseBody(paragraph("段落A") + paragraph("段落B"));

  const records = diffFreeformWordDocument(oldDom, newDom);
  assert.equal(records.length, 1);
  assert.equal(records[0].changeType, ChangeType.DELETED);
  assert.equal(records[0].oldZh, "將被刪除");
});

test("blank paragraphs are ignored, not reported as changes", () => {
  const oldDom = parseBody(paragraph("內容") + paragraph(""));
  const newDom = parseBody(paragraph("內容") + paragraph("") + paragraph(""));

  const records = diffFreeformWordDocument(oldDom, newDom);
  assert.equal(records.length, 0);
});

test("detects table cell changes, aligning rows by content so an inserted row doesn't shift the rest", () => {
  const row = (a: string, b: string) =>
    `<w:tr><w:tc><w:p><w:r><w:t>${a}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>${b}</w:t></w:r></w:p></w:tc></w:tr>`;

  const oldDom = parseBody(`<w:tbl>${row("計算Calculate", "修改Modify")}${row("刪除Delete", "查詢List")}</w:tbl>`);
  const newDom = parseBody(
    `<w:tbl>${row("計算Calculate", "修改Modify")}${row("新增Create", "新欄位New")}${row("刪除Delete", "查詢List")}</w:tbl>`
  );

  const records = diffFreeformWordDocument(oldDom, newDom);
  assert.equal(records.length, 2, "only the two cells in the newly inserted row");
  assert.ok(records.every((r) => r.changeType === ChangeType.ADDED));
  assert.ok(records.some((r) => r.newZh === "新增Create"));
  assert.ok(records.some((r) => r.newZh === "新欄位New"));
});

test("a brand new document (v1) with no old counterpart reports everything as added", () => {
  const newDom = parseBody(paragraph("全新段落"));
  const records = diffFreeformWordDocument(undefined, newDom);
  assert.equal(records.length, 1);
  assert.equal(records[0].changeType, ChangeType.ADDED);
});
