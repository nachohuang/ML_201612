import assert from "node:assert/strict";
import { test } from "node:test";

import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import JSZip from "jszip";

import type { WordMapping } from "./config.ts";
import { inspectWordDocument, isChineseText, WordAdapter } from "./word-adapter.ts";
import type { SegmentUpdate } from "./models.ts";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** Builds a minimal but valid .docx: just enough zip structure for JSZip/our adapter
 * to round-trip word/document.xml (content types + a couple of required parts aren't
 * needed since we only ever read/write document.xml here). */
async function buildDocx(bodyXml: string): Promise<ArrayBuffer> {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W_NS}"><w:body>${bodyXml}</w:body></w:document>`;
  const zip = new JSZip();
  zip.file("word/document.xml", documentXml);
  const buf = await zip.generateAsync({ type: "arraybuffer" });
  return buf as ArrayBuffer;
}

function paragraph(text: string): string {
  return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

function mapping(overrides: Partial<WordMapping> = {}): WordMapping {
  return { mode: "alternating_paragraphs", splitDelimiter: null, zhColumns: [], enColumns: [], ...overrides };
}

async function extractText(zip: JSZip): Promise<string> {
  const file = zip.file("word/document.xml");
  return file ? file.async("string") : "";
}

test("isChineseText", () => {
  assert.equal(isChineseText("這是一段中文說明"), true);
  assert.equal(isChineseText("This is English text"), false);
  assert.equal(isChineseText(""), false);
});

test("alternating paragraphs: extract and apply", async () => {
  const bytes = await buildDocx(paragraph("中文段落一") + paragraph("English paragraph one") + paragraph("中文段落二") + paragraph("English paragraph two"));

  const adapter = new WordAdapter(JSZip, DOMParser, XMLSerializer, mapping());
  const doc = await adapter.load(bytes);
  const segments = adapter.extractSegments(doc);

  assert.equal(segments.length, 2);
  assert.equal(segments[0].zhText, "中文段落一");
  assert.equal(segments[0].enText, "English paragraph one");
  assert.equal(segments[0].enLocationId, "para:1");

  const updates: SegmentUpdate[] = [{ enLocationId: segments[0].enLocationId, enText: "Revised paragraph one" }];
  adapter.applyTranslations(doc, updates);
  const outBytes = await adapter.save(doc);

  const reloaded = await JSZip.loadAsync(outBytes);
  const xml = await extractText(reloaded);
  assert.ok(xml.includes("Revised paragraph one"));
  assert.ok(xml.includes("中文段落二")); // untouched
});

test("alternating paragraphs: v1 zh-only document has no english yet (insert_after)", async () => {
  const bytes = await buildDocx(paragraph("中文段落一") + paragraph("中文段落二"));

  const adapter = new WordAdapter(JSZip, DOMParser, XMLSerializer, mapping());
  const doc = await adapter.load(bytes);
  const segments = adapter.extractSegments(doc);

  assert.equal(segments.length, 2);
  assert.equal(segments[0].enLocationId, "para:0:insert_after");
  assert.equal(segments[1].enLocationId, "para:1:insert_after");

  // Apply both translations in one call — exercises the descending-index ordering
  // that keeps insert_after targets valid even after earlier insertions.
  adapter.applyTranslations(doc, [
    { enLocationId: segments[0].enLocationId, enText: "English one" },
    { enLocationId: segments[1].enLocationId, enText: "English two" },
  ]);

  const outBytes = await adapter.save(doc);
  const reloaded = await adapter.load(outBytes);
  const reloadedSegments = new WordAdapter(JSZip, DOMParser, XMLSerializer, mapping()).extractSegments(reloaded);
  assert.equal(reloadedSegments.length, 2);
  assert.equal(reloadedSegments[0].zhText, "中文段落一");
  assert.equal(reloadedSegments[0].enText, "English one");
  assert.equal(reloadedSegments[1].zhText, "中文段落二");
  assert.equal(reloadedSegments[1].enText, "English two");
});

test("same paragraph split", async () => {
  const bytes = await buildDocx(paragraph("中文內容|English content"));
  const adapter = new WordAdapter(JSZip, DOMParser, XMLSerializer, mapping({ mode: "same_paragraph_split", splitDelimiter: "|" }));
  const doc = await adapter.load(bytes);
  const segments = adapter.extractSegments(doc);

  assert.equal(segments.length, 1);
  assert.equal(segments[0].zhText, "中文內容");
  assert.equal(segments[0].enText, "English content");

  adapter.applyTranslations(doc, [{ enLocationId: segments[0].enLocationId, enText: "New English content" }]);
  const outBytes = await adapter.save(doc);
  const xml = await extractText(await JSZip.loadAsync(outBytes));
  assert.ok(xml.includes("中文內容|New English content"));
});

test("table based", async () => {
  const tableXml =
    "<w:tbl>" +
    "<w:tr><w:tc><w:p><w:r><w:t>中文</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>English</w:t></w:r></w:p></w:tc></w:tr>" +
    "<w:tr><w:tc><w:p><w:r><w:t>第一項</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc></w:tr>" +
    "<w:tr><w:tc><w:p><w:r><w:t>第二項</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc></w:tr>" +
    "</w:tbl>";
  const bytes = await buildDocx(tableXml);

  const adapter = new WordAdapter(JSZip, DOMParser, XMLSerializer, mapping({ mode: "table_based", zhColumns: ["0"], enColumns: ["1"] }));
  const doc = await adapter.load(bytes);
  const segments = adapter.extractSegments(doc);

  assert.equal(segments.length, 2);
  assert.equal(segments[0].zhText, "第一項");
  assert.equal(segments[0].enLocationId, "table:0:1:1");

  adapter.applyTranslations(doc, [{ enLocationId: segments[0].enLocationId, enText: "Item one" }]);
  const outBytes = await adapter.save(doc);
  const xml = await extractText(await JSZip.loadAsync(outBytes));
  assert.ok(xml.includes("Item one"));
});

test("inspectWordDocument", async () => {
  const bytes = await buildDocx(paragraph("中文段落") + paragraph("English paragraph"));
  const rows = await inspectWordDocument(JSZip, DOMParser, bytes);
  assert.equal(rows[0].isChinese, true);
  assert.equal(rows[1].isChinese, false);
});

test("round trip preserves other parts of the zip untouched", async () => {
  const zip = new JSZip();
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W_NS}"><w:body>${paragraph("中文段落一") + paragraph("English paragraph one")}</w:body></w:document>`;
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", "<w:styles/>");
  const bytes = (await zip.generateAsync({ type: "arraybuffer" })) as ArrayBuffer;

  const adapter = new WordAdapter(JSZip, DOMParser, XMLSerializer, mapping());
  const doc = await adapter.load(bytes);
  const segments = adapter.extractSegments(doc);
  adapter.applyTranslations(doc, [{ enLocationId: segments[0].enLocationId, enText: "Revised" }]);
  const outBytes = await adapter.save(doc);

  const reloaded = await JSZip.loadAsync(outBytes);
  assert.ok(reloaded.file("word/styles.xml"), "unrelated zip parts must survive untouched");
  assert.equal(await reloaded.file("word/styles.xml")!.async("string"), "<w:styles/>");
});
