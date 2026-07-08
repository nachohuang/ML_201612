/** Mirrors file_types/word_adapter.py (the Python CLI version), reimplemented over
 * raw OOXML instead of python-docx: a .docx is a zip of XML parts, and word/document.xml
 * is manipulated directly via DOMParser/XMLSerializer, mirroring python-docx's own
 * approach at a lower level (it also just wraps the same XML tree).
 *
 * JSZip is passed in the same way as ExcelJS elsewhere in this tool (dependency
 * injection: browser global via <script src="libs/jszip.min.js">, real npm package
 * under Node tests). DOMParser/XMLSerializer are native Web APIs with zero footprint in
 * the browser bundle; Node has neither built in, so unit tests inject @xmldom/xmldom's
 * implementations (a devDependency only — never shipped to the browser bundle).
 *
 * Supports three zh/en layouts, picked via WordMapping.mode — same three modes as the
 * Python version, same rationale for each (see that module's original docstring):
 * - alternating_paragraphs: a zh paragraph followed by its en paragraph. v1 documents
 *   are zh-only, so the en paragraph may not exist yet (insert_after handling below).
 * - same_paragraph_split: a single paragraph holding "zh<delimiter>en".
 * - table_based: a docx table with dedicated zh/en columns (like the Excel layout).
 *
 * Only top-level body paragraphs/tables are considered (matching python-docx's
 * `document.paragraphs`/`document.tables`, which exclude paragraphs/tables nested
 * inside another table's cell).
 */

import type JSZipNS from "jszip";

import type { WordMapping } from "./config.ts";
import type { MinimalDocument, MinimalDOMParser, MinimalElement, MinimalXMLSerializer } from "./minimal-dom.ts";
import type { Segment, SegmentUpdate } from "./models.ts";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DOCUMENT_XML_PATH = "word/document.xml";

export interface WordDoc {
  zip: JSZipNS;
  dom: MinimalDocument;
}

export function isChineseText(text: string, threshold = 0.3): boolean {
  const stripped = text.trim();
  if (!stripped) return false;
  const chars = [...stripped];
  const cjk = chars.filter((ch) => {
    const code = ch.codePointAt(0) ?? 0;
    return code >= 0x4e00 && code <= 0x9fff;
  }).length;
  return cjk / chars.length >= threshold;
}

export class WordAdapter {
  private readonly JSZip: typeof JSZipNS;
  private readonly DOMParserCtor: new () => MinimalDOMParser;
  private readonly XMLSerializerCtor: new () => MinimalXMLSerializer;
  private readonly mapping: WordMapping;

  constructor(
    JSZip: typeof JSZipNS,
    DOMParserCtor: new () => MinimalDOMParser,
    XMLSerializerCtor: new () => MinimalXMLSerializer,
    mapping: WordMapping
  ) {
    this.JSZip = JSZip;
    this.DOMParserCtor = DOMParserCtor;
    this.XMLSerializerCtor = XMLSerializerCtor;
    this.mapping = mapping;
  }

  async load(bytes: ArrayBuffer): Promise<WordDoc> {
    const zip = await this.JSZip.loadAsync(bytes);
    const file = zip.file(DOCUMENT_XML_PATH);
    if (!file) throw new Error("這不是有效的 .docx 檔案（找不到 word/document.xml）");
    const xmlText = await file.async("string");
    const dom = new this.DOMParserCtor().parseFromString(xmlText, "application/xml");
    assertValidXml(dom);
    return { zip, dom };
  }

  async save(doc: WordDoc): Promise<ArrayBuffer> {
    const xmlText = new this.XMLSerializerCtor().serializeToString(doc.dom);
    doc.zip.file(DOCUMENT_XML_PATH, xmlText);
    const out = await doc.zip.generateAsync({ type: "arraybuffer" });
    return out as ArrayBuffer;
  }

  extractSegments(doc: WordDoc): Segment[] {
    switch (this.mapping.mode) {
      case "freeform":
        return []; // handled separately via diffFreeformWordDocument, not a Segment flow
      case "alternating_paragraphs":
        return extractAlternating(doc.dom);
      case "same_paragraph_split":
        return extractSameSplit(doc.dom, this.mapping.splitDelimiter || "\n");
      case "table_based":
        return extractTableBased(doc.dom, this.mapping);
      default:
        throw new Error(`Unknown word mapping mode: ${this.mapping.mode as string}`);
    }
  }

  applyTranslations(doc: WordDoc, updates: SegmentUpdate[]): void {
    const direct: SegmentUpdate[] = [];
    const tails: SegmentUpdate[] = [];
    const inserts: SegmentUpdate[] = [];
    const tableUpdates: SegmentUpdate[] = [];

    for (const u of updates) {
      const parts = u.enLocationId.split(":");
      if (parts[0] === "para" && parts.length === 2) direct.push(u);
      else if (parts[0] === "para" && parts[parts.length - 1] === "tail") tails.push(u);
      else if (parts[0] === "para" && parts[parts.length - 1] === "insert_after") inserts.push(u);
      else if (parts[0] === "table") tableUpdates.push(u);
      else throw new Error(`Unrecognized en_location_id: ${u.enLocationId}`);
    }

    const paragraphs = getBodyParagraphs(doc.dom);

    for (const u of direct) {
      const idx = Number(u.enLocationId.split(":")[1]);
      setParagraphText(doc.dom, paragraphs[idx], u.enText);
    }

    const delimiter = this.mapping.splitDelimiter || "\n";
    for (const u of tails) {
      const idx = Number(u.enLocationId.split(":")[1]);
      const p = paragraphs[idx];
      const fullText = getParagraphText(p);
      const splitIdx = fullText.indexOf(delimiter);
      const zhPart = splitIdx === -1 ? fullText : fullText.slice(0, splitIdx);
      setParagraphText(doc.dom, p, `${zhPart}${delimiter}${u.enText}`);
    }

    // Insertions shift paragraph indices below the insertion point, so process
    // highest-index-first: every not-yet-processed index is still below every
    // insertion point already handled, so it can't have shifted yet (same reasoning
    // as the Python version's identical ordering rule).
    for (const u of [...inserts].sort((a, b) => Number(b.enLocationId.split(":")[1]) - Number(a.enLocationId.split(":")[1]))) {
      const idx = Number(u.enLocationId.split(":")[1]);
      insertParagraphAfter(doc.dom, paragraphs[idx], u.enText);
    }

    if (tableUpdates.length > 0) {
      const tables = getBodyTables(doc.dom);
      for (const u of tableUpdates) {
        const [, tIdxStr, rIdxStr, cIdxStr] = u.enLocationId.split(":");
        const table = tables[Number(tIdxStr)];
        const rows = getDirectChildElements(table, "tr");
        const row = rows[Number(rIdxStr)];
        const cells = getDirectChildElements(row, "tc");
        setCellText(doc.dom, cells[Number(cIdxStr)], u.enText);
      }
    }
  }
}

function assertValidXml(dom: MinimalDocument): void {
  const errorNode = dom.getElementsByTagName("parsererror")[0];
  if (errorNode) throw new Error(`無法解析 Word 文件內容: ${errorNode.textContent ?? ""}`);
}

export function getBodyParagraphs(dom: MinimalDocument): MinimalElement[] {
  const body = dom.getElementsByTagNameNS(W_NS, "body")[0];
  return getDirectChildElements(body, "p");
}

export function getBodyTables(dom: MinimalDocument): MinimalElement[] {
  const body = dom.getElementsByTagNameNS(W_NS, "body")[0];
  return getDirectChildElements(body, "tbl");
}

export function getDirectChildElements(parent: MinimalElement, localName: string): MinimalElement[] {
  const result: MinimalElement[] = [];
  for (const node of Array.from(parent.childNodes)) {
    if (node.nodeType === 1 && (node as MinimalElement).localName === localName) result.push(node as MinimalElement);
  }
  return result;
}

export function getParagraphText(p: MinimalElement): string {
  const texts = p.getElementsByTagNameNS(W_NS, "t");
  let out = "";
  for (let i = 0; i < texts.length; i++) out += texts[i].textContent ?? "";
  return out;
}

export function getCellText(cell: MinimalElement): string {
  return getDirectChildElements(cell, "p")
    .map((p) => getParagraphText(p))
    .join("\n");
}

/** Overwrites a paragraph's visible text. Keeps the first run's formatting (font,
 * bold, etc.) and clears any other runs — full run-level fidelity isn't preserved
 * when the translated text doesn't align with original run boundaries (same trade-off
 * as the Python version). */
function setParagraphText(dom: MinimalDocument, p: MinimalElement, text: string): void {
  const runs = getDirectChildElements(p, "r");
  if (runs.length === 0) {
    p.appendChild(buildRun(dom, text));
    return;
  }
  const firstRunTexts = runs[0].getElementsByTagNameNS(W_NS, "t");
  if (firstRunTexts.length > 0) {
    firstRunTexts[0].textContent = text;
    firstRunTexts[0].setAttribute("xml:space", "preserve");
    for (let i = 1; i < firstRunTexts.length; i++) firstRunTexts[i].textContent = "";
  } else {
    const t = dom.createElementNS(W_NS, "w:t");
    t.setAttribute("xml:space", "preserve");
    t.textContent = text;
    runs[0].appendChild(t);
  }
  for (let i = 1; i < runs.length; i++) {
    const texts = runs[i].getElementsByTagNameNS(W_NS, "t");
    for (let j = 0; j < texts.length; j++) texts[j].textContent = "";
  }
}

function setCellText(dom: MinimalDocument, cell: MinimalElement, text: string): void {
  const paragraphs = getDirectChildElements(cell, "p");
  if (paragraphs.length === 0) {
    const p = dom.createElementNS(W_NS, "w:p");
    cell.appendChild(p);
    setParagraphText(dom, p, text);
    return;
  }
  setParagraphText(dom, paragraphs[0], text);
  // python-docx's `cell.text = value` collapses a multi-paragraph cell down to one —
  // mirror that rather than leaving stale extra paragraphs behind.
  for (let i = 1; i < paragraphs.length; i++) cell.removeChild(paragraphs[i]);
}

function buildRun(dom: MinimalDocument, text: string): MinimalElement {
  const run = dom.createElementNS(W_NS, "w:r");
  const t = dom.createElementNS(W_NS, "w:t");
  t.setAttribute("xml:space", "preserve");
  t.textContent = text;
  run.appendChild(t);
  return run;
}

function insertParagraphAfter(dom: MinimalDocument, referenceP: MinimalElement, text: string): void {
  const newP = dom.createElementNS(W_NS, "w:p");
  newP.appendChild(buildRun(dom, text));
  referenceP.parentNode!.insertBefore(newP, referenceP.nextSibling);
}

function extractAlternating(dom: MinimalDocument): Segment[] {
  const paragraphs = getBodyParagraphs(dom);
  const n = paragraphs.length;
  const segments: Segment[] = [];
  let i = 0;
  while (i < n) {
    const text = getParagraphText(paragraphs[i]);
    if (!isChineseText(text)) {
      i++;
      continue;
    }
    const zhIdx = i;
    let j = i + 1;
    while (j < n && getParagraphText(paragraphs[j]).trim() === "") j++;

    let enIdx: number | null = null;
    let nextI: number;
    if (j < n && getParagraphText(paragraphs[j]).trim() !== "" && !isChineseText(getParagraphText(paragraphs[j]))) {
      enIdx = j;
      nextI = j + 1;
    } else {
      nextI = i + 1;
    }

    segments.push({
      locationId: `para:${zhIdx}`,
      enLocationId: enIdx !== null ? `para:${enIdx}` : `para:${zhIdx}:insert_after`,
      rowKey: null,
      fieldName: null,
      zhText: getParagraphText(paragraphs[zhIdx]),
      enText: enIdx !== null ? getParagraphText(paragraphs[enIdx]) : "",
    });
    i = nextI;
  }
  return segments;
}

function extractSameSplit(dom: MinimalDocument, delimiter: string): Segment[] {
  const paragraphs = getBodyParagraphs(dom);
  const segments: Segment[] = [];
  paragraphs.forEach((p, idx) => {
    const text = getParagraphText(p);
    let zhPart: string;
    let enPart: string;
    const splitIdx = text.indexOf(delimiter);
    if (splitIdx !== -1) {
      zhPart = text.slice(0, splitIdx);
      enPart = text.slice(splitIdx + delimiter.length);
    } else if (isChineseText(text)) {
      zhPart = text;
      enPart = "";
    } else {
      return;
    }
    segments.push({
      locationId: `para:${idx}:head`,
      enLocationId: `para:${idx}:tail`,
      rowKey: null,
      fieldName: null,
      zhText: zhPart.trim(),
      enText: enPart.trim(),
    });
  });
  return segments;
}

function extractTableBased(dom: MinimalDocument, mapping: WordMapping): Segment[] {
  const tables = getBodyTables(dom);
  const segments: Segment[] = [];
  tables.forEach((table, tIdx) => {
    const rows = getDirectChildElements(table, "tr");
    rows.forEach((row, rIdx) => {
      if (rIdx === 0) return; // header row
      const cells = getDirectChildElements(row, "tc");
      mapping.zhColumns.forEach((zhC, i) => {
        const enC = mapping.enColumns[i];
        const zhCell = cells[Number(zhC)];
        const enCell = cells[Number(enC)];
        if (!zhCell || !enCell) return;
        const zhText = getCellText(zhCell);
        if (!zhText.trim()) return;
        segments.push({
          locationId: `table:${tIdx}:${rIdx}:${zhC}`,
          enLocationId: `table:${tIdx}:${rIdx}:${enC}`,
          rowKey: null,
          fieldName: null,
          zhText,
          enText: getCellText(enCell),
        });
      });
    });
  });
  return segments;
}

export interface InspectParagraph {
  index: number;
  length: number;
  cjkRatio: number;
  isChinese: boolean;
  preview: string;
}

/** Lists each paragraph's index/CJK ratio/preview, to help pick word_mapping.mode
 * without guessing at a real document's layout. */
export async function inspectWordDocument(
  JSZip: typeof JSZipNS,
  DOMParserCtor: new () => MinimalDOMParser,
  bytes: ArrayBuffer,
  sampleChars = 60
): Promise<InspectParagraph[]> {
  const zip = await JSZip.loadAsync(bytes);
  const file = zip.file(DOCUMENT_XML_PATH);
  if (!file) throw new Error("這不是有效的 .docx 檔案（找不到 word/document.xml）");
  const xmlText = await file.async("string");
  const dom = new DOMParserCtor().parseFromString(xmlText, "application/xml");
  assertValidXml(dom);

  return getBodyParagraphs(dom).map((p, idx) => {
    const text = getParagraphText(p);
    const chars = [...text];
    const cjk = chars.filter((ch) => {
      const code = ch.codePointAt(0) ?? 0;
      return code >= 0x4e00 && code <= 0x9fff;
    }).length;
    return {
      index: idx,
      length: chars.length,
      cjkRatio: chars.length ? Math.round((cjk / chars.length) * 100) / 100 : 0,
      isChinese: isChineseText(text),
      preview: text.slice(0, sampleChars),
    };
  });
}
