/** Whole-document diff for Word content that isn't a clean zh/en split — real
 * requirement documents often write Chinese immediately followed by English inline
 * within the same paragraph or table cell (e.g. "計算Calculate"), with no delimiter and
 * no separate paragraph/column per language. None of the three configured
 * word_mapping modes fit that pattern, but the content still needs its changes
 * tracked (same principle as freeform Excel sheets — see plan Context: this is
 * meaningful requirement content, not something to silently ignore).
 *
 * Paragraphs (and table rows) are aligned with the same LCS-based approach as
 * diffing/align.ts, using each paragraph/row's own normalized text as its alignment
 * key — not fixed position. A naive "compare index i to index i" was tried first and
 * failed badly on a real document: inserting one new paragraph shifts every
 * subsequent index, so it misread "one paragraph added" as "every paragraph after it
 * modified." Aligning by content avoids that.
 *
 * A changed unit is reported for human review, but — like freeform Excel cells —
 * there's no separate English slot to translate into, so these records never populate
 * the translation step.
 */

import { getOpcodes } from "./align.ts";
import { ChangeType } from "./models.ts";
import type { ChangeRecord } from "./models.ts";
import { normalize } from "./text-normalize.ts";
import { getBodyParagraphs, getBodyTables, getCellText, getDirectChildElements, getParagraphText } from "./word-adapter.ts";
import type { MinimalDocument, MinimalElement } from "./minimal-dom.ts";

function added(locationId: string, text: string): ChangeRecord {
  return { locationId, fieldName: null, changeType: ChangeType.ADDED, newZh: text };
}
function deleted(locationId: string, text: string): ChangeRecord {
  return { locationId, fieldName: null, changeType: ChangeType.DELETED, oldZh: text };
}
function modified(locationId: string, oldText: string, newText: string): ChangeRecord {
  return { locationId, fieldName: null, changeType: ChangeType.MODIFIED, oldZh: oldText, newZh: newText };
}

/** Filters out blank entries (nothing meaningful to diff) while remembering each
 * surviving entry's original index, so location IDs still point at the real
 * paragraph/row position after filtering. */
function nonBlank(texts: string[]): { texts: string[]; indices: number[] } {
  const kept: string[] = [];
  const indices: number[] = [];
  texts.forEach((text, idx) => {
    if (text.trim() !== "") {
      kept.push(text);
      indices.push(idx);
    }
  });
  return { texts: kept, indices };
}

function diffAlignedText(
  oldTexts: string[],
  newTexts: string[],
  locationIdFor: (idx: number) => string
): ChangeRecord[] {
  const old_ = nonBlank(oldTexts);
  const new_ = nonBlank(newTexts);
  const opcodes = getOpcodes(old_.texts.map(normalize), new_.texts.map(normalize));
  const records: ChangeRecord[] = [];

  for (const { tag, i1, i2, j1, j2 } of opcodes) {
    if (tag === "equal") continue;
    if (tag === "insert") {
      for (let j = j1; j < j2; j++) records.push(added(locationIdFor(new_.indices[j]), new_.texts[j]));
    } else if (tag === "delete") {
      for (let i = i1; i < i2; i++) records.push(deleted(locationIdFor(old_.indices[i]), old_.texts[i]));
    } else {
      const oldLen = i2 - i1;
      const newLen = j2 - j1;
      if (oldLen === newLen) {
        for (let k = 0; k < oldLen; k++) {
          records.push(modified(locationIdFor(new_.indices[j1 + k]), old_.texts[i1 + k], new_.texts[j1 + k]));
        }
      } else {
        for (let i = i1; i < i2; i++) records.push(deleted(locationIdFor(old_.indices[i]), old_.texts[i]));
        for (let j = j1; j < j2; j++) records.push(added(locationIdFor(new_.indices[j]), new_.texts[j]));
      }
    }
  }
  return records;
}

function diffTableRows(oldRows: MinimalElement[], newRows: MinimalElement[], tableIdx: number): ChangeRecord[] {
  const rowKey = (row: MinimalElement) =>
    getDirectChildElements(row, "tc")
      .map((c) => getCellText(c).trim())
      .join("|");

  const opcodes = getOpcodes(oldRows.map((r) => normalize(rowKey(r))), newRows.map((r) => normalize(rowKey(r))));
  const records: ChangeRecord[] = [];

  const emitRow = (row: MinimalElement, rIdx: number, kind: "added" | "deleted") => {
    getDirectChildElements(row, "tc").forEach((cell, cIdx) => {
      const text = getCellText(cell).trim();
      if (!text) return;
      const locationId = `table:${tableIdx}:${rIdx}:${cIdx}`;
      records.push(kind === "added" ? added(locationId, text) : deleted(locationId, text));
    });
  };

  for (const { tag, i1, i2, j1, j2 } of opcodes) {
    if (tag === "equal") continue;
    if (tag === "insert") {
      for (let j = j1; j < j2; j++) emitRow(newRows[j], j, "added");
    } else if (tag === "delete") {
      for (let i = i1; i < i2; i++) emitRow(oldRows[i], i, "deleted");
    } else {
      const oldLen = i2 - i1;
      const newLen = j2 - j1;
      if (oldLen === newLen) {
        for (let k = 0; k < oldLen; k++) {
          const oldCells = getDirectChildElements(oldRows[i1 + k], "tc");
          const newCells = getDirectChildElements(newRows[j1 + k], "tc");
          for (let c = 0; c < Math.max(oldCells.length, newCells.length); c++) {
            const oldText = oldCells[c] ? getCellText(oldCells[c]).trim() : "";
            const newText = newCells[c] ? getCellText(newCells[c]).trim() : "";
            if (oldText === newText) continue;
            const locationId = `table:${tableIdx}:${j1 + k}:${c}`;
            if (oldText === "") records.push(added(locationId, newText));
            else if (newText === "") records.push(deleted(locationId, oldText));
            else records.push(modified(locationId, oldText, newText));
          }
        }
      } else {
        for (let i = i1; i < i2; i++) emitRow(oldRows[i], i, "deleted");
        for (let j = j1; j < j2; j++) emitRow(newRows[j], j, "added");
      }
    }
  }
  return records;
}

export function diffFreeformWordDocument(oldDom: MinimalDocument | undefined, newDom: MinimalDocument): ChangeRecord[] {
  const records: ChangeRecord[] = [];

  const oldParagraphs = oldDom ? getBodyParagraphs(oldDom) : [];
  const newParagraphs = getBodyParagraphs(newDom);
  records.push(
    ...diffAlignedText(
      oldParagraphs.map(getParagraphText),
      newParagraphs.map(getParagraphText),
      (idx) => `para:${idx}`
    )
  );

  const oldTables = oldDom ? getBodyTables(oldDom) : [];
  const newTables = getBodyTables(newDom);
  // Tables themselves are matched by position (whole-table insert/reorder is rare
  // compared to a paragraph or row being added) — only rows within a matched pair get
  // the full LCS alignment treatment.
  for (let t = 0; t < Math.max(oldTables.length, newTables.length); t++) {
    const oldRows = oldTables[t] ? getDirectChildElements(oldTables[t], "tr") : [];
    const newRows = newTables[t] ? getDirectChildElements(newTables[t], "tr") : [];
    records.push(...diffTableRows(oldRows, newRows, t));
  }

  return records;
}
