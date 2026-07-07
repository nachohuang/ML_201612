/** Mirrors diffing/align.py. There's no difflib.SequenceMatcher equivalent in JS, and we
 * deliberately avoid pulling in a third-party diff library (only exceljs/jszip are
 * approved dependencies for this tool) — so this hand-rolls a standard LCS-based diff.
 * It won't reproduce difflib's exact Ratcliff/Obershelp heuristics, but produces the
 * same equal/replace/insert/delete opcode shape, which is all align_and_diff needs.
 *
 * Only the Chinese side drives alignment/change detection — English is never
 * customer-authored (see plan Context), so comparing it would surface our own
 * translation choices as false "changes".
 */

import { ChangeType } from "./models.ts";
import type { ChangeRecord, Segment } from "./models.ts";
import { normalize } from "./text-normalize.ts";

export interface Opcode {
  tag: "equal" | "replace" | "insert" | "delete";
  i1: number;
  i2: number;
  j1: number;
  j2: number;
}

function lcsLengths(a: string[], b: string[]): number[][] {
  const n = a.length;
  const m = b.length;
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  return dp;
}

type Step = { tag: "equal" | "delete" | "insert"; i: number; j: number };

function walkSteps(a: string[], b: string[], dp: number[][]): Step[] {
  const steps: Step[] = [];
  let i = 0;
  let j = 0;
  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) {
      steps.push({ tag: "equal", i, j });
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      steps.push({ tag: "delete", i, j });
      i++;
    } else {
      steps.push({ tag: "insert", i, j });
      j++;
    }
  }
  while (i < a.length) {
    steps.push({ tag: "delete", i, j });
    i++;
  }
  while (j < b.length) {
    steps.push({ tag: "insert", i, j });
    j++;
  }
  return steps;
}

function groupSteps(steps: Step[]): Opcode[] {
  const runs: Opcode[] = [];
  let idx = 0;
  while (idx < steps.length) {
    const tag = steps[idx].tag;
    const start = idx;
    while (idx < steps.length && steps[idx].tag === tag) idx++;
    const first = steps[start];
    const last = steps[idx - 1];
    if (tag === "equal") {
      runs.push({ tag, i1: first.i, i2: last.i + 1, j1: first.j, j2: last.j + 1 });
    } else if (tag === "delete") {
      runs.push({ tag, i1: first.i, i2: last.i + 1, j1: first.j, j2: first.j });
    } else {
      runs.push({ tag, i1: first.i, i2: first.i, j1: first.j, j2: last.j + 1 });
    }
  }
  return runs;
}

/** Adjacent delete+insert runs (in either order) collapse into one "replace" opcode,
 * matching difflib's opcode vocabulary. */
function mergeReplaces(runs: Opcode[]): Opcode[] {
  const merged: Opcode[] = [];
  let idx = 0;
  while (idx < runs.length) {
    const cur = runs[idx];
    const next = runs[idx + 1];
    if (next && ((cur.tag === "delete" && next.tag === "insert") || (cur.tag === "insert" && next.tag === "delete"))) {
      const delRun = cur.tag === "delete" ? cur : next;
      const insRun = cur.tag === "insert" ? cur : next;
      merged.push({ tag: "replace", i1: delRun.i1, i2: delRun.i2, j1: insRun.j1, j2: insRun.j2 });
      idx += 2;
    } else {
      merged.push(cur);
      idx += 1;
    }
  }
  return merged;
}

export function getOpcodes(a: string[], b: string[]): Opcode[] {
  const dp = lcsLengths(a, b);
  const steps = walkSteps(a, b, dp);
  return mergeReplaces(groupSteps(steps));
}

function keys(segments: Segment[]): string[] {
  return segments.map((s, i) =>
    s.rowKey !== null ? `key:${normalize(s.rowKey)}:${s.fieldName ?? ""}` : `pos:${i}:${s.fieldName ?? ""}`
  );
}

export function alignAndDiff(oldSegments: Segment[], newSegments: Segment[]): ChangeRecord[] {
  const opcodes = getOpcodes(keys(oldSegments), keys(newSegments));
  const records: ChangeRecord[] = [];

  for (const { tag, i1, i2, j1, j2 } of opcodes) {
    if (tag === "equal") {
      for (let k = 0; k < i2 - i1; k++) records.push(diffPair(oldSegments[i1 + k], newSegments[j1 + k]));
    } else if (tag === "insert") {
      for (let j = j1; j < j2; j++) records.push(added(newSegments[j]));
    } else if (tag === "delete") {
      for (let i = i1; i < i2; i++) records.push(deleted(oldSegments[i]));
    } else {
      const oldSlice = oldSegments.slice(i1, i2);
      const newSlice = newSegments.slice(j1, j2);
      if (oldSlice.length === newSlice.length) {
        for (let k = 0; k < oldSlice.length; k++) records.push(diffPair(oldSlice[k], newSlice[k]));
      } else {
        for (const old of oldSlice) records.push(deleted(old));
        for (const nw of newSlice) records.push(added(nw));
      }
    }
  }
  return records;
}

function diffPair(oldSeg: Segment, newSeg: Segment): ChangeRecord {
  const changeType = normalize(oldSeg.zhText) === normalize(newSeg.zhText) ? ChangeType.UNCHANGED : ChangeType.MODIFIED;
  return {
    locationId: newSeg.locationId,
    fieldName: newSeg.fieldName,
    changeType,
    oldZh: oldSeg.zhText,
    newZh: newSeg.zhText,
    oldEn: oldSeg.enText,
    newEn: oldSeg.enText, // carried forward until a fresh translation is applied
    newEnLocationId: newSeg.enLocationId,
  };
}

function added(newSeg: Segment): ChangeRecord {
  return {
    locationId: newSeg.locationId,
    fieldName: newSeg.fieldName,
    changeType: ChangeType.ADDED,
    newZh: newSeg.zhText,
    newEnLocationId: newSeg.enLocationId,
  };
}

function deleted(oldSeg: Segment): ChangeRecord {
  return {
    locationId: oldSeg.locationId,
    fieldName: oldSeg.fieldName,
    changeType: ChangeType.DELETED,
    oldZh: oldSeg.zhText,
    oldEn: oldSeg.enText,
  };
}
