/** Mirrors src/doc_update_tool/models.py — kept in lockstep so the diff/tracking
 * logic ported below can be checked line-by-line against the Python original. */

export type ChangeType = "新增" | "修改" | "刪除" | "不變";

export const ChangeType = {
  ADDED: "新增" as ChangeType,
  MODIFIED: "修改" as ChangeType,
  DELETED: "刪除" as ChangeType,
  UNCHANGED: "不變" as ChangeType,
};

/** One zh/en pair extracted from a document (an Excel row-field). */
export interface Segment {
  /** Stable position identifier of the zh side (e.g. 'Sheet1!C12'), shown in diff reports. */
  locationId: string;
  /** Where the corresponding translation is written back to (e.g. 'Sheet1!D12'). */
  enLocationId: string;
  /** Content-based key used to align rows across versions (e.g. a requirement ID column). */
  rowKey: string | null;
  /** Human-readable column name, shown in diff reports. */
  fieldName: string | null;
  zhText: string;
  enText: string;
}

export interface ChangeRecord {
  locationId: string;
  fieldName: string | null;
  changeType: ChangeType;
  oldZh?: string;
  newZh?: string;
  oldEn?: string;
  newEn?: string;
  /** Where to write a fresh translation in the *new* document. Set for ADDED/MODIFIED;
   * undefined for DELETED (no corresponding location in the new document). */
  newEnLocationId?: string;
}

export interface TranslationItem {
  changeId: string;
  zhText: string;
  context?: string | null;
}

export interface TranslationResult {
  changeId: string;
  enText: string;
}

export interface SegmentUpdate {
  enLocationId: string;
  enText: string;
}
