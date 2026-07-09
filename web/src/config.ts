/** Mirrors config.py's ExcelMapping/TrackingSheetConfig — but since the browser tool
 * (Path B: upload/download, no client config file) exposes these as a settings form
 * instead of a YAML file, this module also owns default values + localStorage
 * persistence so the user only has to configure each sheet once per browser/device.
 *
 * Real customer workbooks have multiple tabs with genuinely different column layouts
 * (not just the same layout repeated), so the mapping is keyed per sheet name rather
 * than being a single flat mapping applied to "the" sheet.
 */

export interface SheetMapping {
  sheetName: string;
  /** "freeform" (default for newly-seen sheets): compare every cell, no columns to
   * configure — for screen mockups/report layouts, which real requirement documents
   * have far more of than actual zh/en tables. "table": the original column-role
   * picker flow, for sheets that really are a list of rows with a translatable column. */
  mode: "table" | "freeform";
  headerRow: number;
  keyColumns: string[];
  zhColumns: string[];
  enColumns: string[];
}

export interface TrackingColumns {
  versionNo: string;
  receivedDate: string;
  sender: string;
  sourceFileName: string;
  emailPath: string;
  bilingualFilePath: string;
  diffReportPath: string;
  updateSummary: string;
  isFirstVersion: string;
  overseasConfirmStatus: string;
  overseasConfirmDate: string;
  overseasConfirmMethod: string;
  customerConfirmStatus: string;
  customerConfirmDate: string;
  customerConfirmMethod: string;
  status: string;
  notes: string;
}

export interface TrackingSheetMapping {
  sheetName: string;
  headerRow: number;
  columns: TrackingColumns;
}

export interface WordMapping {
  /** "freeform" (default): whole-document compare, no zh/en split configured — for
   * documents that already mix Chinese and English inline with no clean delimiter
   * (common in practice — see plan Context). The other three are the original
   * structured layouts, for documents that really do separate the two languages. */
  mode: "freeform" | "alternating_paragraphs" | "same_paragraph_split" | "table_based";
  splitDelimiter: string | null;
  /** table_based only: 0-based column indices as strings, e.g. ["0"]. */
  zhColumns: string[];
  enColumns: string[];
}

export interface Settings {
  /** Keyed by worksheet name — each tab is configured independently. */
  sheetMappings: Record<string, SheetMapping>;
  trackingSheetMapping: TrackingSheetMapping;
  wordMapping: WordMapping;
}

export const TRACKING_FIELD_LABELS: Record<keyof TrackingColumns, string> = {
  versionNo: "版本編號",
  receivedDate: "收到日期",
  sender: "寄件窗口",
  sourceFileName: "客戶原始檔名",
  emailPath: "信件檔名",
  bilingualFilePath: "雙語檔檔名",
  diffReportPath: "差異報告檔名",
  updateSummary: "更新摘要",
  isFirstVersion: "是否首次文件",
  overseasConfirmStatus: "海外確認狀態",
  overseasConfirmDate: "海外確認日期",
  overseasConfirmMethod: "海外確認方式",
  customerConfirmStatus: "客戶確認狀態",
  customerConfirmDate: "客戶確認日期",
  customerConfirmMethod: "客戶確認方式",
  status: "狀態",
  notes: "備註",
};

export const DEFAULT_SETTINGS: Settings = {
  sheetMappings: {},
  wordMapping: {
    mode: "freeform",
    splitDelimiter: null,
    zhColumns: [],
    enColumns: [],
  },
  trackingSheetMapping: {
    sheetName: "追蹤",
    headerRow: 1,
    columns: {
      versionNo: "A",
      receivedDate: "B",
      sender: "C",
      sourceFileName: "Q",
      emailPath: "D",
      bilingualFilePath: "E",
      diffReportPath: "F",
      updateSummary: "G",
      isFirstVersion: "H",
      overseasConfirmStatus: "I",
      overseasConfirmDate: "J",
      overseasConfirmMethod: "K",
      customerConfirmStatus: "L",
      customerConfirmDate: "M",
      customerConfirmMethod: "N",
      status: "O",
      notes: "P",
    },
  },
};

const STORAGE_KEY = "docUpdateTool.settings.v2";

function clone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value));
}

export function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return clone(DEFAULT_SETTINGS);
    const parsed = JSON.parse(raw);
    return {
      ...clone(DEFAULT_SETTINGS),
      ...parsed,
      trackingSheetMapping: {
        ...clone(DEFAULT_SETTINGS.trackingSheetMapping),
        ...parsed.trackingSheetMapping,
        columns: { ...clone(DEFAULT_SETTINGS.trackingSheetMapping.columns), ...parsed.trackingSheetMapping?.columns },
      },
      wordMapping: { ...clone(DEFAULT_SETTINGS.wordMapping), ...parsed.wordMapping },
    };
  } catch {
    return clone(DEFAULT_SETTINGS);
  }
}

export function saveSettings(settings: Settings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

export function saveSheetMapping(settings: Settings, mapping: SheetMapping): Settings {
  const updated: Settings = { ...settings, sheetMappings: { ...settings.sheetMappings, [mapping.sheetName]: mapping } };
  saveSettings(updated);
  return updated;
}
