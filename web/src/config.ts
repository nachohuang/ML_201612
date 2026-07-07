/** Mirrors config.py's ExcelMapping/TrackingSheetConfig — but since the browser tool
 * (Path B: upload/download, no client config file) exposes these as a settings form
 * instead of a YAML file, this module also owns default values + localStorage
 * persistence so the user only has to fill the form in once per browser/device. */

export interface ExcelMapping {
  sheetName: string; // "auto" or an exact sheet name
  headerRow: number;
  keyColumns: string[];
  zhColumns: string[];
  enColumns: string[];
}

export interface TrackingColumns {
  versionNo: string;
  receivedDate: string;
  sender: string;
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

export interface Settings {
  excelMapping: ExcelMapping;
  trackingSheetMapping: TrackingSheetMapping;
}

export const DEFAULT_SETTINGS: Settings = {
  excelMapping: {
    sheetName: "auto",
    headerRow: 1,
    keyColumns: ["A"],
    zhColumns: ["C"],
    enColumns: ["D"],
  },
  trackingSheetMapping: {
    sheetName: "追蹤",
    headerRow: 1,
    columns: {
      versionNo: "A",
      receivedDate: "B",
      sender: "C",
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

const STORAGE_KEY = "docUpdateTool.settings.v1";

function clone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value));
}

export function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return clone(DEFAULT_SETTINGS);
    return { ...clone(DEFAULT_SETTINGS), ...JSON.parse(raw) };
  } catch {
    return clone(DEFAULT_SETTINGS);
  }
}

export function saveSettings(settings: Settings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}
