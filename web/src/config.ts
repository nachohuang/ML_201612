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

export interface Settings {
  /** Keyed by worksheet name — each tab is configured independently. */
  sheetMappings: Record<string, SheetMapping>;
  trackingSheetMapping: TrackingSheetMapping;
}

export const DEFAULT_SETTINGS: Settings = {
  sheetMappings: {},
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
