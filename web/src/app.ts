/** Page wiring for the single-page tool. Everything happens in one browser session
 * (select files -> review/translate -> download) so there's no need for the
 * process/apply-translations two-step state file the Python CLI needed to bridge two
 * separate process invocations — see plan Context.
 */

import type ExcelJSNS from "exceljs";
import type JSZipNS from "jszip";

import { alignAndDiff } from "./align.ts";
import { DEFAULT_SETTINGS, loadSettings, saveSettings } from "./config.ts";
import type { ExcelMapping, Settings, TrackingSheetMapping } from "./config.ts";
import { buildDiffReportWorkbook } from "./diff-report.ts";
import { ExcelAdapter } from "./excel-adapter.ts";
import { ChangeType } from "./models.ts";
import type { ChangeRecord, Segment, SegmentUpdate } from "./models.ts";
import { readRow, upsertRow } from "./tracking-sheet.ts";
import { buildZip } from "./zip-bundle.ts";
import type { ZipEntry } from "./zip-bundle.ts";

declare const ExcelJS: typeof ExcelJSNS;
declare const JSZip: typeof JSZipNS;

// ---- small DOM helpers ----------------------------------------------------

function $<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`missing element #${id}`);
  return el as T;
}

function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
  return file.arrayBuffer();
}

function setStatus(message: string, isError = false): void {
  const el = $<HTMLDivElement>("status");
  el.textContent = message;
  el.classList.toggle("error", isError);
}

function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

function splitColumns(value: string): string[] {
  return value
    .split(",")
    .map((s) => s.trim().toUpperCase())
    .filter(Boolean);
}

function todayCompact(): string {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
}

// ---- settings form ----------------------------------------------------------

const TRACKING_FIELD_LABELS: Record<keyof TrackingSheetMapping["columns"], string> = {
  versionNo: "版本編號",
  receivedDate: "收到日期",
  sender: "寄件窗口",
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

function renderTrackingColumnInputs(settings: Settings): void {
  const container = $<HTMLDivElement>("trackingColumnGrid");
  container.innerHTML = "";
  for (const key of Object.keys(TRACKING_FIELD_LABELS) as Array<keyof TrackingSheetMapping["columns"]>) {
    const wrapper = document.createElement("label");
    wrapper.className = "col-input";
    wrapper.textContent = TRACKING_FIELD_LABELS[key];
    const input = document.createElement("input");
    input.type = "text";
    input.size = 3;
    input.dataset.trackingCol = key;
    input.value = settings.trackingSheetMapping.columns[key];
    wrapper.appendChild(input);
    container.appendChild(wrapper);
  }
}

function fillSettingsForm(settings: Settings): void {
  $<HTMLInputElement>("excelSheetName").value = settings.excelMapping.sheetName;
  $<HTMLInputElement>("excelHeaderRow").value = String(settings.excelMapping.headerRow);
  $<HTMLInputElement>("excelKeyColumns").value = settings.excelMapping.keyColumns.join(",");
  $<HTMLInputElement>("excelZhColumns").value = settings.excelMapping.zhColumns.join(",");
  $<HTMLInputElement>("excelEnColumns").value = settings.excelMapping.enColumns.join(",");
  $<HTMLInputElement>("trackingSheetName").value = settings.trackingSheetMapping.sheetName;
  $<HTMLInputElement>("trackingHeaderRow").value = String(settings.trackingSheetMapping.headerRow);
  renderTrackingColumnInputs(settings);
}

function readSettingsForm(): Settings {
  const excelMapping: ExcelMapping = {
    sheetName: $<HTMLInputElement>("excelSheetName").value.trim() || "auto",
    headerRow: Number($<HTMLInputElement>("excelHeaderRow").value) || 1,
    keyColumns: splitColumns($<HTMLInputElement>("excelKeyColumns").value),
    zhColumns: splitColumns($<HTMLInputElement>("excelZhColumns").value),
    enColumns: splitColumns($<HTMLInputElement>("excelEnColumns").value),
  };

  const columns = { ...DEFAULT_SETTINGS.trackingSheetMapping.columns };
  for (const input of trackingColumnInputs()) {
    const key = input.dataset.trackingCol as keyof TrackingSheetMapping["columns"];
    columns[key] = input.value.trim().toUpperCase() || columns[key];
  }

  const trackingSheetMapping: TrackingSheetMapping = {
    sheetName: $<HTMLInputElement>("trackingSheetName").value.trim() || "追蹤",
    headerRow: Number($<HTMLInputElement>("trackingHeaderRow").value) || 1,
    columns,
  };

  return { excelMapping, trackingSheetMapping };
}

function trackingColumnInputs(): HTMLInputElement[] {
  return Array.from($<HTMLDivElement>("trackingColumnGrid").querySelectorAll("input[data-tracking-col]"));
}

// ---- application state ------------------------------------------------------

interface AnalysisState {
  excelMapping: ExcelMapping;
  newDocBytes: ArrayBuffer;
  newDocWorkbook: ExcelJSNS.Workbook;
  isFirstVersion: boolean;
  records: ChangeRecord[]; // empty for v1
  translationItems: Array<{ changeId: string; zhText: string; fieldName: string | null; oldEn?: string }>;
}

let state: AnalysisState | null = null;

// ---- step 1: analyze ---------------------------------------------------------

async function handleAnalyze(): Promise<void> {
  try {
    setStatus("分析中…");
    const settings = readSettingsForm();
    saveSettings(settings);

    const newDocFile = $<HTMLInputElement>("newDocFile").files?.[0];
    if (!newDocFile) throw new Error("請先選擇新版文件");
    const prevDocFile = $<HTMLInputElement>("prevDocFile").files?.[0];

    const adapter = new ExcelAdapter(ExcelJS, settings.excelMapping);
    const newDocBytes = await readFileAsArrayBuffer(newDocFile);
    const newDocWorkbook = await adapter.load(newDocBytes);
    const newSegments = adapter.extractSegments(newDocWorkbook);

    let records: ChangeRecord[] = [];
    let translationItems: AnalysisState["translationItems"];
    const isFirstVersion = !prevDocFile;

    if (isFirstVersion) {
      translationItems = newSegments
        .filter((s) => s.zhText.trim() !== "")
        .map((s) => ({ changeId: s.enLocationId, zhText: s.zhText, fieldName: s.fieldName }));
    } else {
      const prevBytes = await readFileAsArrayBuffer(prevDocFile);
      const prevWorkbook = await adapter.load(prevBytes);
      const prevSegments: Segment[] = adapter.extractSegments(prevWorkbook);
      records = alignAndDiff(prevSegments, newSegments);
      translationItems = records
        .filter((r) => (r.changeType === ChangeType.ADDED || r.changeType === ChangeType.MODIFIED) && r.newEnLocationId)
        .map((r) => ({
          changeId: r.newEnLocationId as string,
          zhText: r.newZh ?? "",
          fieldName: r.fieldName,
          // MODIFIED rows are pre-filled with the old translation as a starting point
          // to revise, mirroring how a human translator would work from the prior text
          oldEn: r.changeType === ChangeType.MODIFIED ? r.oldEn : undefined,
        }));
    }

    state = { excelMapping: settings.excelMapping, newDocBytes, newDocWorkbook, isFirstVersion, records, translationItems };

    renderDiffTable(records);
    renderTranslationTable(translationItems);
    $<HTMLInputElement>("versionNo").value = isFirstVersion ? "1" : $<HTMLInputElement>("versionNo").value || "";
    $<HTMLInputElement>("versionNo").readOnly = isFirstVersion;

    $<HTMLDivElement>("resultsSection").hidden = false;
    setStatus(
      isFirstVersion
        ? `首次文件，無前版可比對，共 ${translationItems.length} 筆需要翻譯。`
        : `比對完成，共 ${translationItems.length} 筆新增/修改需要翻譯。`
    );
  } catch (err) {
    setStatus(`發生錯誤：${(err as Error).message}`, true);
  }
}

function renderDiffTable(records: ChangeRecord[]): void {
  const section = $<HTMLDivElement>("diffSection");
  const tbody = $<HTMLTableSectionElement>("diffTableBody");
  tbody.innerHTML = "";
  const visible = records.filter((r) => r.changeType !== ChangeType.UNCHANGED);
  section.hidden = visible.length === 0;

  for (const r of visible) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escapeHtml(r.locationId)}</td><td>${escapeHtml(r.fieldName ?? "")}</td><td>${escapeHtml(r.oldZh ?? "")}</td><td>${escapeHtml(r.newZh ?? "")}</td><td>${escapeHtml(r.changeType)}</td>`;
    tbody.appendChild(tr);
  }
}

function renderTranslationTable(items: AnalysisState["translationItems"]): void {
  const tbody = $<HTMLTableSectionElement>("translationTableBody");
  tbody.innerHTML = "";
  for (const item of items) {
    const tr = document.createElement("tr");
    const zhTd = document.createElement("td");
    zhTd.textContent = item.zhText;
    const fieldTd = document.createElement("td");
    fieldTd.textContent = item.fieldName ?? "";
    const enTd = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    input.dataset.changeId = item.changeId;
    input.className = "translation-input";
    input.value = item.oldEn ?? "";
    enTd.appendChild(input);
    tr.append(fieldTd, zhTd, enTd);
    tbody.appendChild(tr);
  }
}

function escapeHtml(text: string): string {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

// ---- step 2: generate + download --------------------------------------------

async function handleGenerate(): Promise<void> {
  if (!state) {
    setStatus("請先執行「分析文件」", true);
    return;
  }
  try {
    setStatus("產生歸檔中…");

    const inputs = Array.from(
      $<HTMLTableSectionElement>("translationTableBody").querySelectorAll<HTMLInputElement>("input[data-change-id]")
    );
    const empty = inputs.filter((i) => i.value.trim() === "");
    if (empty.length > 0) throw new Error(`還有 ${empty.length} 筆尚未填寫英文翻譯`);

    const updates: SegmentUpdate[] = inputs.map((i) => ({
      enLocationId: i.dataset.changeId as string,
      enText: i.value.trim(),
    }));

    const adapter = new ExcelAdapter(ExcelJS, state.excelMapping);
    adapter.applyTranslations(state.newDocWorkbook, updates);
    const bilingualBytes = await adapter.save(state.newDocWorkbook);

    const versionNo = Number($<HTMLInputElement>("versionNo").value);
    if (!versionNo || versionNo < 1) throw new Error("請輸入正確的版本編號");

    const dateStamp = todayCompact();
    const entries: ZipEntry[] = [{ name: "bilingual.xlsx", bytes: bilingualBytes }];

    if (!state.isFirstVersion) {
      const diffBytes = await buildDiffReportWorkbook(ExcelJS, state.records);
      entries.push({ name: "diff_report.xlsx", bytes: diffBytes });
    }

    const msgFile = $<HTMLInputElement>("msgFile").files?.[0];
    if (msgFile) entries.push({ name: msgFile.name, bytes: await readFileAsArrayBuffer(msgFile) });

    const trackingFile = $<HTMLInputElement>("trackingFile").files?.[0];
    if (trackingFile) {
      const settings = readSettingsForm();
      const trackingBytes = await readFileAsArrayBuffer(trackingFile);
      const updatedTracking = await upsertRow(ExcelJS, trackingBytes, settings.trackingSheetMapping, versionNo, {
        versionNo,
        receivedDate: $<HTMLInputElement>("receivedDate").value,
        sender: $<HTMLInputElement>("sender").value,
        emailPath: msgFile?.name ?? "",
        bilingualFilePath: `v${String(versionNo).padStart(2, "0")}_${dateStamp}/bilingual.xlsx`,
        diffReportPath: state.isFirstVersion ? "" : `v${String(versionNo).padStart(2, "0")}_${dateStamp}/diff_report.xlsx`,
        updateSummary: state.isFirstVersion ? "首次提供文件" : "",
        isFirstVersion: state.isFirstVersion ? "是" : "否",
        overseasConfirmStatus: "未確認",
        customerConfirmStatus: "未確認",
        status: "待確認",
      });
      entries.push({ name: trackingFile.name, bytes: updatedTracking });
    }

    entries.push({
      name: "使用說明.txt",
      bytes: new TextEncoder().encode(buildReadme(versionNo, dateStamp, !!trackingFile)).buffer as ArrayBuffer,
    });

    const zipBlob = await buildZip(JSZip, entries);
    downloadBlob(zipBlob, `v${String(versionNo).padStart(2, "0")}_${dateStamp}.zip`);
    setStatus("已產生歸檔 zip，請解壓縮後依「使用說明.txt」放到正確資料夾。");
  } catch (err) {
    setStatus(`發生錯誤：${(err as Error).message}`, true);
  }
}

function buildReadme(versionNo: number, dateStamp: string, hasTracking: boolean): string {
  const folder = `v${String(versionNo).padStart(2, "0")}_${dateStamp}`;
  const lines = [
    `本次產出（版本 ${versionNo}）使用說明：`,
    "",
    `1. 把 bilingual.xlsx（及 diff_report.xlsx，若有）放進 archive/${folder}/ 資料夾`,
    "2. 同樣的 bilingual.xlsx 也複製一份到 output/ 資料夾，供上傳 Teams/SharePoint",
  ];
  if (hasTracking) {
    lines.push("3. 把追蹤表檔案覆蓋回原本存放的位置（取代舊檔）");
  }
  return lines.join("\n");
}

// ---- confirm section ---------------------------------------------------------

async function handleConfirm(): Promise<void> {
  try {
    setStatus("更新確認狀態中…");
    const trackingFile = $<HTMLInputElement>("confirmTrackingFile").files?.[0];
    if (!trackingFile) throw new Error("請先選擇追蹤表檔案");

    const versionNo = Number($<HTMLInputElement>("confirmVersionNo").value);
    const party = $<HTMLSelectElement>("confirmParty").value as "overseas" | "customer";
    const date = $<HTMLInputElement>("confirmDate").value;
    const method = $<HTMLInputElement>("confirmMethod").value;
    if (!versionNo || !date || !method) throw new Error("請完整填寫版本編號/日期/方式");

    const settings = readSettingsForm();
    const bytes = await readFileAsArrayBuffer(trackingFile);
    const current = (await readRow(ExcelJS, bytes, settings.trackingSheetMapping, versionNo)) ?? {};

    const other = party === "overseas" ? "customerConfirmStatus" : "overseasConfirmStatus";
    const updates: Record<string, string> = {
      [`${party}ConfirmStatus`]: "已確認",
      [`${party}ConfirmDate`]: date,
      [`${party}ConfirmMethod`]: method,
    };
    if (current[other] === "已確認") updates.status = "已完成";

    const updated = await upsertRow(ExcelJS, bytes, settings.trackingSheetMapping, versionNo, updates);
    downloadBlob(new Blob([updated]), trackingFile.name);
    setStatus(`已更新版本 ${versionNo} 的 ${party === "overseas" ? "海外團隊" : "客戶"} 確認狀態，請覆蓋回原檔案。`);
  } catch (err) {
    setStatus(`發生錯誤：${(err as Error).message}`, true);
  }
}

// ---- wiring -------------------------------------------------------------------

function init(): void {
  fillSettingsForm(loadSettings());
  $<HTMLButtonElement>("saveSettingsBtn").addEventListener("click", () => {
    saveSettings(readSettingsForm());
    setStatus("設定已儲存。");
  });
  $<HTMLButtonElement>("analyzeBtn").addEventListener("click", () => void handleAnalyze());
  $<HTMLButtonElement>("generateBtn").addEventListener("click", () => void handleGenerate());
  $<HTMLButtonElement>("confirmBtn").addEventListener("click", () => void handleConfirm());
}

document.addEventListener("DOMContentLoaded", init);
