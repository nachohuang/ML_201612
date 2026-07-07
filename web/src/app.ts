/** Page wiring for the single-page tool. Everything happens in one browser session
 * (select files -> review/translate -> download) so there's no need for the
 * process/apply-translations two-step state file the Python CLI needed to bridge two
 * separate process invocations — see plan Context.
 */

import type ExcelJSNS from "exceljs";
import type JSZipNS from "jszip";

import { alignAndDiff } from "./align.ts";
import { DEFAULT_SETTINGS, loadSettings, saveSettings } from "./config.ts";
import type { Settings, SheetMapping, TrackingSheetMapping } from "./config.ts";
import { buildDiffReportWorkbook } from "./diff-report.ts";
import { parseEmlFile } from "./eml-parser.ts";
import { ExcelAdapter, inspectWorkbook } from "./excel-adapter.ts";
import type { InspectSheet } from "./excel-adapter.ts";
import { ChangeType } from "./models.ts";
import type { ChangeRecord, Segment, SegmentUpdate } from "./models.ts";
import { listEntries, readRow, upsertRow } from "./tracking-sheet.ts";
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

function todayCompact(): string {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
}

function escapeHtml(text: string): string {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

/** Wires a drop zone element so dragging a file onto it (or clicking it) populates the
 * given file input — the native "Choose File" button inside keeps working as a fallback. */
function wireDropZone(zoneId: string, inputId: string): void {
  const zone = $<HTMLElement>(zoneId);
  const input = $<HTMLInputElement>(inputId);
  zone.addEventListener("dragover", (e) => {
    e.preventDefault();
    zone.classList.add("drag-over");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("drag-over"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("drag-over");
    if (e.dataTransfer?.files?.length) {
      input.files = e.dataTransfer.files;
      input.dispatchEvent(new Event("change", { bubbles: true }));
    }
  });
}

// ---- tracking sheet settings form (fixed set of fields, kept as a simple form) -----

const TRACKING_FIELD_LABELS: Record<keyof TrackingSheetMapping["columns"], string> = {
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

function fillTrackingSettingsForm(settings: Settings): void {
  $<HTMLInputElement>("trackingSheetName").value = settings.trackingSheetMapping.sheetName;
  $<HTMLInputElement>("trackingHeaderRow").value = String(settings.trackingSheetMapping.headerRow);
  renderTrackingColumnInputs(settings);
}

function readTrackingSettingsForm(): TrackingSheetMapping {
  const columns = { ...DEFAULT_SETTINGS.trackingSheetMapping.columns };
  for (const input of trackingColumnInputs()) {
    const key = input.dataset.trackingCol as keyof TrackingSheetMapping["columns"];
    columns[key] = input.value.trim().toUpperCase() || columns[key];
  }
  return {
    sheetName: $<HTMLInputElement>("trackingSheetName").value.trim() || "追蹤",
    headerRow: Number($<HTMLInputElement>("trackingHeaderRow").value) || 1,
    columns,
  };
}

function trackingColumnInputs(): HTMLInputElement[] {
  return Array.from($<HTMLDivElement>("trackingColumnGrid").querySelectorAll("input[data-tracking-col]"));
}

// ---- per-sheet column mapping (interactive, point-and-click) ----------------------

const ROLE_OPTIONS: Array<[string, string]> = [
  ["ignore", "忽略"],
  ["key", "唯一鍵"],
  ["zh1", "中文（配對1）"],
  ["en1", "英文（配對1）"],
  ["zh2", "中文（配對2）"],
  ["en2", "英文（配對2）"],
  ["zh3", "中文（配對3）"],
  ["en3", "英文（配對3）"],
];

function roleForColumn(col: string, existing?: SheetMapping): string {
  if (!existing) return "ignore";
  if (existing.keyColumns.includes(col)) return "key";
  const zhIdx = existing.zhColumns.indexOf(col);
  if (zhIdx !== -1) return `zh${zhIdx + 1}`;
  const enIdx = existing.enColumns.indexOf(col);
  if (enIdx !== -1) return `en${enIdx + 1}`;
  return "ignore";
}

function renderSheetMappingCards(sheets: InspectSheet[], existing: Record<string, SheetMapping>): void {
  const container = $<HTMLDivElement>("sheetMappingContainer");
  container.innerHTML = "";

  for (const sheet of sheets) {
    const card = document.createElement("div");
    card.className = "sheet-mapping-card";
    card.dataset.sheetName = sheet.sheetName;

    const heading = document.createElement("h3");
    heading.textContent = `頁籤：${sheet.sheetName}`;
    card.appendChild(heading);

    const headerRowLabel = document.createElement("label");
    headerRowLabel.textContent = "表頭列號：";
    const headerRowInput = document.createElement("input");
    headerRowInput.type = "number";
    headerRowInput.min = "1";
    headerRowInput.value = String(existing[sheet.sheetName]?.headerRow ?? 1);
    headerRowInput.className = "sheet-header-row";
    headerRowInput.style.width = "4em";
    headerRowLabel.appendChild(headerRowInput);
    card.appendChild(headerRowLabel);

    const table = document.createElement("table");
    const theadRow = document.createElement("tr");
    theadRow.innerHTML = "<th>欄位</th>" + sheet.columns.map((c) => `<th>${escapeHtml(c.column)}</th>`).join("");
    table.appendChild(theadRow);

    const previewLabels = ["第1列", "第2列", "第3列", "第4列"];
    for (let r = 0; r < 4; r++) {
      const row = document.createElement("tr");
      const cells = sheet.columns
        .map((c) => `<td>${escapeHtml(String((r === 0 ? c.header : c.samples[r - 1]) ?? ""))}</td>`)
        .join("");
      row.innerHTML = `<td class="hint">${previewLabels[r]}</td>${cells}`;
      table.appendChild(row);
    }

    const roleRow = document.createElement("tr");
    const roleLabelCell = document.createElement("td");
    roleLabelCell.textContent = "欄位用途";
    roleRow.appendChild(roleLabelCell);
    for (const col of sheet.columns) {
      const td = document.createElement("td");
      const select = document.createElement("select");
      select.dataset.col = col.column;
      for (const [value, label] of ROLE_OPTIONS) {
        const opt = document.createElement("option");
        opt.value = value;
        opt.textContent = label;
        select.appendChild(opt);
      }
      select.value = roleForColumn(col.column, existing[sheet.sheetName]);
      td.appendChild(select);
      roleRow.appendChild(td);
    }
    table.appendChild(roleRow);

    card.appendChild(table);
    container.appendChild(card);
  }

  $<HTMLDivElement>("sheetMappingSection").hidden = sheets.length === 0;
}

function collectSheetMappingsFromUI(): Record<string, SheetMapping> {
  const cards = Array.from(document.querySelectorAll<HTMLElement>(".sheet-mapping-card"));
  const result: Record<string, SheetMapping> = {};

  for (const card of cards) {
    const sheetName = card.dataset.sheetName as string;
    const headerRow = Number(card.querySelector<HTMLInputElement>(".sheet-header-row")!.value) || 1;
    const keyColumns: string[] = [];
    const pairs: Record<string, { zh?: string; en?: string }> = {};

    for (const select of Array.from(card.querySelectorAll<HTMLSelectElement>("select[data-col]"))) {
      const col = select.dataset.col as string;
      const role = select.value;
      if (role === "key") keyColumns.push(col);
      else if (role.startsWith("zh")) (pairs[role.slice(2)] ??= {}).zh = col;
      else if (role.startsWith("en")) (pairs[role.slice(2)] ??= {}).en = col;
    }

    const zhColumns: string[] = [];
    const enColumns: string[] = [];
    for (const key of Object.keys(pairs).sort()) {
      const pair = pairs[key];
      if (pair.zh && pair.en) {
        zhColumns.push(pair.zh);
        enColumns.push(pair.en);
      }
    }

    result[sheetName] = { sheetName, headerRow, keyColumns, zhColumns, enColumns };
  }
  return result;
}

async function handleNewDocSelected(): Promise<void> {
  const file = $<HTMLInputElement>("newDocFile").files?.[0];
  if (!file) return;
  try {
    const bytes = await readFileAsArrayBuffer(file);
    const sheets = await inspectWorkbook(ExcelJS, bytes);
    const settings = loadSettings();
    renderSheetMappingCards(sheets, settings.sheetMappings);
  } catch (err) {
    setStatus(`讀取文件失敗：${(err as Error).message}`, true);
  }
  await checkVersionMatch();
}

// ---- eml parsing -------------------------------------------------------------------

async function handleEmlSelected(): Promise<void> {
  const file = $<HTMLInputElement>("emlFile").files?.[0];
  if (!file) return;
  try {
    const meta = await parseEmlFile(file);
    if (meta.sender) $<HTMLInputElement>("sender").value = meta.sender;
    if (meta.receivedDate) $<HTMLInputElement>("receivedDate").value = meta.receivedDate;
    setStatus(
      meta.sender || meta.receivedDate
        ? "已從信件自動帶入寄件者/日期，請確認正確，需要的話可自行修改。"
        : "無法從這份信件解析出寄件者/日期，請手動填寫。"
    );
  } catch {
    setStatus("無法解析這份 .eml 信件，請手動填寫寄件者/日期。", true);
  }
}

// ---- version matching by filename against the tracking sheet ----------------------

let resolvedVersionNo: number | null = null;

// handleNewDocSelected() and the trackingFile input's own "change" listener can both
// call checkVersionMatch() around the same time (e.g. selecting a new document also
// re-checks the already-uploaded tracking file). Without this guard, a slower call
// started earlier could resolve later and clobber a newer call's result with stale
// data, since both mutate the same DOM state.
let versionMatchRequestId = 0;

async function checkVersionMatch(): Promise<void> {
  const requestId = ++versionMatchRequestId;
  const trackingFile = $<HTMLInputElement>("trackingFile").files?.[0];
  const newDocFile = $<HTMLInputElement>("newDocFile").files?.[0];
  const section = $<HTMLDivElement>("versionMatchSection");

  if (!trackingFile || !newDocFile) {
    section.hidden = true;
    resolvedVersionNo = null;
    return;
  }

  let entries;
  try {
    const trackingBytes = await readFileAsArrayBuffer(trackingFile);
    const mapping = readTrackingSettingsForm();
    entries = await listEntries(ExcelJS, trackingBytes, mapping);
  } catch {
    // A File handle can become unreadable if the input was cleared/reselected while
    // this (now-superseded) read was in flight — harmless as long as a newer call is
    // the one that ends up applying its result, so just bail out quietly here.
    return;
  }

  if (requestId !== versionMatchRequestId) return; // superseded by a newer call

  const exact = entries.find((e) => e.sourceFileName === newDocFile.name);
  if (exact) {
    resolvedVersionNo = exact.versionNo;
    section.hidden = true;
    $<HTMLInputElement>("versionNo").value = String(exact.versionNo);
    setStatus(`追蹤表中找到相同檔名，將更新版本 ${exact.versionNo} 的紀錄。`);
    return;
  }

  if (entries.length === 0) {
    resolvedVersionNo = 1;
    section.hidden = true;
    $<HTMLInputElement>("versionNo").value = "1";
    return;
  }

  // No exact match but the tracking sheet has history — don't guess, ask.
  resolvedVersionNo = null;
  section.hidden = false;
  const select = $<HTMLSelectElement>("matchExistingSelect");
  select.innerHTML = '<option value="">-- 請選擇這是延續哪一份既有文件 --</option>';
  for (const entry of entries) {
    const opt = document.createElement("option");
    opt.value = String(entry.versionNo);
    opt.textContent = `版本 ${entry.versionNo}：${entry.sourceFileName || "(無檔名記錄)"}`;
    select.appendChild(opt);
  }
  select.value = "";
  $<HTMLInputElement>("confirmNewFileCheckbox").checked = false;
  $<HTMLElement>("versionMatchInfo").textContent = `在追蹤表中找不到檔名「${newDocFile.name}」。`;
}

function wireVersionMatchControls(): void {
  const select = $<HTMLSelectElement>("matchExistingSelect");
  const checkbox = $<HTMLInputElement>("confirmNewFileCheckbox");

  select.addEventListener("change", () => {
    if (select.value) {
      checkbox.checked = false;
      const matchedVersion = Number(select.value);
      resolvedVersionNo = matchedVersion + 1;
      $<HTMLInputElement>("versionNo").value = String(resolvedVersionNo);
    } else {
      resolvedVersionNo = null;
    }
  });

  checkbox.addEventListener("change", () => {
    if (checkbox.checked) {
      select.value = "";
      resolvedVersionNo = 1;
      $<HTMLInputElement>("versionNo").value = "1";
    } else {
      resolvedVersionNo = null;
    }
  });
}

// ---- application state ------------------------------------------------------

interface AnalysisState {
  sheetMappings: Record<string, SheetMapping>;
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

    const newDocFile = $<HTMLInputElement>("newDocFile").files?.[0];
    if (!newDocFile) throw new Error("請先選擇新版文件");
    const prevDocFile = $<HTMLInputElement>("prevDocFile").files?.[0];
    const trackingFile = $<HTMLInputElement>("trackingFile").files?.[0];

    if (trackingFile && !$<HTMLDivElement>("versionMatchSection").hidden && resolvedVersionNo === null) {
      throw new Error("追蹤表找不到這個檔名，請先在上面選擇對應的既有文件，或勾選「確定是新的檔名」");
    }

    const sheetMappings = collectSheetMappingsFromUI();
    if (Object.values(sheetMappings).every((m) => m.zhColumns.length === 0)) {
      throw new Error("請至少為一個頁籤設定「中文」欄位");
    }
    saveSettings({ ...loadSettings(), sheetMappings });

    const adapter = new ExcelAdapter(ExcelJS, sheetMappings);
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

    state = { sheetMappings, newDocBytes, newDocWorkbook, isFirstVersion, records, translationItems };

    renderDiffTable(records);
    renderTranslationTable(translationItems);
    if (resolvedVersionNo === null) {
      $<HTMLInputElement>("versionNo").value = isFirstVersion ? "1" : $<HTMLInputElement>("versionNo").value || "";
    }

    $<HTMLDivElement>("resultsSection").hidden = false;
    setStatus(
      isFirstVersion
        ? `首次文件，無前版可比對，共 ${translationItems.length} 筆需要翻譯。`
        : `比對完成，共 ${translationItems.length} 筆新增/修改需要翻譯（可往下看每個頁籤的完整差異報告）。`
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
    const locTd = document.createElement("td");
    locTd.textContent = item.changeId.split("!")[0]; // sheet name, for context
    const fieldTd = document.createElement("td");
    fieldTd.textContent = item.fieldName ?? "";
    const zhTd = document.createElement("td");
    zhTd.textContent = item.zhText;
    const enTd = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    input.dataset.changeId = item.changeId;
    input.className = "translation-input";
    input.value = item.oldEn ?? "";
    enTd.appendChild(input);
    tr.append(locTd, fieldTd, zhTd, enTd);
    tbody.appendChild(tr);
  }
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

    const adapter = new ExcelAdapter(ExcelJS, state.sheetMappings);
    adapter.applyTranslations(state.newDocWorkbook, updates);
    const bilingualBytes = await adapter.save(state.newDocWorkbook);

    const versionNo = Number($<HTMLInputElement>("versionNo").value);
    if (!versionNo || versionNo < 1) throw new Error("請輸入正確的版本編號");

    const newDocFile = $<HTMLInputElement>("newDocFile").files?.[0];
    const dateStamp = todayCompact();
    const entries: ZipEntry[] = [{ name: "bilingual.xlsx", bytes: bilingualBytes }];

    if (!state.isFirstVersion) {
      const diffBytes = await buildDiffReportWorkbook(ExcelJS, state.records);
      entries.push({ name: "diff_report.xlsx", bytes: diffBytes });
    }

    const emlFile = $<HTMLInputElement>("emlFile").files?.[0];
    if (emlFile) entries.push({ name: emlFile.name, bytes: await readFileAsArrayBuffer(emlFile) });

    const trackingFile = $<HTMLInputElement>("trackingFile").files?.[0];
    if (trackingFile) {
      const mapping = readTrackingSettingsForm();
      const trackingBytes = await readFileAsArrayBuffer(trackingFile);
      const updatedTracking = await upsertRow(ExcelJS, trackingBytes, mapping, versionNo, {
        versionNo,
        receivedDate: $<HTMLInputElement>("receivedDate").value,
        sender: $<HTMLInputElement>("sender").value,
        sourceFileName: newDocFile?.name ?? "",
        emailPath: emlFile?.name ?? "",
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

    const mapping = readTrackingSettingsForm();
    const bytes = await readFileAsArrayBuffer(trackingFile);
    const current = (await readRow(ExcelJS, bytes, mapping, versionNo)) ?? {};

    const other = party === "overseas" ? "customerConfirmStatus" : "overseasConfirmStatus";
    const updates: Record<string, string> = {
      [`${party}ConfirmStatus`]: "已確認",
      [`${party}ConfirmDate`]: date,
      [`${party}ConfirmMethod`]: method,
    };
    if (current[other] === "已確認") updates.status = "已完成";

    const updated = await upsertRow(ExcelJS, bytes, mapping, versionNo, updates);
    downloadBlob(new Blob([updated]), trackingFile.name);
    setStatus(`已更新版本 ${versionNo} 的 ${party === "overseas" ? "海外團隊" : "客戶"} 確認狀態，請覆蓋回原檔案。`);
  } catch (err) {
    setStatus(`發生錯誤：${(err as Error).message}`, true);
  }
}

// ---- wiring -------------------------------------------------------------------

function init(): void {
  fillTrackingSettingsForm(loadSettings());
  $<HTMLButtonElement>("saveSettingsBtn").addEventListener("click", () => {
    saveSettings({ ...loadSettings(), trackingSheetMapping: readTrackingSettingsForm() });
    setStatus("設定已儲存。");
  });

  wireDropZone("newDocDropZone", "newDocFile");
  wireDropZone("prevDocDropZone", "prevDocFile");
  wireDropZone("emlDropZone", "emlFile");
  wireDropZone("trackingDropZone", "trackingFile");

  $<HTMLInputElement>("newDocFile").addEventListener("change", () => void handleNewDocSelected());
  $<HTMLInputElement>("trackingFile").addEventListener("change", () => void checkVersionMatch());
  $<HTMLInputElement>("emlFile").addEventListener("change", () => void handleEmlSelected());
  wireVersionMatchControls();

  $<HTMLButtonElement>("analyzeBtn").addEventListener("click", () => void handleAnalyze());
  $<HTMLButtonElement>("generateBtn").addEventListener("click", () => void handleGenerate());
  $<HTMLButtonElement>("confirmBtn").addEventListener("click", () => void handleConfirm());
}

document.addEventListener("DOMContentLoaded", init);
