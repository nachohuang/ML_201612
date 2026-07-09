/** Page wiring for the single-page tool. Everything happens in one browser session
 * (select files -> review/translate -> download) so there's no need for the
 * process/apply-translations two-step state file the Python CLI needed to bridge two
 * separate process invocations — see plan Context.
 *
 * The page has two tabs: "新舊文件上傳及比對" (batch document upload/compare/translate,
 * one card per new document — several requirement documents can arrive in the same
 * customer email) and "追蹤表管理" (import-or-create the tracking sheet and edit it
 * directly on screen). Each document card is built entirely in JS rather than static
 * HTML, since the number of cards is dynamic (however many files were selected).
 */

import type ExcelJSNS from "exceljs";
import type JSZipNS from "jszip";

import { alignAndDiff } from "./align.ts";
import { DEFAULT_SETTINGS, loadSettings, saveSettings, TRACKING_FIELD_LABELS } from "./config.ts";
import type { Settings, SheetMapping, TrackingColumns, TrackingSheetMapping, WordMapping } from "./config.ts";
import { buildCombinedReportWorkbook } from "./diff-report.ts";
import type { ReportRow } from "./diff-report.ts";
import { parseEmlFile } from "./eml-parser.ts";
import { ExcelAdapter, inspectWorkbook, resolveWorksheet } from "./excel-adapter.ts";
import type { InspectSheet } from "./excel-adapter.ts";
import { diffFreeformSheet } from "./freeform-diff.ts";
import { ChangeType } from "./models.ts";
import type { ChangeRecord, Segment, SegmentUpdate } from "./models.ts";
import { buildTrackingWorkbook, listAllRows, listEntries, upsertRow } from "./tracking-sheet.ts";
import type { TrackingRow } from "./tracking-sheet.ts";
import { WordAdapter } from "./word-adapter.ts";
import type { WordDoc } from "./word-adapter.ts";
import { diffFreeformWordDocument } from "./word-freeform-diff.ts";
import { buildZip } from "./zip-bundle.ts";
import type { ZipEntry } from "./zip-bundle.ts";

declare const ExcelJS: typeof ExcelJSNS;
declare const JSZip: typeof JSZipNS;
// DOMParser/XMLSerializer are native browser globals (unlike ExcelJS/JSZip, nothing to
// load via a <script> tag) — lib.dom.d.ts already declares them for TypeScript.

// ---- small DOM helpers ----------------------------------------------------

function $<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`missing element #${id}`);
  return el as T;
}

function setStatus(message: string, isError = false): void {
  const el = $<HTMLDivElement>("status");
  el.textContent = message;
  el.classList.toggle("error", isError);
}

function setTrackingStatus(message: string, isError = false): void {
  const el = $<HTMLDivElement>("trackingStatus");
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

function getDocType(file: File): "excel" | "word" {
  return file.name.toLowerCase().endsWith(".docx") ? "word" : "excel";
}

function splitColumns(value: string): string[] {
  return value
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
}

function sanitizeBaseName(fileName: string): string {
  const base = fileName.replace(/\.[^.]+$/, "");
  return base.replace(/[^\w一-鿿-]+/g, "_") || "file";
}

/** Wires a drop zone element so dragging a file onto it (or clicking it) populates the
 * given file input — the native "Choose File" button inside keeps working as a fallback. */
function wireDropZoneEl(zone: HTMLElement, input: HTMLInputElement): void {
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

// ---- tab navigation ----------------------------------------------------------------

function switchTab(tab: "docs" | "tracking"): void {
  $<HTMLElement>("docsPage").hidden = tab !== "docs";
  $<HTMLElement>("trackingPage").hidden = tab !== "tracking";
  $<HTMLButtonElement>("tabBtnDocs").classList.toggle("active", tab === "docs");
  $<HTMLButtonElement>("tabBtnTracking").classList.toggle("active", tab === "tracking");
}

// ---- tracking sheet column-mapping settings (fixed set of fields) -----------------

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

// ---- Word mapping form component (built per document, not a static form) ----------

interface WordMappingSection {
  root: HTMLElement;
  getMapping(): WordMapping;
}

function buildWordMappingSection(initial: WordMapping): WordMappingSection {
  const root = document.createElement("div");

  const modeRow = document.createElement("div");
  modeRow.className = "field-row";
  const modeLabel = document.createElement("label");
  modeLabel.textContent = "比對方式";
  const modeSelect = document.createElement("select");
  modeSelect.className = "word-mode";
  modeSelect.innerHTML =
    '<option value="freeform">自由格式（整篇比對，中英文已混在同段落/儲存格內，不需設定欄位）</option>' +
    '<option value="alternating_paragraphs">交替段落（一段中文緊接著一段對應英文）</option>' +
    '<option value="same_paragraph_split">同段落切分（同一段落用分隔符號隔開中英文）</option>' +
    '<option value="table_based">表格式（表格中有專門的中文欄／英文欄）</option>';
  modeSelect.value = initial.mode;
  modeLabel.appendChild(modeSelect);
  modeRow.appendChild(modeLabel);
  root.appendChild(modeRow);

  const delimiterRow = document.createElement("div");
  delimiterRow.className = "field-row";
  const delimiterLabel = document.createElement("label");
  delimiterLabel.textContent = "分隔符號";
  const delimiterInput = document.createElement("input");
  delimiterInput.type = "text";
  delimiterInput.placeholder = "例如 | 或 -";
  delimiterInput.value = initial.splitDelimiter ?? "";
  delimiterLabel.appendChild(delimiterInput);
  delimiterRow.appendChild(delimiterLabel);
  root.appendChild(delimiterRow);

  const tableColumnsRow = document.createElement("div");
  tableColumnsRow.className = "field-row";
  const zhLabel = document.createElement("label");
  zhLabel.textContent = "中文欄位（0為第一欄，逗號分隔）";
  const zhInput = document.createElement("input");
  zhInput.type = "text";
  zhInput.placeholder = "例如 0";
  zhInput.value = initial.zhColumns.join(",");
  zhLabel.appendChild(zhInput);
  const enLabel = document.createElement("label");
  enLabel.textContent = "英文欄位";
  const enInput = document.createElement("input");
  enInput.type = "text";
  enInput.placeholder = "例如 1";
  enInput.value = initial.enColumns.join(",");
  enLabel.appendChild(enInput);
  tableColumnsRow.append(zhLabel, enLabel);
  root.appendChild(tableColumnsRow);

  const hint = document.createElement("p");
  hint.className = "hint";
  hint.textContent = "大部分需求文件的中英文是混在同一段落/儲存格內，沒有清楚分開，這種情況請用「自由格式」。";
  root.appendChild(hint);

  function updateVisibility(): void {
    delimiterRow.hidden = modeSelect.value !== "same_paragraph_split";
    tableColumnsRow.hidden = modeSelect.value !== "table_based";
  }
  modeSelect.addEventListener("change", updateVisibility);
  updateVisibility();

  function getMapping(): WordMapping {
    return {
      mode: modeSelect.value as WordMapping["mode"],
      splitDelimiter: delimiterInput.value || null,
      zhColumns: splitColumns(zhInput.value),
      enColumns: splitColumns(enInput.value),
    };
  }

  return { root, getMapping };
}

// ---- Excel per-sheet mapping component (point-and-click column roles) -------------

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

interface SheetMappingSection {
  root: HTMLElement;
  getMappings(): Record<string, SheetMapping>;
}

interface SheetCardRef {
  sheetName: string;
  enableCheckbox: HTMLInputElement;
  modeSelect: HTMLSelectElement;
  headerRowInput: HTMLInputElement;
  roleSelects: HTMLSelectElement[];
}

function buildSheetMappingSection(sheets: InspectSheet[], existing: Record<string, SheetMapping>): SheetMappingSection {
  const root = document.createElement("div");
  const refs: SheetCardRef[] = [];

  for (const sheet of sheets) {
    const priorMapping = existing[sheet.sheetName];
    // Every tab is included by default now (screen mockups/report layouts are
    // meaningful requirement content too, not just zh/en tables — see plan Context),
    // defaulting to freeform (whole-cell compare, nothing to configure) unless the
    // user previously set this sheet up as a real zh/en table.
    const mode: "table" | "freeform" = priorMapping?.mode === "table" ? "table" : "freeform";

    const card = document.createElement("div");
    card.className = "sheet-mapping-card";
    card.dataset.sheetName = sheet.sheetName;

    const heading = document.createElement("label");
    heading.className = "sheet-card-heading";
    const enableCheckbox = document.createElement("input");
    enableCheckbox.type = "checkbox";
    enableCheckbox.className = "sheet-enable";
    enableCheckbox.checked = true;
    heading.appendChild(enableCheckbox);
    heading.append(` 比對此頁籤：${sheet.sheetName} （${sheet.columns.length} 欄）`);
    card.appendChild(heading);

    const detail = document.createElement("div");
    detail.className = "sheet-card-detail";
    enableCheckbox.addEventListener("change", () => {
      detail.hidden = !enableCheckbox.checked;
    });

    const modeLabel = document.createElement("label");
    modeLabel.textContent = "比對方式：";
    const modeSelect = document.createElement("select");
    modeSelect.className = "sheet-mode";
    modeSelect.innerHTML =
      '<option value="freeform">自由格式（整格比對，畫面/報表版面，不用設定欄位）</option>' +
      '<option value="table">表格（設定中英文欄位，用於清單/列表類頁籤）</option>';
    modeSelect.value = mode;
    modeLabel.appendChild(modeSelect);
    detail.appendChild(modeLabel);

    const tableDetail = document.createElement("div");
    tableDetail.className = "sheet-table-detail";
    tableDetail.hidden = mode !== "table";
    modeSelect.addEventListener("change", () => {
      tableDetail.hidden = modeSelect.value !== "table";
    });

    const headerRowLabel = document.createElement("label");
    headerRowLabel.textContent = "表頭列號：";
    const headerRowInput = document.createElement("input");
    headerRowInput.type = "number";
    headerRowInput.min = "1";
    headerRowInput.value = String(priorMapping?.headerRow ?? 1);
    headerRowInput.className = "sheet-header-row";
    headerRowInput.style.width = "4em";
    headerRowLabel.appendChild(headerRowInput);
    tableDetail.appendChild(headerRowLabel);

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
    const roleSelects: HTMLSelectElement[] = [];
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
      select.value = roleForColumn(col.column, priorMapping);
      td.appendChild(select);
      roleRow.appendChild(td);
      roleSelects.push(select);
    }
    table.appendChild(roleRow);

    tableDetail.appendChild(table);
    detail.appendChild(tableDetail);
    card.appendChild(detail);
    root.appendChild(card);

    refs.push({ sheetName: sheet.sheetName, enableCheckbox, modeSelect, headerRowInput, roleSelects });
  }

  function getMappings(): Record<string, SheetMapping> {
    const result: Record<string, SheetMapping> = {};
    for (const ref of refs) {
      if (!ref.enableCheckbox.checked) continue;
      const mode = ref.modeSelect.value as "table" | "freeform";

      if (mode === "freeform") {
        result[ref.sheetName] = { sheetName: ref.sheetName, mode, headerRow: 1, keyColumns: [], zhColumns: [], enColumns: [] };
        continue;
      }

      const headerRow = Number(ref.headerRowInput.value) || 1;
      const keyColumns: string[] = [];
      const pairs: Record<string, { zh?: string; en?: string }> = {};

      for (const select of ref.roleSelects) {
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

      result[ref.sheetName] = { sheetName: ref.sheetName, mode, headerRow, keyColumns, zhColumns, enColumns };
    }
    return result;
  }

  return { root, getMappings };
}

// ---- per-document batch state ------------------------------------------------------

type TranslationItems = Array<{ changeId: string; zhText: string; fieldName: string | null; oldEn?: string }>;

interface DocItem {
  id: number;
  file: File;
  docType: "excel" | "word";
  root: HTMLElement;
  prevFileInput: HTMLInputElement;
  versionNoInput: HTMLInputElement;
  senderInput: HTMLInputElement;
  receivedDateInput: HTMLInputElement;
  versionMatchSection: HTMLElement;
  versionMatchInfo: HTMLElement;
  matchExistingSelect: HTMLSelectElement;
  confirmNewFileCheckbox: HTMLInputElement;
  analyzeBtn: HTMLButtonElement;
  statusEl: HTMLElement;
  diffSection: HTMLElement;
  diffTableBody: HTMLTableSectionElement;
  translationSection: HTMLElement;
  translationTableBody: HTMLTableSectionElement;
  resolvedVersionNo: number | null;
  versionMatchRequestId: number;
  /** Resolves once the mapping UI (sheet cards for Excel, built asynchronously since it
   * needs to inspect the workbook; the Word form, built synchronously) is ready. */
  mappingReady: Promise<void>;
  sheetMapping?: SheetMappingSection;
  wordMapping?: WordMappingSection;
  analyzed?: {
    isFirstVersion: boolean;
    records: ChangeRecord[];
    excel?: { sheetMappings: Record<string, SheetMapping>; workbook: ExcelJSNS.Workbook };
    word?: { mapping: WordMapping; doc: WordDoc };
  };
}

let docItems: DocItem[] = [];
let nextItemId = 0;
let lastEmlMeta: { sender?: string | null; receivedDate?: string | null } | null = null;

function segmentsAsAdded(segments: Segment[]): ChangeRecord[] {
  return segments
    .filter((s) => s.zhText.trim() !== "")
    .map((s) => ({
      locationId: s.locationId,
      fieldName: s.fieldName,
      changeType: ChangeType.ADDED,
      newZh: s.zhText,
      newEnLocationId: s.enLocationId,
    }));
}

function buildTranslationItems(records: ChangeRecord[]): TranslationItems {
  return records
    .filter((r) => (r.changeType === ChangeType.ADDED || r.changeType === ChangeType.MODIFIED) && r.newEnLocationId)
    .map((r) => ({
      changeId: r.newEnLocationId as string,
      zhText: r.newZh ?? "",
      fieldName: r.fieldName,
      // MODIFIED rows are pre-filled with the old translation as a starting point to
      // revise, mirroring how a human translator would work from the prior text.
      oldEn: r.changeType === ChangeType.MODIFIED ? r.oldEn : undefined,
    }));
}

function renderItemDiffTable(item: DocItem, records: ChangeRecord[]): void {
  const tbody = item.diffTableBody;
  tbody.innerHTML = "";
  const visible = records.filter((r) => r.changeType !== ChangeType.UNCHANGED);
  item.diffSection.hidden = visible.length === 0;

  for (const r of visible) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escapeHtml(r.locationId)}</td><td>${escapeHtml(r.fieldName ?? "")}</td><td>${escapeHtml(r.oldZh ?? "")}</td><td>${escapeHtml(r.newZh ?? "")}</td><td>${escapeHtml(r.changeType)}</td>`;
    const noteTd = document.createElement("td");
    const noteInput = document.createElement("input");
    noteInput.type = "text";
    noteInput.className = "note-input";
    noteInput.dataset.locationId = r.locationId;
    noteTd.appendChild(noteInput);
    tr.appendChild(noteTd);
    tbody.appendChild(tr);
  }
}

function renderItemTranslationTable(item: DocItem, items: TranslationItems): void {
  const tbody = item.translationTableBody;
  tbody.innerHTML = "";
  item.translationSection.hidden = items.length === 0;

  for (const it of items) {
    const tr = document.createElement("tr");
    const locTd = document.createElement("td");
    locTd.textContent = it.changeId;
    const fieldTd = document.createElement("td");
    fieldTd.textContent = it.fieldName ?? "";
    const zhTd = document.createElement("td");
    zhTd.textContent = it.zhText;
    const enTd = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    input.dataset.changeId = it.changeId;
    input.className = "translation-input";
    input.value = it.oldEn ?? "";
    enTd.appendChild(input);
    tr.append(locTd, fieldTd, zhTd, enTd);
    tbody.appendChild(tr);
  }
}

// ---- version matching by filename against the (shared) tracking sheet -------------

// buildDocItem() wires this per item, and the shared trackingFile input's own "change"
// listener re-checks every item — a slower earlier call could resolve after a faster
// later one and clobber a newer call's result, so each item guards with its own
// monotonic request id (mirrors the single-document tool's original race-condition fix).
async function checkVersionMatch(item: DocItem): Promise<void> {
  const requestId = ++item.versionMatchRequestId;
  const trackingFile = $<HTMLInputElement>("trackingFile").files?.[0];

  if (!trackingFile) {
    item.versionMatchSection.hidden = true;
    item.resolvedVersionNo = null;
    return;
  }

  let entries;
  try {
    const trackingBytes = await trackingFile.arrayBuffer();
    const mapping = readTrackingSettingsForm();
    entries = await listEntries(ExcelJS, trackingBytes, mapping);
  } catch {
    return; // superseded call whose file handle became unreadable — harmless, ignore
  }

  if (requestId !== item.versionMatchRequestId) return; // superseded by a newer call

  const exact = entries.find((e) => e.sourceFileName === item.file.name);
  if (exact) {
    item.resolvedVersionNo = exact.versionNo;
    item.versionMatchSection.hidden = true;
    item.versionNoInput.value = String(exact.versionNo);
    return;
  }

  if (entries.length === 0) {
    item.resolvedVersionNo = 1;
    item.versionMatchSection.hidden = true;
    item.versionNoInput.value = "1";
    return;
  }

  // No exact match but the tracking sheet has history — don't guess, ask.
  item.resolvedVersionNo = null;
  item.versionMatchSection.hidden = false;
  const select = item.matchExistingSelect;
  select.innerHTML = '<option value="">-- 請選擇這是延續哪一份既有文件 --</option>';
  for (const entry of entries) {
    const opt = document.createElement("option");
    opt.value = String(entry.versionNo);
    opt.textContent = `版本 ${entry.versionNo}：${entry.sourceFileName || "(無檔名記錄)"}`;
    select.appendChild(opt);
  }
  select.value = "";
  item.confirmNewFileCheckbox.checked = false;
  item.versionMatchInfo.textContent = `在追蹤表中找不到檔名「${item.file.name}」。`;
}

// ---- building one document's card --------------------------------------------------

function buildDocItem(file: File): DocItem {
  const id = nextItemId++;
  const docType = getDocType(file);

  const root = document.createElement("div");
  root.className = "doc-item-card";
  root.dataset.itemId = String(id);

  const heading = document.createElement("h3");
  heading.textContent = `${file.name}（${docType === "word" ? "Word 文件" : "Excel 活頁簿"}）`;
  root.appendChild(heading);

  const prevZone = document.createElement("div");
  prevZone.className = "drop-zone";
  const prevLabel = document.createElement("label");
  prevLabel.textContent = "前一版雙語檔（沒有就代表這是首次文件，可拖拉檔案到這裡）";
  const prevInput = document.createElement("input");
  prevInput.type = "file";
  prevInput.accept = ".xlsx,.docx";
  prevInput.className = "prev-file-input";
  prevLabel.appendChild(prevInput);
  prevZone.appendChild(prevLabel);
  root.appendChild(prevZone);
  wireDropZoneEl(prevZone, prevInput);

  const meta = document.createElement("div");
  meta.className = "doc-item-meta";

  const versionLabel = document.createElement("label");
  versionLabel.textContent = "版本編號";
  const versionInput = document.createElement("input");
  versionInput.type = "number";
  versionInput.min = "1";
  versionInput.style.width = "5em";
  versionInput.className = "item-version-no";
  versionLabel.appendChild(versionInput);
  meta.appendChild(versionLabel);

  const senderLabel = document.createElement("label");
  senderLabel.textContent = "寄件窗口";
  const senderInput = document.createElement("input");
  senderInput.type = "text";
  senderInput.className = "item-sender";
  senderLabel.appendChild(senderInput);
  meta.appendChild(senderLabel);

  const dateLabel = document.createElement("label");
  dateLabel.textContent = "收到日期";
  const dateInput = document.createElement("input");
  dateInput.type = "date";
  dateInput.className = "item-received-date";
  dateLabel.appendChild(dateInput);
  meta.appendChild(dateLabel);
  root.appendChild(meta);

  if (lastEmlMeta?.sender) senderInput.value = lastEmlMeta.sender;
  if (lastEmlMeta?.receivedDate) dateInput.value = lastEmlMeta.receivedDate;

  const versionMatchSection = document.createElement("div");
  versionMatchSection.className = "version-match-section";
  versionMatchSection.hidden = true;
  const versionMatchInfo = document.createElement("p");
  versionMatchSection.appendChild(versionMatchInfo);
  const matchRow = document.createElement("div");
  matchRow.className = "field-row";
  const matchLabel = document.createElement("label");
  matchLabel.textContent = "對應到既有文件";
  const matchSelect = document.createElement("select");
  matchSelect.className = "match-existing-select";
  matchLabel.appendChild(matchSelect);
  matchRow.appendChild(matchLabel);
  versionMatchSection.appendChild(matchRow);
  const confirmRow = document.createElement("div");
  confirmRow.className = "field-row";
  const confirmLabel = document.createElement("label");
  const confirmCheckbox = document.createElement("input");
  confirmCheckbox.type = "checkbox";
  confirmCheckbox.className = "confirm-new-file-checkbox";
  confirmLabel.appendChild(confirmCheckbox);
  confirmLabel.append(" 確定這是全新的文件（沒有前一版）");
  confirmRow.appendChild(confirmLabel);
  versionMatchSection.appendChild(confirmRow);
  root.appendChild(versionMatchSection);

  const mappingContainer = document.createElement("div");
  root.appendChild(mappingContainer);

  const analyzeBtn = document.createElement("button");
  analyzeBtn.type = "button";
  analyzeBtn.textContent = "分析此文件";
  analyzeBtn.className = "analyze-btn";
  root.appendChild(analyzeBtn);

  const statusEl = document.createElement("div");
  statusEl.className = "item-status";
  root.appendChild(statusEl);

  const diffSection = document.createElement("div");
  diffSection.className = "doc-item-diff";
  diffSection.hidden = true;
  const diffHeading = document.createElement("h4");
  diffHeading.textContent = "差異報告";
  diffSection.appendChild(diffHeading);
  const diffTable = document.createElement("table");
  diffTable.innerHTML =
    '<thead><tr><th>位置</th><th>欄位名稱</th><th>舊內容</th><th>新內容</th><th>變動類型</th><th>備註</th></tr></thead><tbody class="diff-table-body"></tbody>';
  diffSection.appendChild(diffTable);
  root.appendChild(diffSection);

  const translationSection = document.createElement("div");
  translationSection.className = "doc-item-translation";
  translationSection.hidden = true;
  const translationHeading = document.createElement("h4");
  translationHeading.textContent = "填寫英文翻譯";
  translationSection.appendChild(translationHeading);
  const translationTable = document.createElement("table");
  translationTable.innerHTML =
    '<thead><tr><th>位置</th><th>欄位名稱</th><th>中文原文</th><th>英文翻譯</th></tr></thead><tbody class="translation-table-body"></tbody>';
  translationSection.appendChild(translationTable);
  const translationHint = document.createElement("p");
  translationHint.className = "hint";
  translationHint.textContent = "修改項目已預先帶入舊翻譯供你參考修改；新增項目請自行填寫。";
  translationSection.appendChild(translationHint);
  root.appendChild(translationSection);

  const item: DocItem = {
    id,
    file,
    docType,
    root,
    prevFileInput: prevInput,
    versionNoInput: versionInput,
    senderInput,
    receivedDateInput: dateInput,
    versionMatchSection,
    versionMatchInfo,
    matchExistingSelect: matchSelect,
    confirmNewFileCheckbox: confirmCheckbox,
    analyzeBtn,
    statusEl,
    diffSection,
    diffTableBody: diffTable.querySelector("tbody") as HTMLTableSectionElement,
    translationSection,
    translationTableBody: translationTable.querySelector("tbody") as HTMLTableSectionElement,
    resolvedVersionNo: null,
    versionMatchRequestId: 0,
    mappingReady: Promise.resolve(),
  };

  prevInput.addEventListener("change", () => void checkVersionMatch(item));
  matchSelect.addEventListener("change", () => {
    if (matchSelect.value) {
      confirmCheckbox.checked = false;
      const matchedVersion = Number(matchSelect.value);
      item.resolvedVersionNo = matchedVersion + 1;
      versionInput.value = String(item.resolvedVersionNo);
    } else {
      item.resolvedVersionNo = null;
    }
  });
  confirmCheckbox.addEventListener("change", () => {
    if (confirmCheckbox.checked) {
      matchSelect.value = "";
      item.resolvedVersionNo = 1;
      versionInput.value = "1";
    } else {
      item.resolvedVersionNo = null;
    }
  });
  analyzeBtn.addEventListener("click", () => void analyzeItem(item));

  if (docType === "word") {
    const wordSection = buildWordMappingSection(loadSettings().wordMapping);
    item.wordMapping = wordSection;
    mappingContainer.appendChild(wordSection.root);
  } else {
    const loadingMsg = document.createElement("p");
    loadingMsg.className = "hint";
    loadingMsg.textContent = "讀取工作表中…";
    mappingContainer.appendChild(loadingMsg);

    item.mappingReady = (async () => {
      try {
        const bytes = await file.arrayBuffer();
        const sheets = await inspectWorkbook(ExcelJS, bytes);
        mappingContainer.innerHTML = "";
        const mappingHint = document.createElement("p");
        mappingHint.className = "hint";
        mappingHint.textContent = "每個頁籤（工作表）結構可能不同，請分別點選每欄的用途；設定會依頁籤名稱自動記住，下次上傳同樣頁籤不用重填。";
        mappingContainer.appendChild(mappingHint);
        const section = buildSheetMappingSection(sheets, loadSettings().sheetMappings);
        item.sheetMapping = section;
        mappingContainer.appendChild(section.root);
      } catch (err) {
        mappingContainer.innerHTML = "";
        const errEl = document.createElement("p");
        errEl.className = "item-status error";
        errEl.textContent = `讀取文件失敗：${(err as Error).message}`;
        mappingContainer.appendChild(errEl);
      }
    })();
  }

  void checkVersionMatch(item);

  return item;
}

// ---- new-document batch selection --------------------------------------------------

function handleNewDocFilesSelected(): void {
  const files = Array.from($<HTMLInputElement>("newDocFiles").files ?? []);

  const listEl = $<HTMLUListElement>("newDocFileList");
  listEl.innerHTML = "";
  for (const f of files) {
    const li = document.createElement("li");
    li.textContent = `📄 ${f.name}`;
    listEl.appendChild(li);
  }

  const container = $<HTMLDivElement>("docItemsContainer");
  container.innerHTML = "";
  docItems = files.map((f) => buildDocItem(f));
  for (const item of docItems) container.appendChild(item.root);
}

// ---- eml parsing (shared across the whole batch — one email, several attachments) -

async function handleEmlSelected(): Promise<void> {
  const file = $<HTMLInputElement>("emlFile").files?.[0];
  if (!file) return;
  try {
    const meta = await parseEmlFile(file);
    lastEmlMeta = meta;
    for (const item of docItems) {
      if (meta.sender) item.senderInput.value = meta.sender;
      if (meta.receivedDate) item.receivedDateInput.value = meta.receivedDate;
    }
    setStatus(
      meta.sender || meta.receivedDate
        ? "已從信件自動帶入寄件者/日期到所有文件，請確認正確，需要的話可自行修改。"
        : "無法從這份信件解析出寄件者/日期，請手動填寫。"
    );
  } catch {
    setStatus("無法解析這份 .eml 信件，請手動填寫寄件者/日期。", true);
  }
}

// ---- analyze one document ----------------------------------------------------------

async function analyzeItem(item: DocItem): Promise<void> {
  try {
    item.statusEl.textContent = "分析中…";
    item.statusEl.classList.remove("error");
    await item.mappingReady;

    const prevFile = item.prevFileInput.files?.[0];
    const isFirstVersion = !prevFile;
    let records: ChangeRecord[];
    let excelState: { sheetMappings: Record<string, SheetMapping>; workbook: ExcelJSNS.Workbook } | undefined;
    let wordState: { mapping: WordMapping; doc: WordDoc } | undefined;

    if (item.docType === "word") {
      if (!item.wordMapping) throw new Error("Word 設定尚未就緒");
      const mapping = item.wordMapping.getMapping();
      saveSettings({ ...loadSettings(), wordMapping: mapping });

      const adapter = new WordAdapter(JSZip, DOMParser, XMLSerializer, mapping);
      const newDoc = await adapter.load(await item.file.arrayBuffer());

      if (mapping.mode === "freeform") {
        const prevDom = prevFile ? (await adapter.load(await prevFile.arrayBuffer())).dom : undefined;
        records = diffFreeformWordDocument(prevDom, newDoc.dom);
      } else {
        const newSegments = adapter.extractSegments(newDoc);
        if (isFirstVersion) {
          records = segmentsAsAdded(newSegments);
        } else {
          const prevDoc = await adapter.load(await prevFile.arrayBuffer());
          const prevSegments = adapter.extractSegments(prevDoc);
          records = alignAndDiff(prevSegments, newSegments);
        }
      }
      wordState = { mapping, doc: newDoc };
    } else {
      if (!item.sheetMapping) throw new Error("頁籤設定尚未就緒");
      const sheetMappings = item.sheetMapping.getMappings();
      if (Object.keys(sheetMappings).length === 0) throw new Error("請至少勾選一個頁籤進行比對");
      saveSettings({ ...loadSettings(), sheetMappings });

      const adapter = new ExcelAdapter(ExcelJS, sheetMappings);
      const newWorkbook = await adapter.load(await item.file.arrayBuffer());
      const newSegments = adapter.extractSegments(newWorkbook);
      const freeformMappings = Object.values(sheetMappings).filter((m) => m.mode === "freeform");

      let prevWorkbook: ExcelJSNS.Workbook | undefined;
      if (isFirstVersion) {
        records = segmentsAsAdded(newSegments);
      } else {
        prevWorkbook = await adapter.load(await prevFile.arrayBuffer());
        const prevSegments = adapter.extractSegments(prevWorkbook, newWorkbook);
        records = alignAndDiff(prevSegments, newSegments);
      }

      for (const mapping of freeformMappings) {
        const newWs = resolveWorksheet(newWorkbook, mapping.sheetName);
        if (!newWs) continue;
        const oldWs = prevWorkbook ? resolveWorksheet(prevWorkbook, mapping.sheetName, newWorkbook) : undefined;
        records.push(...diffFreeformSheet(oldWs, newWs));
      }

      excelState = { sheetMappings, workbook: newWorkbook };
    }

    const translationItems = buildTranslationItems(records);
    item.analyzed = { isFirstVersion, records, excel: excelState, word: wordState };
    renderItemDiffTable(item, records);
    renderItemTranslationTable(item, translationItems);
    if (item.resolvedVersionNo !== null) item.versionNoInput.value = String(item.resolvedVersionNo);
    else if (!item.versionNoInput.value) item.versionNoInput.value = isFirstVersion ? "1" : "";

    item.statusEl.textContent = isFirstVersion
      ? `首次文件，無前版可比對，共 ${translationItems.length} 筆需要翻譯。`
      : `比對完成，共 ${translationItems.length} 筆新增/修改需要翻譯（可往下看完整差異報告）。`;
  } catch (err) {
    item.statusEl.textContent = `發生錯誤：${(err as Error).message}`;
    item.statusEl.classList.add("error");
  }
}

// ---- generate + download the whole batch -------------------------------------------

function buildReadme(hasTracking: boolean): string {
  const lines = [
    "本次批次產出使用說明：",
    "",
    "1. 每份文件對應一個 bilingual_<檔名>.xlsx/docx，連同 diff_report.xlsx 一起放進 archive/ 對應版本的資料夾",
    "2. 同樣的 bilingual 檔案也複製一份到 output/ 資料夾，供上傳 Teams/SharePoint",
  ];
  if (hasTracking) lines.push("3. 把追蹤表檔案覆蓋回原本存放的位置（取代舊檔）");
  return lines.join("\n");
}

async function handleGenerateAll(): Promise<void> {
  try {
    if (docItems.length === 0) throw new Error("請先選擇至少一份新版文件");

    const notAnalyzed = docItems.find((i) => !i.analyzed);
    if (notAnalyzed) throw new Error(`「${notAnalyzed.file.name}」尚未分析，請先按「分析此文件」`);

    for (const item of docItems) {
      if (!item.versionMatchSection.hidden && item.resolvedVersionNo === null) {
        throw new Error(`「${item.file.name}」追蹤表找不到這個檔名，請先選擇對應的既有文件，或勾選「確定是新的檔名」`);
      }
      const inputs = Array.from(item.translationTableBody.querySelectorAll<HTMLInputElement>("input[data-change-id]"));
      const empty = inputs.filter((i) => i.value.trim() === "");
      if (empty.length > 0) throw new Error(`「${item.file.name}」還有 ${empty.length} 筆尚未填寫英文翻譯`);
      const versionNo = Number(item.versionNoInput.value);
      if (!versionNo || versionNo < 1) throw new Error(`「${item.file.name}」請輸入正確的版本編號`);
    }

    setStatus("產生歸檔中…");

    const dateStamp = todayCompact();
    const entries: ZipEntry[] = [];
    const reportRows: ReportRow[] = [];

    const trackingFile = $<HTMLInputElement>("trackingFile").files?.[0];
    let trackingBytes = trackingFile ? await trackingFile.arrayBuffer() : undefined;
    const trackingMapping = readTrackingSettingsForm();

    for (const item of docItems) {
      const { isFirstVersion, records, excel, word } = item.analyzed!;
      const inputs = Array.from(item.translationTableBody.querySelectorAll<HTMLInputElement>("input[data-change-id]"));
      const updates: SegmentUpdate[] = inputs.map((i) => ({ enLocationId: i.dataset.changeId as string, enText: i.value.trim() }));
      const translationByChangeId = new Map(inputs.map((i) => [i.dataset.changeId as string, i.value.trim()]));

      const noteInputs = Array.from(item.diffTableBody.querySelectorAll<HTMLInputElement>("input.note-input"));
      const notesByLocation = new Map(noteInputs.map((i) => [i.dataset.locationId as string, i.value.trim()]));

      const ext = item.docType === "word" ? "docx" : "xlsx";
      const bilingualEntryName = `bilingual_${sanitizeBaseName(item.file.name)}.${ext}`;

      let bilingualBytes: ArrayBuffer;
      if (excel) {
        const adapter = new ExcelAdapter(ExcelJS, excel.sheetMappings);
        adapter.applyTranslations(excel.workbook, updates);
        bilingualBytes = await adapter.save(excel.workbook);
      } else if (word) {
        const adapter = new WordAdapter(JSZip, DOMParser, XMLSerializer, word.mapping);
        adapter.applyTranslations(word.doc, updates);
        bilingualBytes = await adapter.save(word.doc);
      } else {
        throw new Error("內部錯誤：找不到已分析的文件狀態");
      }
      entries.push({ name: bilingualEntryName, bytes: bilingualBytes });

      for (const r of records) {
        if (r.changeType === ChangeType.UNCHANGED) continue;
        reportRows.push({
          docName: item.file.name,
          locationId: r.locationId,
          fieldName: r.fieldName,
          oldZh: r.oldZh,
          newZh: r.newZh,
          changeType: r.changeType,
          note: notesByLocation.get(r.locationId) ?? "",
          translation: r.newEnLocationId ? translationByChangeId.get(r.newEnLocationId) ?? "" : "",
        });
      }

      const versionNo = Number(item.versionNoInput.value);
      const folder = `v${String(versionNo).padStart(2, "0")}_${dateStamp}`;

      if (trackingBytes) {
        trackingBytes = await upsertRow(ExcelJS, trackingBytes, trackingMapping, versionNo, {
          versionNo,
          receivedDate: item.receivedDateInput.value,
          sender: item.senderInput.value,
          sourceFileName: item.file.name,
          bilingualFilePath: `${folder}/${bilingualEntryName}`,
          diffReportPath: `${folder}/diff_report.xlsx`,
          updateSummary: isFirstVersion ? "首次提供文件" : "",
          isFirstVersion: isFirstVersion ? "是" : "否",
          overseasConfirmStatus: "未確認",
          customerConfirmStatus: "未確認",
          status: "待確認",
        });
      }
    }

    if (reportRows.length > 0) {
      const reportBytes = await buildCombinedReportWorkbook(ExcelJS, reportRows);
      entries.push({ name: "diff_report.xlsx", bytes: reportBytes });
    }

    const emlFile = $<HTMLInputElement>("emlFile").files?.[0];
    if (emlFile) entries.push({ name: emlFile.name, bytes: await emlFile.arrayBuffer() });

    if (trackingFile && trackingBytes) entries.push({ name: trackingFile.name, bytes: trackingBytes });

    entries.push({
      name: "使用說明.txt",
      bytes: new TextEncoder().encode(buildReadme(!!trackingFile)).buffer as ArrayBuffer,
    });

    const zipBlob = await buildZip(JSZip, entries);
    downloadBlob(zipBlob, `批次歸檔_${dateStamp}.zip`);
    setStatus(`已產生歸檔 zip（共 ${docItems.length} 份文件），請解壓縮後依「使用說明.txt」放到正確資料夾。`);
  } catch (err) {
    setStatus(`發生錯誤：${(err as Error).message}`, true);
  }
}

// ---- tracking sheet management tab -------------------------------------------------

let trackingRows: TrackingRow[] = [];
let trackingImportedBytes: ArrayBuffer | undefined;

function blankTrackingRow(): TrackingRow {
  const row = {} as TrackingRow;
  for (const field of Object.keys(TRACKING_FIELD_LABELS) as Array<keyof TrackingColumns>) row[field] = "";
  return row;
}

function buildTrackingRowEl(row: TrackingRow): HTMLTableRowElement {
  const tr = document.createElement("tr");
  for (const field of Object.keys(TRACKING_FIELD_LABELS) as Array<keyof TrackingColumns>) {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = field.toLowerCase().includes("date") ? "date" : "text";
    input.className = "track-input";
    input.dataset.field = field;
    input.value = row[field] ?? "";
    td.appendChild(input);
    tr.appendChild(td);
  }
  return tr;
}

function renderTrackingTable(): void {
  const headRow = $<HTMLTableRowElement>("trackingTableHeadRow");
  headRow.innerHTML = Object.values(TRACKING_FIELD_LABELS)
    .map((label) => `<th>${escapeHtml(label)}</th>`)
    .join("");

  const tbody = $<HTMLTableSectionElement>("trackingTableBody");
  tbody.innerHTML = "";
  for (const row of trackingRows) tbody.appendChild(buildTrackingRowEl(row));

  $<HTMLDivElement>("trackingTableSection").hidden = false;
}

function collectTrackingRowsFromUI(): TrackingRow[] {
  const tbody = $<HTMLTableSectionElement>("trackingTableBody");
  const rows: TrackingRow[] = [];
  for (const tr of Array.from(tbody.querySelectorAll("tr"))) {
    const row = {} as TrackingRow;
    for (const input of Array.from(tr.querySelectorAll<HTMLInputElement>("input[data-field]"))) {
      row[input.dataset.field as keyof TrackingColumns] = input.value.trim();
    }
    rows.push(row);
  }
  return rows;
}

async function handleTrackingMgmtFileSelected(): Promise<void> {
  const file = $<HTMLInputElement>("trackingMgmtFile").files?.[0];
  if (!file) return;
  try {
    const bytes = await file.arrayBuffer();
    const mapping = readTrackingSettingsForm();
    trackingRows = await listAllRows(ExcelJS, bytes, mapping);
    trackingImportedBytes = bytes;
    renderTrackingTable();
    setTrackingStatus(`已匯入追蹤表，共 ${trackingRows.length} 列，可直接在畫面上編輯。`);
  } catch (err) {
    setTrackingStatus(`匯入失敗：${(err as Error).message}`, true);
  }
}

function handleTrackingCreateNew(): void {
  trackingRows = [];
  trackingImportedBytes = undefined;
  renderTrackingTable();
  setTrackingStatus("已建立一份新的空白追蹤表，按「新增列」開始輸入，完成後記得「儲存並下載」。");
}

function handleTrackingAddRow(): void {
  $<HTMLTableSectionElement>("trackingTableBody").appendChild(buildTrackingRowEl(blankTrackingRow()));
  $<HTMLDivElement>("trackingTableSection").hidden = false;
}

async function handleTrackingSave(): Promise<void> {
  try {
    const rows = collectTrackingRowsFromUI();
    const mapping = readTrackingSettingsForm();
    const bytes = await buildTrackingWorkbook(ExcelJS, mapping, rows, trackingImportedBytes);
    const fileName = $<HTMLInputElement>("trackingMgmtFile").files?.[0]?.name ?? "追蹤表.xlsx";
    downloadBlob(new Blob([bytes]), fileName);
    setTrackingStatus("已下載更新後的追蹤表，請覆蓋回原本存放的位置。");
  } catch (err) {
    setTrackingStatus(`儲存失敗：${(err as Error).message}`, true);
  }
}

// ---- wiring -------------------------------------------------------------------

function init(): void {
  fillTrackingSettingsForm(loadSettings());
  $<HTMLButtonElement>("saveSettingsBtn").addEventListener("click", () => {
    saveSettings({ ...loadSettings(), trackingSheetMapping: readTrackingSettingsForm() });
    setTrackingStatus("設定已儲存。");
  });

  $<HTMLButtonElement>("tabBtnDocs").addEventListener("click", () => switchTab("docs"));
  $<HTMLButtonElement>("tabBtnTracking").addEventListener("click", () => switchTab("tracking"));

  wireDropZoneEl($("newDocDropZone"), $<HTMLInputElement>("newDocFiles"));
  wireDropZoneEl($("emlDropZone"), $<HTMLInputElement>("emlFile"));
  wireDropZoneEl($("trackingDropZone"), $<HTMLInputElement>("trackingFile"));
  wireDropZoneEl($("trackingMgmtDropZone"), $<HTMLInputElement>("trackingMgmtFile"));

  $<HTMLInputElement>("newDocFiles").addEventListener("change", handleNewDocFilesSelected);
  $<HTMLInputElement>("emlFile").addEventListener("change", () => void handleEmlSelected());
  $<HTMLInputElement>("trackingFile").addEventListener("change", () => {
    for (const item of docItems) void checkVersionMatch(item);
  });

  $<HTMLButtonElement>("generateBtn").addEventListener("click", () => void handleGenerateAll());

  $<HTMLInputElement>("trackingMgmtFile").addEventListener("change", () => void handleTrackingMgmtFileSelected());
  $<HTMLButtonElement>("trackingCreateNewBtn").addEventListener("click", handleTrackingCreateNew);
  $<HTMLButtonElement>("trackingAddRowBtn").addEventListener("click", handleTrackingAddRow);
  $<HTMLButtonElement>("trackingSaveBtn").addEventListener("click", () => void handleTrackingSave());
}

document.addEventListener("DOMContentLoaded", init);
