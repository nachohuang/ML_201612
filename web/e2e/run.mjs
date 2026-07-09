// End-to-end smoke test: drives the built dist/index.html in a real Chromium instance
// (via Playwright) through the batch-processing golden path, exercising every feature
// added in response to real-user feedback: multiple new documents selected/dropped at
// once (each gets its own card with its own mapping/version/diff/translation/notes),
// multiple worksheet tabs with different layouts, drag-and-drop, .eml sender/date
// auto-parsing shared across the whole batch, tracking-sheet filename-based version
// matching (including the disambiguation prompt, and the fix so two different
// documents that both resolve to "version 1" don't collide in the same tracking
// sheet), per-row notes, and the combined diff+notes+translation report. A separate
// scenario covers the Word paragraph-insertion alignment regression, and another
// covers the tracking-management tab (import/create, on-screen edit, save).
//
// Run with: node e2e/run.mjs

import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import ExcelJS from "exceljs";
import JSZip from "jszip";
import { chromium } from "playwright";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const INDEX_HTML = path.join(__dirname, "..", "dist", "index.html");

const SHEET_REQ = "需求列表";
const SHEET_LOG = "維護紀錄";
const SHEET_SCREEN = "畫面截圖";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function wordParagraph(text) {
  return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

async function buildDocx(filePath, bodyXml) {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W_NS}"><w:body>${bodyXml}</w:body></w:document>`;
  const zip = new JSZip();
  zip.file("word/document.xml", documentXml);
  const buf = await zip.generateAsync({ type: "nodebuffer" });
  await writeFile(filePath, buf);
}

async function buildV1ExcelDoc(filePath) {
  const wb = new ExcelJS.Workbook();
  const reqWs = wb.addWorksheet(SHEET_REQ);
  reqWs.getCell("A1").value = "編號";
  reqWs.getCell("C1").value = "需求說明";
  reqWs.getCell("D1").value = "English";
  reqWs.getCell("A2").value = 1;
  reqWs.getCell("C2").value = "第一項需求";
  reqWs.getCell("A3").value = 2;
  reqWs.getCell("C3").value = "第二項需求";

  const logWs = wb.addWorksheet(SHEET_LOG);
  logWs.getCell("A1").value = "日期";
  logWs.getCell("B1").value = "維護人員";
  logWs.getCell("C1").value = "Maintainer (EN)";
  logWs.getCell("A2").value = "2026-07-01";
  logWs.getCell("B2").value = "陳宜筠";

  // A free-form screen mockup tab — no header row, no clean zh/en columns, and
  // deliberately no merged cells so this test's row counts stay fully deterministic.
  const screenWs = wb.addWorksheet(SHEET_SCREEN);
  screenWs.getCell("A1").value = "查詢畫面";
  screenWs.getCell("A2").value = "帳號";
  screenWs.getCell("B2").value = "幣別";

  await wb.xlsx.writeFile(filePath);
}

async function buildTrackingSheet(filePath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("追蹤");
  const headers = [
    "versionNo", "receivedDate", "sender", "sourceFileName", "emailPath", "bilingualFilePath", "diffReportPath",
    "updateSummary", "isFirstVersion", "overseasConfirmStatus", "overseasConfirmDate",
    "overseasConfirmMethod", "customerConfirmStatus", "customerConfirmDate", "customerConfirmMethod",
    "status", "notes",
  ];
  headers.forEach((h, i) => (ws.getRow(1).getCell(i + 1).value = h));
  await wb.xlsx.writeFile(filePath);
}

function buildEml(senderNameUtf8, senderEmail, dateHeader) {
  const encodedName = Buffer.from(senderNameUtf8, "utf-8").toString("base64");
  return [
    `From: =?UTF-8?B?${encodedName}?= <${senderEmail}>`,
    "Subject: 需求文件更新",
    `Date: ${dateHeader}`,
    "",
    "請查收附件，謝謝。",
  ].join("\r\n");
}

async function zipEntryBuffer(zip, name) {
  const file = zip.file(name);
  assert.ok(file, `zip is missing ${name}`);
  return Buffer.from(await file.async("nodebuffer"));
}

/** Simulates a real multi-file drag-and-drop onto a drop zone (rather than just calling
 * setInputFiles) — this is the one feature with no other code path to fall back on, so
 * it's worth exercising for real, now with more than one file at once. */
async function dropFilesOnZone(page, zoneSelector, filePaths, mimeTypes) {
  const items = [];
  for (let i = 0; i < filePaths.length; i++) {
    const bytes = await readFile(filePaths[i]);
    items.push({ base64: bytes.toString("base64"), fileName: path.basename(filePaths[i]), mimeType: mimeTypes[i] });
  }
  await page.evaluate(
    ({ zoneSelector, items }) => {
      const dt = new DataTransfer();
      for (const it of items) {
        const binary = atob(it.base64);
        const arr = Uint8Array.from(binary, (c) => c.charCodeAt(0));
        dt.items.add(new File([arr], it.fileName, { type: it.mimeType }));
      }
      document.querySelector(zoneSelector).dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: dt }));
    },
    { zoneSelector, items }
  );
  const count = await page.$eval("#newDocFiles", (el) => el.files.length);
  assert.equal(count, filePaths.length, `drag-and-drop onto ${zoneSelector} did not populate all files`);
}

function cardAt(page, idx) {
  return page.locator(".doc-item-card").nth(idx);
}

async function setSheetCardRoles(card, sheetName, roles) {
  const sheetCard = card.locator(`.sheet-mapping-card[data-sheet-name="${sheetName}"]`);
  await sheetCard.locator(".sheet-mode").selectOption("table");
  for (const [col, role] of Object.entries(roles)) {
    await sheetCard.locator(`select[data-col="${col}"]`).selectOption(role);
  }
}

async function readSheetCardRoles(card, sheetName, cols) {
  const sheetCard = card.locator(`.sheet-mapping-card[data-sheet-name="${sheetName}"]`);
  const result = {};
  for (const col of cols) {
    result[col] = await sheetCard.locator(`select[data-col="${col}"]`).inputValue();
  }
  return result;
}

async function main() {
  const tmp = await mkdtemp(path.join(tmpdir(), "docmgr-e2e-"));
  const v1ExcelPath = path.join(tmp, "customer_v1.xlsx");
  const v1WordPath = path.join(tmp, "requirement_v1.docx");
  const trackingPath = path.join(tmp, "tracking.xlsx");
  const emlPath = path.join(tmp, "customer_email.eml");
  await buildV1ExcelDoc(v1ExcelPath);
  await buildDocx(v1WordPath, wordParagraph("計算Calculate") + wordParagraph("查詢List"));
  await buildTrackingSheet(trackingPath);
  await writeFile(emlPath, buildEml("陳宜筠", "yiyun@example.com", "Tue, 07 Jul 2026 10:23:45 +0800"));

  const browser = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" });
  const page = await browser.newPage();
  const consoleErrors = [];
  page.on("pageerror", (err) => consoleErrors.push(String(err)));
  page.on("console", (msg) => {
    if (msg.type() === "error") consoleErrors.push(msg.text());
  });

  await page.goto(`file://${INDEX_HTML}`);

  // ================================================================================
  // Batch v1: drag-and-drop a fresh Excel doc AND a fresh Word doc at the same time —
  // several requirement documents can arrive attached to the same customer email.
  // ================================================================================
  await dropFilesOnZone(
    page,
    "#newDocDropZone",
    [v1ExcelPath, v1WordPath],
    [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ]
  );

  const fileListItems = await page.$$eval("#newDocFileList li", (els) => els.map((el) => el.textContent));
  assert.equal(fileListItems.length, 2, "file list must show every dropped file");
  assert.ok(fileListItems.some((t) => t.includes("customer_v1.xlsx")));
  assert.ok(fileListItems.some((t) => t.includes("requirement_v1.docx")));

  await page.waitForFunction(() => document.querySelectorAll(".doc-item-card").length === 2);
  const excelCard = cardAt(page, 0); // insertion order is preserved from the drop
  const wordCard = cardAt(page, 1);

  // ---- Excel card: every tab included by default, defaulting to freeform ----
  await page.waitForFunction(
    (sheetName) => Array.from(document.querySelectorAll(".sheet-mapping-card")).some((c) => c.dataset.sheetName === sheetName),
    SHEET_REQ
  );
  const cardSheetNames = await excelCard.locator(".sheet-mapping-card").evaluateAll((cards) => cards.map((c) => c.dataset.sheetName));
  assert.deepEqual(new Set(cardSheetNames), new Set([SHEET_REQ, SHEET_LOG, SHEET_SCREEN]), "every worksheet tab must get its own card");
  for (const name of [SHEET_REQ, SHEET_LOG, SHEET_SCREEN]) {
    assert.equal(await excelCard.locator(`.sheet-mapping-card[data-sheet-name="${name}"] .sheet-enable`).isChecked(), true);
  }
  // SHEET_SCREEN deliberately left in its default freeform mode (no column setup).
  await setSheetCardRoles(excelCard, SHEET_REQ, { A: "key", C: "zh1", D: "en1" });
  await setSheetCardRoles(excelCard, SHEET_LOG, { A: "key", B: "zh1", C: "en1" });

  // ---- Word card: freeform is the default, no column setup needed ----
  assert.equal(await wordCard.locator(".word-mode").inputValue(), "freeform", "freeform must be the default word mapping mode");

  // ---- shared eml + tracking file apply to every card in the batch ----
  await page.setInputFiles("#emlFile", emlPath);
  await page.waitForFunction(() => document.querySelector(".doc-item-card .item-sender").value.length > 0);
  for (const card of [excelCard, wordCard]) {
    assert.equal(await card.locator(".item-sender").inputValue(), "陳宜筠", "eml sender must be decoded and applied to every card");
    assert.equal(await card.locator(".item-received-date").inputValue(), "2026-07-07");
  }

  await page.setInputFiles("#trackingFile", trackingPath);
  await page.waitForFunction(() =>
    Array.from(document.querySelectorAll(".doc-item-card")).every((c) => c.querySelector(".item-version-no").value === "1")
  );
  assert.equal(await excelCard.locator(".version-match-section").isHidden(), true);
  assert.equal(await wordCard.locator(".version-match-section").isHidden(), true);

  // ---- analyze both documents, independently ----
  await excelCard.locator(".analyze-btn").click();
  await page.waitForFunction(() => document.querySelectorAll(".doc-item-card")[0].querySelector(".item-status").textContent.includes("首次文件"));
  const excelTranslationInputs = excelCard.locator(".translation-table-body input[data-change-id]");
  assert.equal(await excelTranslationInputs.count(), 3, "2 segments from 需求列表 + 1 from 維護紀錄");
  assert.equal(await excelCard.locator(".diff-table-body tr").count(), 6, "2+1 table rows plus 3 freeform cells from 畫面截圖, all v1 ADDED");

  await wordCard.locator(".analyze-btn").click();
  await page.waitForFunction(() => document.querySelectorAll(".doc-item-card")[1].querySelector(".item-status").textContent.includes("首次文件"));
  assert.equal(await wordCard.locator(".translation-table-body input[data-change-id]").count(), 0, "freeform word mode never needs translation input");
  assert.equal(await wordCard.locator(".diff-table-body tr").count(), 2, "both v1 paragraphs reported as added");

  // ---- fill translations, and a note, on the Excel card ----
  for (let i = 0; i < 3; i++) await excelTranslationInputs.nth(i).fill(`EN-${i}`);
  await excelCard.locator(".diff-table-body tr").nth(0).locator(".note-input").fill("已跟客戶確認過");

  // ---- generate the whole batch in one zip ----
  const [download1] = await Promise.all([page.waitForEvent("download"), page.click("#generateBtn")]);
  const zipPath1 = path.join(tmp, "v1_output.zip");
  await download1.saveAs(zipPath1);
  const zip1 = await JSZip.loadAsync(await readFile(zipPath1));

  assert.ok(zip1.file("bilingual_customer_v1.xlsx"), "excel bilingual output must be named after its source file");
  assert.ok(zip1.file("bilingual_requirement_v1.docx"), "word bilingual output must be named after its source file");
  assert.ok(zip1.file("diff_report.xlsx"), "v1 now also gets a combined report, since every field starts as an ADDED row to review");
  assert.ok(zip1.file(path.basename(emlPath)));
  assert.ok(zip1.file("tracking.xlsx"));

  const excelBilingualBuf = await zipEntryBuffer(zip1, "bilingual_customer_v1.xlsx");
  const excelBilingual = new ExcelJS.Workbook();
  await excelBilingual.xlsx.load(excelBilingualBuf);
  assert.equal(excelBilingual.getWorksheet(SHEET_REQ).getCell("D2").value, "EN-0");
  assert.equal(excelBilingual.getWorksheet(SHEET_REQ).getCell("D3").value, "EN-1");
  assert.equal(excelBilingual.getWorksheet(SHEET_LOG).getCell("C2").value, "EN-2");
  assert.equal(excelBilingual.getWorksheet(SHEET_SCREEN).getCell("A2").value, "帳號", "freeform sheet passes through untouched");

  const reportBuf1 = await zipEntryBuffer(zip1, "diff_report.xlsx");
  const reportWb1 = new ExcelJS.Workbook();
  await reportWb1.xlsx.load(reportBuf1);
  const reportWs1 = reportWb1.worksheets[0];
  assert.deepEqual(reportWs1.getRow(1).values.slice(1), ["文件名稱", "位置", "欄位名稱", "舊內容", "新內容", "變動類型", "備註", "翻譯"]);
  const reportRows1 = [];
  reportWs1.eachRow((row, num) => {
    if (num > 1) reportRows1.push(row.values.slice(1));
  });
  assert.equal(reportRows1.length, 8, "excel v1's 6 rows + word v1's 2 rows, combined into one report");
  assert.ok(reportRows1.some((r) => r[0] === "customer_v1.xlsx" && r[6] === "已跟客戶確認過"), "note must be carried into the combined report");
  assert.ok(reportRows1.some((r) => r[0] === "requirement_v1.docx"), "word document's rows must be present too, tagged with its own name");

  const trackingBuf1 = await zipEntryBuffer(zip1, "tracking.xlsx");
  const trackingWb1 = new ExcelJS.Workbook();
  await trackingWb1.xlsx.load(trackingBuf1);
  const trackingWs1 = trackingWb1.worksheets[0];
  assert.equal(trackingWs1.rowCount, 3, "header + one row per document, both independently version 1");
  assert.equal(trackingWs1.getCell("A2").value, 1);
  assert.equal(trackingWs1.getCell("Q2").value, "customer_v1.xlsx");
  assert.equal(trackingWs1.getCell("A3").value, 1);
  assert.equal(trackingWs1.getCell("Q3").value, "requirement_v1.docx");
  console.log("Batch v1 golden path OK (multi-file drag-drop + per-card mapping + notes + combined report + no version collision)");

  // ================================================================================
  // v2: the customer renamed the Excel file (no exact filename match) and edited it —
  // continues on the SAME page, so localStorage-persisted sheet-mapping settings from
  // v1 should already be applied without re-selection.
  // ================================================================================
  const v1BilingualPath = path.join(tmp, "v1_bilingual.xlsx");
  await writeFile(v1BilingualPath, excelBilingualBuf);
  await writeFile(trackingPath, trackingBuf1); // carry the updated tracking sheet forward

  const v2Wb = new ExcelJS.Workbook();
  await v2Wb.xlsx.load(excelBilingualBuf);
  v2Wb.getWorksheet(SHEET_REQ).getCell("C3").value = "第二項需求（已修改）";
  v2Wb.getWorksheet(SHEET_LOG).getCell("A3").value = "2026-07-14";
  v2Wb.getWorksheet(SHEET_LOG).getCell("B3").value = "王小明";
  v2Wb.getWorksheet(SHEET_SCREEN).getCell("A2").value = "帳號（必填）"; // freeform cell edit
  const v2Path = path.join(tmp, "customer_v2_0714.xlsx"); // deliberately renamed
  await v2Wb.xlsx.writeFile(v2Path);

  await page.setInputFiles("#newDocFiles", [v2Path]);
  await page.waitForFunction(() => document.querySelectorAll(".doc-item-card").length === 1);
  const v2Card = cardAt(page, 0);

  await page.waitForFunction(
    (sheetName) => Array.from(document.querySelectorAll(".sheet-mapping-card")).some((c) => c.dataset.sheetName === sheetName),
    SHEET_REQ
  );
  // saved per-sheet mappings from v1 should already be applied — no re-selection needed
  const restoredRoles = await readSheetCardRoles(v2Card, SHEET_REQ, ["A", "C", "D"]);
  assert.deepEqual(restoredRoles, { A: "key", C: "zh1", D: "en1" }, "sheet-mapping settings must persist across documents via localStorage");

  await page.setInputFiles(".prev-file-input", v1BilingualPath);
  await page.setInputFiles("#trackingFile", []);
  await page.setInputFiles("#trackingFile", trackingPath);

  // filename doesn't exactly match anything on record -> disambiguation prompt must appear
  await page.waitForFunction(() => document.querySelector(".version-match-section").hidden === false);
  const matchInfo = await v2Card.locator(".version-match-section p").first().textContent();
  assert.ok(matchInfo.includes("customer_v2_0714.xlsx"), "prompt should name the unmatched filename");

  await v2Card.locator(".match-existing-select").selectOption("1"); // "this continues from version 1"
  assert.equal(await v2Card.locator(".item-version-no").inputValue(), "2", "picking version 1's entry should resolve to version 2");

  await v2Card.locator(".analyze-btn").click();
  await page.waitForFunction(() => document.querySelector(".item-status").textContent.includes("比對完成"));

  const v2DiffRows = await v2Card.locator(".diff-table-body tr").evaluateAll((rows) => rows.map((r) => Array.from(r.children).map((c) => c.textContent)));
  assert.equal(v2DiffRows.length, 3, "table-mode modified+added, plus the freeform sheet's changed cell");
  assert.ok(v2DiffRows.some((r) => r[0].startsWith(SHEET_REQ) && r[4] === "修改"));
  assert.ok(v2DiffRows.some((r) => r[0].startsWith(SHEET_LOG) && r[4] === "新增"));
  assert.ok(
    v2DiffRows.some((r) => r[0].startsWith(SHEET_SCREEN) && r[4] === "修改" && r[3] === "帳號（必填）"),
    "the freeform sheet's cell edit must show up in the diff report"
  );

  const v2TranslationInputs = v2Card.locator(".translation-table-body input[data-change-id]");
  assert.equal(await v2TranslationInputs.count(), 2, "only the table-mode modified+added rows need translation, not the freeform cell");
  await v2TranslationInputs.nth(0).fill("Item two revised");
  await v2TranslationInputs.nth(1).fill("Xiaoming Wang");

  const [download2] = await Promise.all([page.waitForEvent("download"), page.click("#generateBtn")]);
  const zipPath2 = path.join(tmp, "v2_output.zip");
  await download2.saveAs(zipPath2);
  const zip2 = await JSZip.loadAsync(await readFile(zipPath2));

  assert.ok(zip2.file("diff_report.xlsx"));
  const v2BilingualBuf = await zipEntryBuffer(zip2, "bilingual_customer_v2_0714.xlsx");
  const v2Bilingual = new ExcelJS.Workbook();
  await v2Bilingual.xlsx.load(v2BilingualBuf);
  assert.equal(v2Bilingual.getWorksheet(SHEET_REQ).getCell("D2").value, "EN-0", "unchanged translation carried forward");
  assert.equal(v2Bilingual.getWorksheet(SHEET_REQ).getCell("D3").value, "Item two revised");
  assert.equal(v2Bilingual.getWorksheet(SHEET_LOG).getCell("C3").value, "Xiaoming Wang");
  assert.equal(v2Bilingual.getWorksheet(SHEET_SCREEN).getCell("A2").value, "帳號（必填）", "freeform sheet passes through as-is");

  const trackingBuf2 = await zipEntryBuffer(zip2, "tracking.xlsx");
  const trackingWb2 = new ExcelJS.Workbook();
  await trackingWb2.xlsx.load(trackingBuf2);
  const trackingWs2 = trackingWb2.worksheets[0];
  assert.equal(trackingWs2.rowCount, 4, "header + excel-v1 + word-v1 (untouched) + excel-v2, no duplicate/overwritten rows");
  assert.equal(trackingWs2.getCell("A4").value, 2);
  assert.equal(trackingWs2.getCell("Q4").value, "customer_v2_0714.xlsx");
  console.log("v2 golden path OK (renamed file disambiguated, settings persisted, tracking rows not clobbered)");

  await writeFile(trackingPath, trackingBuf2); // carry the final (4-row) tracking sheet forward

  if (consoleErrors.length > 0) {
    console.error("Console errors were logged during the run:", consoleErrors);
    process.exitCode = 1;
  } else {
    console.log("ALL E2E CHECKS PASSED, no console errors");
  }

  await browser.close();
  return { finalTrackingPath: trackingPath, finalTrackingBuf: trackingBuf2 };
}

/** Word support's freeform mode aligns paragraphs by content (LCS), not fixed position —
 * a plain position-based compare misreads "one paragraph inserted" as "every subsequent
 * paragraph changed", a real bug found against actual customer documents. This isolates
 * that regression: an edit and an unrelated insertion elsewhere must not bleed together. */
async function runWordInsertionRegressionScenario(browser) {
  const page = await browser.newPage();
  const consoleErrors = [];
  page.on("pageerror", (err) => consoleErrors.push(String(err)));
  page.on("console", (msg) => {
    if (msg.type() === "error") consoleErrors.push(msg.text());
  });

  await page.goto(`file://${INDEX_HTML}`);

  const tmp = await mkdtemp(path.join(tmpdir(), "docmgr-e2e-word-"));
  const v1Path = path.join(tmp, "requirement_v1.docx");
  await buildDocx(
    v1Path,
    wordParagraph("計算Calculate (新增Create)") + wordParagraph("查詢List") + wordParagraph("確認Confirm")
  );

  await page.setInputFiles("#newDocFiles", [v1Path]);
  await page.waitForFunction(() => document.querySelectorAll(".doc-item-card").length === 1);
  const card = cardAt(page, 0);
  assert.equal(await card.locator(".word-mode").inputValue(), "freeform");

  await card.locator(".analyze-btn").click();
  await page.waitForFunction(() => document.querySelector(".item-status").textContent.includes("首次文件"));

  const [download1] = await Promise.all([page.waitForEvent("download"), page.click("#generateBtn")]);
  const zipPath1 = path.join(tmp, "v1_output.zip");
  await download1.saveAs(zipPath1);
  const zip1 = await JSZip.loadAsync(await readFile(zipPath1));
  const v1BilingualBuf = Buffer.from(await zip1.file("bilingual_requirement_v1.docx").async("nodebuffer"));
  const v1BilingualPath = path.join(tmp, "v1_bilingual.docx");
  await writeFile(v1BilingualPath, v1BilingualBuf);

  // v2: first paragraph edited in place (unrelated to the insertion below), and a
  // brand new paragraph inserted further down between two otherwise-unchanged ones.
  const v2Path = path.join(tmp, "requirement_v2.docx");
  await buildDocx(
    v2Path,
    wordParagraph("計算Calculate（已修改）") +
      wordParagraph("查詢List") +
      wordParagraph("新增的段落New paragraph") +
      wordParagraph("確認Confirm")
  );

  await page.setInputFiles("#newDocFiles", [v2Path]);
  await page.waitForFunction(() => document.querySelectorAll(".doc-item-card").length === 1);
  const v2Card = cardAt(page, 0);
  await page.setInputFiles(".prev-file-input", v1BilingualPath);
  // no tracking file in this scenario, so the version number is never auto-resolved —
  // fill it in manually, same as a user would.
  await v2Card.locator(".item-version-no").fill("2");

  await v2Card.locator(".analyze-btn").click();
  await page.waitForFunction(() => document.querySelector(".item-status").textContent.includes("比對完成"));

  const diffRows = await v2Card.locator(".diff-table-body tr").evaluateAll((rows) => rows.map((r) => Array.from(r.children).map((c) => c.textContent)));
  assert.equal(diffRows.length, 2, "one modified paragraph + one added paragraph, insertion must not shift the rest");
  assert.ok(diffRows.some((r) => r[4] === "修改" && r[3].includes("已修改")));
  assert.ok(diffRows.some((r) => r[4] === "新增" && r[3].includes("新增的段落")));
  assert.equal(await v2Card.locator(".translation-table-body input").count(), 0, "freeform mode never needs translation input");

  const [download2] = await Promise.all([page.waitForEvent("download"), page.click("#generateBtn")]);
  const zipPath2 = path.join(tmp, "v2_output.zip");
  await download2.saveAs(zipPath2);
  const zip2 = await JSZip.loadAsync(await readFile(zipPath2));
  assert.ok(zip2.file("diff_report.xlsx"));

  if (consoleErrors.length > 0) {
    console.error("Word insertion-regression scenario console errors:", consoleErrors);
    process.exitCode = 1;
  } else {
    console.log("Word freeform insertion-regression OK (insertion correctly isolated, no translation slots)");
  }
  await page.close();
}

/** Tracking-management tab: import an existing tracking sheet, edit it directly on
 * screen (including adding a brand-new row), and save it back out — and separately,
 * building one from scratch when there's no existing tracking sheet yet. */
async function runTrackingTabScenario(browser, existingTrackingPath) {
  const page = await browser.newPage();
  const consoleErrors = [];
  page.on("pageerror", (err) => consoleErrors.push(String(err)));
  page.on("console", (msg) => {
    if (msg.type() === "error") consoleErrors.push(msg.text());
  });

  await page.goto(`file://${INDEX_HTML}`);
  await page.click("#tabBtnTracking");
  assert.equal(await page.isHidden("#docsPage"), true);
  assert.equal(await page.isHidden("#trackingPage"), false);

  // ---- import an existing tracking sheet and edit it on screen ----
  await page.setInputFiles("#trackingMgmtFile", existingTrackingPath);
  await page.waitForFunction(() => document.getElementById("trackingTableSection").hidden === false);
  const importedRowCount = await page.locator("#trackingTableBody tr").count();
  assert.equal(importedRowCount, 3, "every existing tracking data row must be shown for on-screen editing");

  // edit an existing cell directly (status column, field key "status")
  const firstStatusInput = page.locator('#trackingTableBody tr').nth(0).locator('input[data-field="status"]');
  await firstStatusInput.fill("已完成");

  // add a brand-new row and fill a couple of fields
  await page.click("#trackingAddRowBtn");
  assert.equal(await page.locator("#trackingTableBody tr").count(), importedRowCount + 1, "新增列 must append an editable row");
  const newRow = page.locator("#trackingTableBody tr").last();
  await newRow.locator('input[data-field="versionNo"]').fill("3");
  await newRow.locator('input[data-field="sourceFileName"]').fill("手動新增的列.xlsx");

  const [download] = await Promise.all([page.waitForEvent("download"), page.click("#trackingSaveBtn")]);
  const savedPath = path.join(path.dirname(existingTrackingPath), "tracking_saved.xlsx");
  await download.saveAs(savedPath);

  const savedWb = new ExcelJS.Workbook();
  await savedWb.xlsx.load(await readFile(savedPath));
  const savedWs = savedWb.worksheets[0];
  assert.equal(savedWs.rowCount, importedRowCount + 2, "header + all original rows + the new row");
  assert.equal(savedWs.getCell("O2").value, "已完成", "edited cell must be reflected in the saved file");
  const lastDataRow = importedRowCount + 2;
  assert.equal(savedWs.getCell(`A${lastDataRow}`).value, "3");
  assert.equal(savedWs.getCell(`Q${lastDataRow}`).value, "手動新增的列.xlsx");
  console.log("Tracking-management tab OK (import, on-screen edit, add row, save)");

  // ---- create a brand-new tracking sheet from scratch ----
  await page.reload();
  await page.click("#tabBtnTracking");
  await page.click("#trackingCreateNewBtn");
  await page.waitForFunction(() => document.getElementById("trackingTableSection").hidden === false);
  assert.equal(await page.locator("#trackingTableBody tr").count(), 0, "a brand-new tracking sheet starts with zero rows");

  await page.click("#trackingAddRowBtn");
  await page.locator('#trackingTableBody tr').nth(0).locator('input[data-field="versionNo"]').fill("1");
  await page.locator('#trackingTableBody tr').nth(0).locator('input[data-field="sourceFileName"]').fill("第一份文件.docx");

  const [newDownload] = await Promise.all([page.waitForEvent("download"), page.click("#trackingSaveBtn")]);
  const newSavedPath = path.join(path.dirname(existingTrackingPath), "tracking_new.xlsx");
  await newDownload.saveAs(newSavedPath);

  const newWb = new ExcelJS.Workbook();
  await newWb.xlsx.load(await readFile(newSavedPath));
  const newWs = newWb.worksheets[0];
  assert.equal(newWs.getCell("A1").value, "版本編號", "a from-scratch sheet must get header labels");
  assert.equal(newWs.getCell("A2").value, "1");
  assert.equal(newWs.getCell("Q2").value, "第一份文件.docx");
  console.log("Tracking-management tab OK (brand-new tracking sheet created from scratch)");

  if (consoleErrors.length > 0) {
    console.error("Tracking tab scenario console errors:", consoleErrors);
    process.exitCode = 1;
  } else {
    console.log("Tracking tab scenario: no console errors");
  }
  await page.close();
}

async function run() {
  const { finalTrackingPath } = await main();

  const browser = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" });
  await runWordInsertionRegressionScenario(browser);
  await runTrackingTabScenario(browser, finalTrackingPath);
  await browser.close();
}

run().catch((err) => {
  console.error(err);
  process.exit(1);
});
