// End-to-end smoke test: drives the built dist/index.html in a real Chromium instance
// (via Playwright) through the v1 -> v2 golden path, exercising every feature added in
// response to real-user feedback: multiple worksheet tabs with different layouts,
// drag-and-drop file input, .eml sender/date auto-parsing, and tracking-sheet
// filename-based version matching (including the disambiguation prompt when a
// filename doesn't match anything on record).
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

async function buildV1Doc(filePath) {
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

/** Simulates a real drag-and-drop onto a drop zone, rather than just calling
 * setInputFiles directly on the underlying <input> — this is the one feature that has
 * no other code path to fall back on, so it's worth exercising for real. */
async function dropFileOnZone(page, zoneSelector, inputId, filePath, mimeType) {
  const bytes = await readFile(filePath);
  const base64 = bytes.toString("base64");
  const fileName = path.basename(filePath);
  await page.evaluate(
    ({ zoneSelector, base64, fileName, mimeType }) => {
      const binary = atob(base64);
      const arr = Uint8Array.from(binary, (c) => c.charCodeAt(0));
      const file = new File([arr], fileName, { type: mimeType });
      const dt = new DataTransfer();
      dt.items.add(file);
      const zone = document.querySelector(zoneSelector);
      zone.dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: dt }));
    },
    { zoneSelector, base64, fileName, mimeType }
  );
  // sanity check the drop actually landed on the target input, not just the DOM event
  const count = await page.$eval(`#${inputId}`, (el) => el.files.length);
  assert.equal(count, 1, `drag-and-drop onto ${zoneSelector} did not populate #${inputId}`);
}

async function setSheetCardRoles(page, sheetName, roles) {
  const card = page.locator(`.sheet-mapping-card[data-sheet-name="${sheetName}"]`);
  // Sheets default to collapsed/unchecked (real documents can have 20+ tabs, most of
  // which aren't translatable tables) — must opt in before the column selects show.
  await card.locator(".sheet-enable").check();
  for (const [col, role] of Object.entries(roles)) {
    await card.locator(`select[data-col="${col}"]`).selectOption(role);
  }
}

async function readSheetCardRoles(page, sheetName, cols) {
  const card = page.locator(`.sheet-mapping-card[data-sheet-name="${sheetName}"]`);
  const result = {};
  for (const col of cols) {
    result[col] = await card.locator(`select[data-col="${col}"]`).inputValue();
  }
  return result;
}

async function main() {
  const tmp = await mkdtemp(path.join(tmpdir(), "docmgr-e2e-"));
  const v1Path = path.join(tmp, "customer_v1.xlsx");
  const trackingPath = path.join(tmp, "tracking.xlsx");
  const emlPath = path.join(tmp, "customer_email.eml");
  await buildV1Doc(v1Path);
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

  // ---- v1: drag-and-drop the new doc, upload eml + tracking sheet ----
  await dropFileOnZone(page, "#newDocDropZone", "newDocFile", v1Path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  await page.waitForSelector(".sheet-mapping-card");
  const cardNames = await page.$$eval(".sheet-mapping-card", (cards) => cards.map((c) => c.dataset.sheetName));
  assert.deepEqual(new Set(cardNames), new Set([SHEET_REQ, SHEET_LOG]), "both worksheet tabs must get their own card");

  await setSheetCardRoles(page, SHEET_REQ, { A: "key", C: "zh1", D: "en1" });
  await setSheetCardRoles(page, SHEET_LOG, { A: "key", B: "zh1", C: "en1" });

  await page.setInputFiles("#emlFile", emlPath);
  await page.waitForFunction(() => document.getElementById("sender").value.length > 0);
  assert.equal(await page.inputValue("#sender"), "陳宜筠", "eml sender must be decoded from MIME-encoded header");
  assert.equal(await page.inputValue("#receivedDate"), "2026-07-07");

  await page.setInputFiles("#trackingFile", trackingPath);
  // tracking sheet has zero rows yet, so version resolves to 1 automatically, no prompt
  await page.waitForFunction(() => document.getElementById("versionNo").value === "1");
  assert.equal(await page.isHidden("#versionMatchSection"), true);

  await page.click("#analyzeBtn");
  await page.waitForFunction(() => document.getElementById("status").textContent.includes("首次文件"));

  const v1Inputs = await page.$$("#translationTableBody input[data-change-id]");
  assert.equal(v1Inputs.length, 3, "2 segments from 需求列表 + 1 from 維護紀錄");
  for (let i = 0; i < v1Inputs.length; i++) {
    await v1Inputs[i].fill(`EN-${i}`);
  }

  const [v1Download] = await Promise.all([page.waitForEvent("download"), page.click("#generateBtn")]);
  const v1ZipPath = path.join(tmp, "v1_output.zip");
  await v1Download.saveAs(v1ZipPath);

  const v1Zip = await JSZip.loadAsync(await readFile(v1ZipPath));
  assert.ok(!v1Zip.file("diff_report.xlsx"), "v1 (first version) must not have a diff report");

  const v1BilingualBuf = await zipEntryBuffer(v1Zip, "bilingual.xlsx");
  const v1Bilingual = new ExcelJS.Workbook();
  await v1Bilingual.xlsx.load(v1BilingualBuf);
  assert.equal(v1Bilingual.getWorksheet(SHEET_REQ).getCell("D2").value, "EN-0");
  assert.equal(v1Bilingual.getWorksheet(SHEET_REQ).getCell("D3").value, "EN-1");
  assert.equal(v1Bilingual.getWorksheet(SHEET_LOG).getCell("C2").value, "EN-2");

  const v1TrackingBuf = await zipEntryBuffer(v1Zip, "tracking.xlsx");
  const v1Tracking = new ExcelJS.Workbook();
  await v1Tracking.xlsx.load(v1TrackingBuf);
  const v1TrackingWs = v1Tracking.worksheets[0];
  assert.equal(v1TrackingWs.getCell("A2").value, 1);
  assert.equal(v1TrackingWs.getCell("C2").value, "陳宜筠");
  assert.equal(v1TrackingWs.getCell("Q2").value, "customer_v1.xlsx");
  console.log("v1 golden path OK (multi-sheet + eml + drag-and-drop + auto version 1)");

  // ---- v2: customer renamed the file, so filename won't exact-match the tracking row ----
  const v1BilingualPath = path.join(tmp, "v1_bilingual.xlsx");
  await writeFile(v1BilingualPath, v1BilingualBuf);
  await writeFile(trackingPath, v1TrackingBuf); // carry the updated tracking sheet forward

  const v2Wb = new ExcelJS.Workbook();
  await v2Wb.xlsx.load(v1BilingualBuf);
  v2Wb.getWorksheet(SHEET_REQ).getCell("C3").value = "第二項需求（已修改）";
  v2Wb.getWorksheet(SHEET_LOG).getCell("A3").value = "2026-07-14";
  v2Wb.getWorksheet(SHEET_LOG).getCell("B3").value = "王小明";
  const v2Path = path.join(tmp, "customer_v2_0714.xlsx"); // deliberately different filename
  await v2Wb.xlsx.writeFile(v2Path);

  // Playwright's setInputFiles only fires "change" on the first call for a given
  // input — clearing first forces a real change event on this second use, matching
  // what a real browser does every time a user picks a file via the native dialog.
  await page.setInputFiles("#newDocFile", []);
  await page.setInputFiles("#newDocFile", v2Path);
  await page.waitForFunction(
    (sheetName) => {
      const cards = document.querySelectorAll(".sheet-mapping-card");
      return cards.length === 2 && Array.from(cards).some((c) => c.dataset.sheetName === sheetName);
    },
    SHEET_REQ
  );
  // saved per-sheet mappings from v1 should already be applied — no re-selection needed
  const restoredRoles = await readSheetCardRoles(page, SHEET_REQ, ["A", "C", "D"]);
  assert.deepEqual(restoredRoles, { A: "key", C: "zh1", D: "en1" });

  await page.setInputFiles("#prevDocFile", v1BilingualPath);
  await page.setInputFiles("#trackingFile", []);
  await page.setInputFiles("#trackingFile", trackingPath);

  // filename doesn't match anything on record -> disambiguation prompt must appear
  await page.waitForFunction(() => document.getElementById("versionMatchSection").hidden === false);
  const matchInfo = await page.textContent("#versionMatchInfo");
  assert.ok(matchInfo.includes("customer_v2_0714.xlsx"), "prompt should name the unmatched filename");

  await page.selectOption("#matchExistingSelect", "1"); // "this continues from version 1"
  assert.equal(await page.inputValue("#versionNo"), "2", "picking version 1's entry should resolve to version 2");

  await page.click("#analyzeBtn");
  await page.waitForFunction(() => document.getElementById("status").textContent.includes("比對完成"));

  const diffRows = await page.$$eval("#diffTableBody tr", (rows) =>
    rows.map((r) => Array.from(r.children).map((c) => c.textContent))
  );
  assert.equal(diffRows.length, 2, "expected exactly 2 diff rows across both sheets (modified + added)");
  assert.ok(diffRows.some((r) => r[0].startsWith(SHEET_REQ) && r[4] === "修改"));
  assert.ok(diffRows.some((r) => r[0].startsWith(SHEET_LOG) && r[4] === "新增"));

  const v2Inputs = await page.$$("#translationTableBody input[data-change-id]");
  assert.equal(v2Inputs.length, 2, "only the modified+added rows need translation");
  await v2Inputs[0].fill("Item two revised");
  await v2Inputs[1].fill("Xiaoming Wang");

  const [v2Download] = await Promise.all([page.waitForEvent("download"), page.click("#generateBtn")]);
  const v2ZipPath = path.join(tmp, "v2_output.zip");
  await v2Download.saveAs(v2ZipPath);

  const v2Zip = await JSZip.loadAsync(await readFile(v2ZipPath));
  assert.ok(v2Zip.file("diff_report.xlsx"), "v2 zip must contain a diff report");

  const v2BilingualBuf = await zipEntryBuffer(v2Zip, "bilingual.xlsx");
  const v2Bilingual = new ExcelJS.Workbook();
  await v2Bilingual.xlsx.load(v2BilingualBuf);
  assert.equal(v2Bilingual.getWorksheet(SHEET_REQ).getCell("D2").value, "EN-0", "unchanged translation carried forward");
  assert.equal(v2Bilingual.getWorksheet(SHEET_REQ).getCell("D3").value, "Item two revised");
  assert.equal(v2Bilingual.getWorksheet(SHEET_LOG).getCell("C3").value, "Xiaoming Wang");

  const v2TrackingBuf = await zipEntryBuffer(v2Zip, "tracking.xlsx");
  const v2Tracking = new ExcelJS.Workbook();
  await v2Tracking.xlsx.load(v2TrackingBuf);
  const v2TrackingWs = v2Tracking.worksheets[0];
  assert.equal(v2TrackingWs.rowCount, 3, "header + v1 + v2, no duplicate rows");
  assert.equal(v2TrackingWs.getCell("A3").value, 2);
  assert.equal(v2TrackingWs.getCell("Q3").value, "customer_v2_0714.xlsx");
  console.log("v2 golden path OK (renamed file correctly disambiguated to version 2)");

  if (consoleErrors.length > 0) {
    console.error("Console errors were logged during the run:", consoleErrors);
    process.exitCode = 1;
  } else {
    console.log("ALL E2E CHECKS PASSED, no console errors");
  }

  await browser.close();
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
