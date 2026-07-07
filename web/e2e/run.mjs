// End-to-end smoke test: drives the built dist/index.html in a real Chromium instance
// (via Playwright) through the v1 -> v2 golden path, mirroring
// tests/test_pipeline_e2e.py from the Python CLI version.
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

async function buildV1Doc(filePath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.getCell("A1").value = "編號";
  ws.getCell("C1").value = "需求說明";
  ws.getCell("D1").value = "English";
  ws.getCell("A2").value = 1;
  ws.getCell("C2").value = "第一項需求";
  ws.getCell("A3").value = 2;
  ws.getCell("C3").value = "第二項需求";
  await wb.xlsx.writeFile(filePath);
}

async function buildTrackingSheet(filePath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("追蹤");
  const headers = [
    "versionNo", "receivedDate", "sender", "emailPath", "bilingualFilePath", "diffReportPath",
    "updateSummary", "isFirstVersion", "overseasConfirmStatus", "overseasConfirmDate",
    "overseasConfirmMethod", "customerConfirmStatus", "customerConfirmDate", "customerConfirmMethod",
    "status", "notes",
  ];
  headers.forEach((h, i) => (ws.getRow(1).getCell(i + 1).value = h));
  await wb.xlsx.writeFile(filePath);
}

async function zipEntryBuffer(zip, name) {
  const file = zip.file(name);
  assert.ok(file, `zip is missing ${name}`);
  return Buffer.from(await file.async("nodebuffer"));
}

async function main() {
  const tmp = await mkdtemp(path.join(tmpdir(), "docmgr-e2e-"));
  const v1Path = path.join(tmp, "customer_v1.xlsx");
  const trackingPath = path.join(tmp, "tracking.xlsx");
  await buildV1Doc(v1Path);
  await buildTrackingSheet(trackingPath);

  const browser = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" });
  const page = await browser.newPage();
  const consoleErrors = [];
  page.on("pageerror", (err) => consoleErrors.push(String(err)));
  page.on("console", (msg) => {
    if (msg.type() === "error") consoleErrors.push(msg.text());
  });

  await page.goto(`file://${INDEX_HTML}`);

  // Sanity check: default settings are pre-filled correctly on load.
  assert.equal(await page.inputValue("#excelZhColumns"), "C");
  assert.equal(await page.inputValue("#excelEnColumns"), "D");

  // ---- v1: zh-only document, no previous version, with a tracking sheet attached ----
  await page.setInputFiles("#newDocFile", v1Path);
  await page.setInputFiles("#trackingFile", trackingPath);
  await page.fill("#sender", "customer-a@example.com");
  await page.fill("#receivedDate", "2026-07-08");
  await page.click("#analyzeBtn");
  await page.waitForFunction(() => document.getElementById("status").textContent.includes("首次文件"));

  assert.equal(await page.inputValue("#versionNo"), "1");
  const v1Inputs = await page.$$("#translationTableBody input[data-change-id]");
  assert.equal(v1Inputs.length, 2, "v1 should have 2 translation items");
  await page.fill(
    '#translationTableBody tr:nth-child(1) input[data-change-id]',
    "Item one"
  );
  await page.fill(
    '#translationTableBody tr:nth-child(2) input[data-change-id]',
    "Item two"
  );

  const [v1Download] = await Promise.all([page.waitForEvent("download"), page.click("#generateBtn")]);
  const v1ZipPath = path.join(tmp, "v1_output.zip");
  await v1Download.saveAs(v1ZipPath);

  const v1Zip = await JSZip.loadAsync(await readFile(v1ZipPath));
  assert.ok(v1Zip.file("bilingual.xlsx"), "v1 zip must contain bilingual.xlsx");
  assert.ok(!v1Zip.file("diff_report.xlsx"), "v1 (first version) must not have a diff report");
  assert.ok(v1Zip.file("tracking.xlsx"), "v1 zip must contain the updated tracking sheet");
  assert.ok(v1Zip.file("使用說明.txt"), "v1 zip must contain the readme");

  const v1BilingualBuf = await zipEntryBuffer(v1Zip, "bilingual.xlsx");
  const v1Bilingual = new ExcelJS.Workbook();
  await v1Bilingual.xlsx.load(v1BilingualBuf);
  const v1Ws = v1Bilingual.worksheets[0];
  assert.equal(v1Ws.getCell("D2").value, "Item one");
  assert.equal(v1Ws.getCell("D3").value, "Item two");

  const v1TrackingBuf = await zipEntryBuffer(v1Zip, "tracking.xlsx");
  const v1Tracking = new ExcelJS.Workbook();
  await v1Tracking.xlsx.load(v1TrackingBuf);
  const v1TrackingWs = v1Tracking.worksheets[0];
  assert.equal(v1TrackingWs.getCell("A2").value, 1);
  assert.equal(v1TrackingWs.getCell("H2").value, "是");
  assert.equal(v1TrackingWs.getCell("O2").value, "待確認");
  assert.equal(v1TrackingWs.getCell("C2").value, "customer-a@example.com");
  console.log("v1 golden path OK");

  // ---- v2: customer edited the bilingual file — old English + new zh mixed together ----
  const v1BilingualPath = path.join(tmp, "v1_bilingual.xlsx");
  await writeFile(v1BilingualPath, v1BilingualBuf);
  // Feed the v1-updated tracking sheet forward (not the pristine original) so v2's
  // upsert appends as row 3, matching what a real user carrying the file forward would see.
  // Overwrite the same tracking.xlsx path: a real user replaces the same file each round.
  await writeFile(trackingPath, v1TrackingBuf);

  const v2Wb = new ExcelJS.Workbook();
  await v2Wb.xlsx.load(v1BilingualBuf);
  const v2Ws = v2Wb.worksheets[0];
  v2Ws.getCell("C3").value = "第二項需求（已修改）";
  v2Ws.getCell("A4").value = 3;
  v2Ws.getCell("C4").value = "第三項需求（新增）";
  const v2Path = path.join(tmp, "customer_v2.xlsx");
  await v2Wb.xlsx.writeFile(v2Path);

  await page.setInputFiles("#newDocFile", v2Path);
  await page.setInputFiles("#prevDocFile", v1BilingualPath);
  await page.setInputFiles("#trackingFile", trackingPath); // carry forward the updated sheet
  await page.click("#analyzeBtn");
  await page.waitForFunction(() => document.getElementById("status").textContent.includes("比對完成"));
  // versionNo only becomes editable again once analyze() resets it for a non-first version
  await page.fill("#versionNo", "2");

  const diffRows = await page.$$eval("#diffTableBody tr", (rows) =>
    rows.map((r) => Array.from(r.children).map((c) => c.textContent))
  );
  assert.equal(diffRows.length, 2, "expected exactly 2 diff rows (modified + added)");
  assert.ok(diffRows.some((r) => r[4] === "修改"));
  assert.ok(diffRows.some((r) => r[4] === "新增"));

  const v2Inputs = await page.$$("#translationTableBody input[data-change-id]");
  assert.equal(v2Inputs.length, 2, "v2 should only need translation for modified+added rows");
  await page.fill('#translationTableBody tr:nth-child(1) input[data-change-id]', "Item two revised");
  await page.fill('#translationTableBody tr:nth-child(2) input[data-change-id]', "Item three");

  const [v2Download] = await Promise.all([page.waitForEvent("download"), page.click("#generateBtn")]);
  const v2ZipPath = path.join(tmp, "v2_output.zip");
  await v2Download.saveAs(v2ZipPath);

  const v2Zip = await JSZip.loadAsync(await readFile(v2ZipPath));
  assert.ok(v2Zip.file("diff_report.xlsx"), "v2 zip must contain a diff report");

  const v2BilingualBuf = await zipEntryBuffer(v2Zip, "bilingual.xlsx");
  const v2Bilingual = new ExcelJS.Workbook();
  await v2Bilingual.xlsx.load(v2BilingualBuf);
  const v2ResultWs = v2Bilingual.worksheets[0];
  assert.equal(v2ResultWs.getCell("D2").value, "Item one", "unchanged row keeps its original translation");
  assert.equal(v2ResultWs.getCell("D3").value, "Item two revised");
  assert.equal(v2ResultWs.getCell("D4").value, "Item three");

  const v2TrackingBuf = await zipEntryBuffer(v2Zip, "tracking.xlsx");
  const v2Tracking = new ExcelJS.Workbook();
  await v2Tracking.xlsx.load(v2TrackingBuf);
  const v2TrackingWs = v2Tracking.worksheets[0];
  assert.equal(v2TrackingWs.rowCount, 3, "header + v1 + v2, no duplicate rows");
  assert.equal(v2TrackingWs.getCell("A3").value, 2);
  console.log("v2 golden path OK");

  // ---- confirm section: mark both parties confirmed on the v2 tracking sheet ----
  const v2TrackingPath = path.join(tmp, "v2_tracking.xlsx");
  await writeFile(v2TrackingPath, v2TrackingBuf);

  await page.setInputFiles("#confirmTrackingFile", v2TrackingPath);
  await page.fill("#confirmVersionNo", "2");
  await page.selectOption("#confirmParty", "overseas");
  await page.fill("#confirmDate", "2026-07-10");
  await page.fill("#confirmMethod", "Teams");
  const [confirm1] = await Promise.all([page.waitForEvent("download"), page.click("#confirmBtn")]);
  const afterOverseasPath = path.join(tmp, "after_overseas.xlsx");
  await confirm1.saveAs(afterOverseasPath);

  await page.setInputFiles("#confirmTrackingFile", afterOverseasPath);
  await page.selectOption("#confirmParty", "customer");
  await page.fill("#confirmDate", "2026-07-11");
  await page.fill("#confirmMethod", "回信");
  const [confirm2] = await Promise.all([page.waitForEvent("download"), page.click("#confirmBtn")]);
  const afterBothPath = path.join(tmp, "after_both.xlsx");
  await confirm2.saveAs(afterBothPath);

  const finalTracking = new ExcelJS.Workbook();
  await finalTracking.xlsx.load(await readFile(afterBothPath));
  const finalWs = finalTracking.worksheets[0];
  assert.equal(finalWs.getCell("I3").value, "已確認"); // overseasConfirmStatus
  assert.equal(finalWs.getCell("L3").value, "已確認"); // customerConfirmStatus
  assert.equal(finalWs.getCell("O3").value, "已完成"); // status
  console.log("confirm flow OK");

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
