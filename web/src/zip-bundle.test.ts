import assert from "node:assert/strict";
import { test } from "node:test";

import JSZip from "jszip";

import { buildZip } from "./zip-bundle.ts";

test("buildZip produces a zip containing all named entries", async () => {
  const encoder = new TextEncoder();
  const blob = await buildZip(JSZip, [
    { name: "bilingual.xlsx", bytes: encoder.encode("fake xlsx bytes").buffer as ArrayBuffer },
    { name: "diff_report.xlsx", bytes: encoder.encode("fake diff bytes").buffer as ArrayBuffer },
  ]);

  const buffer = Buffer.from(await blob.arrayBuffer());
  const reopened = await JSZip.loadAsync(buffer);
  assert.deepEqual(Object.keys(reopened.files).sort(), ["bilingual.xlsx", "diff_report.xlsx"]);
  assert.equal(await reopened.file("bilingual.xlsx")!.async("string"), "fake xlsx bytes");
});
