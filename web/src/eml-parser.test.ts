import assert from "node:assert/strict";
import { test } from "node:test";

import { parseEml } from "./eml-parser.ts";

test("parses plain ASCII From/Subject/Date headers", () => {
  const eml = [
    "From: Jane Doe <jane@example.com>",
    "Subject: Requirement update",
    "Date: Tue, 07 Jul 2026 10:23:45 +0800",
    "",
    "body text",
  ].join("\r\n");

  const meta = parseEml(eml);
  assert.equal(meta.sender, "Jane Doe");
  assert.equal(meta.subject, "Requirement update");
  assert.equal(meta.receivedDate, "2026-07-07");
});

test("decodes base64-encoded MIME words (Chinese sender name)", () => {
  const encodedName = Buffer.from("陳宜筠", "utf-8").toString("base64");
  const eml = [
    `From: =?UTF-8?B?${encodedName}?= <yiyun@example.com>`,
    "Subject: =?UTF-8?B?6ZyA5rGC5pu05pawIQ==?=",
    "Date: Wed, 08 Jul 2026 09:00:00 +0800",
    "",
    "body",
  ].join("\r\n");

  const meta = parseEml(eml);
  assert.equal(meta.sender, "陳宜筠");
  assert.equal(meta.receivedDate, "2026-07-08");
});

test("decodes quoted-printable MIME words", () => {
  // "測試" in UTF-8 quoted-printable
  const eml = ["From: =?UTF-8?Q?=E6=B8=AC=E8=A9=A6?= <test@example.com>", "Subject: hi", "", "body"].join("\r\n");

  const meta = parseEml(eml);
  assert.equal(meta.sender, "測試");
});

test("handles folded header lines", () => {
  const eml = ["Subject: this is a very long subject", " that continues on the next line", "", "body"].join("\r\n");

  const meta = parseEml(eml);
  assert.equal(meta.subject, "this is a very long subject that continues on the next line");
});

test("returns nulls for missing headers instead of throwing", () => {
  const meta = parseEml("body only, no headers at all");
  assert.equal(meta.sender, null);
  assert.equal(meta.subject, null);
  assert.equal(meta.receivedDate, null);
});

test("sender falls back to the bare email when there's no display name", () => {
  const eml = ["From: noreply@example.com", "", "body"].join("\r\n");
  const meta = parseEml(eml);
  assert.equal(meta.sender, "noreply@example.com");
});
