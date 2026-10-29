import test from "node:test";
import assert from "node:assert/strict";
import { detectFormat } from "../dist/detection/detect-format.js";

test("format detection prefers file signatures and supports metadata fallback", () => {
  assert.equal(detectFormat("misleading.pdf", "image/png", Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10])), "png");
  assert.equal(detectFormat("report.docx"), "docx");
  assert.equal(detectFormat("data.bin", "application/pdf"), "pdf");
  assert.equal(detectFormat("renamed.bin", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", Uint8Array.from([0x50, 0x4b, 3, 4])), "docx");
  assert.equal(detectFormat("legacy.xls", "application/octet-stream", Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])), "xls");
  assert.equal(detectFormat("archive.bin"), "unknown");
});
