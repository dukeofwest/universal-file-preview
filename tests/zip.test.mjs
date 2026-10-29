import test from "node:test";
import assert from "node:assert/strict";
import { makeZip } from "./helpers/zip.mjs";
import { ZipArchive } from "../dist/archive/zip-reader.js";

test("ZIP reader reads stored and DEFLATE-compressed entries and checks CRC", () => {
  const payload = "A bounded archive with repeated repeated repeated data.";
  for (const deflate of [false, true]) {
    const bytes = makeZip({ "folder/data.txt": payload }, { deflate });
    const zip = ZipArchive.open(bytes);
    assert.equal(zip.has("/folder/data.txt"), true);
    assert.equal(new TextDecoder().decode(zip.read("folder/data.txt")), payload);
  }
  const corrupt = makeZip({ "data.txt": "payload" });
  const nameLength = corrupt.readUInt16LE(26), dataOffset = 30 + nameLength;
  corrupt[dataOffset] ^= 0xff;
  assert.throws(() => ZipArchive.open(corrupt).read("data.txt"), /checksum mismatch/);
});

test("ZIP reader rejects truncated, traversal, and oversized archives", () => {
  assert.throws(() => ZipArchive.open(Uint8Array.from([0x50, 0x4b])), /too short/);
  assert.throws(() => ZipArchive.open(makeZip({ "../escape": "x" })), /Unsafe ZIP/);
  assert.throws(() => ZipArchive.open(makeZip({ "large.txt": "too large" }), { maxEntrySize: 2 }), /size limit/);
  const truncated = makeZip({ "data.txt": "repeated repeated repeated data" }, { deflate: true });
  const centralOffset = truncated.readUInt32LE(truncated.length - 6);
  const declaredCompressedSize = truncated.readUInt32LE(centralOffset + 20);
  truncated.writeUInt32LE(declaredCompressedSize - 1, centralOffset + 20);
  assert.throws(() => ZipArchive.open(truncated).read("data.txt"), /Truncated DEFLATE|output size/);
});
