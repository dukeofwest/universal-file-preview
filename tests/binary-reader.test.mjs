import test from "node:test";
import assert from "node:assert/strict";
import { BinaryReader } from "../dist/binary/binary-reader.js";

test("BinaryReader reads little and big endian values and checks bounds", () => {
  const reader = new BinaryReader(Uint8Array.from([1, 2, 3, 4, 65, 66]));
  assert.equal(reader.readUint16LE(), 0x0201);
  assert.equal(reader.readUint16BE(), 0x0304);
  assert.equal(reader.readString(2), "AB");
  assert.throws(() => reader.readUint8(), /Unexpected end/);
  assert.throws(() => reader.seek(-1), /Invalid byte offset/);
});
