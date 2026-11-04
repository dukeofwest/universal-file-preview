import test from "node:test";
import assert from "node:assert/strict";
import { constants, crc32 as zlibCrc32, deflateRawSync } from "node:zlib";
import { inflateRaw } from "../dist/archive/inflate.js";
import { crc32 } from "../dist/archive/zip-reader.js";
import { PreviewError } from "../dist/errors.js";

let seed = 0x9e3779b9;
const random = () => (seed = (seed * 1103515245 + 12345) >>> 0) / 4294967296;
const fails = (run, message) => assert.throws(run, error => error instanceof PreviewError && error.code === "invalid-file" && message.test(error.message), String(message));

/** Writes DEFLATE bit streams: plain values least-significant bit first, Huffman codes most-significant bit first. */
class BitWriter {
  #bytes = []; #byte = 0; #count = 0;
  bits(value, count) {
    for (let bit = 0; bit < count; bit++) {
      this.#byte |= ((value >>> bit) & 1) << this.#count;
      if (++this.#count === 8) { this.#bytes.push(this.#byte); this.#byte = 0; this.#count = 0; }
    }
    return this;
  }
  code(code, length) { for (let bit = length - 1; bit >= 0; bit--) this.bits((code >>> bit) & 1, 1); return this; }
  /** Pads to a byte boundary, as a stored block requires, then appends raw bytes. */
  raw(...bytes) { if (this.#count) { this.#bytes.push(this.#byte); this.#byte = 0; this.#count = 0; } this.#bytes.push(...bytes); return this; }
  done() { return Uint8Array.from(this.#count ? [...this.#bytes, this.#byte] : this.#bytes); }
}
/** Fixed Huffman literal/length code (RFC 1951 §3.2.6). */
const fixedLiteral = (writer, symbol) => symbol < 144 ? writer.code(0x30 + symbol, 8) : symbol < 256 ? writer.code(0x190 + symbol - 144, 9) : symbol < 280 ? writer.code(symbol - 256, 7) : writer.code(0xc0 + symbol - 280, 8);
const fixedBlock = () => new BitWriter().bits(1, 1).bits(1, 2);
/** A dynamic block header with only code-length codes for the symbols 16, 17, 18 and 0 (in that transmission order). */
const dynamicHeader = (lengths16, lengths17, lengths18, lengths0) => new BitWriter().bits(1, 1).bits(2, 2).bits(0, 5).bits(0, 5).bits(0, 4).bits(lengths16, 3).bits(lengths17, 3).bits(lengths18, 3).bits(lengths0, 3);

// -----------------------------------------------------------------------------------------------
// CRC-32
// -----------------------------------------------------------------------------------------------

function bitwiseCrc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) { crc ^= byte; for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0); }
  return (crc ^ 0xffffffff) >>> 0;
}

test("table CRC-32 equals the bitwise CRC-32 and zlib for random buffers", () => {
  assert.equal(crc32(new TextEncoder().encode("123456789")), 0xcbf43926, "standard check value");
  assert.equal(crc32(new Uint8Array(0)), 0);
  for (let index = 0; index < 300; index++) {
    const bytes = Uint8Array.from({ length: Math.floor(random() ** 2 * 70_000) }, () => Math.floor(random() * 256));
    const expected = bitwiseCrc32(bytes);
    assert.equal(crc32(bytes), expected, `buffer ${index} (${bytes.length} bytes)`);
    assert.equal(crc32(bytes), zlibCrc32(bytes) >>> 0);
  }
  for (const byte of [0, 0x80, 0xff]) assert.equal(crc32(new Uint8Array(4096).fill(byte)), bitwiseCrc32(new Uint8Array(4096).fill(byte)));
});

// -----------------------------------------------------------------------------------------------
// Differential decoding against zlib
// -----------------------------------------------------------------------------------------------

const words = Array.from({ length: 500 }, () => Array.from({ length: 2 + Math.floor(random() * 9) }, () => String.fromCharCode(97 + Math.floor(random() * 26))).join(""));
const generators = [
  length => Uint8Array.from({ length }, () => Math.floor(random() * 256)),
  length => { let text = ""; while (text.length < length) text += words[Math.floor(random() * words.length)] + (random() < 0.1 ? "\n" : " "); return new TextEncoder().encode(text.slice(0, length)); },
  length => { let text = ""; for (let i = 0; text.length < length; i++) text += `<c r="A${i}" s="${i % 7}"><v>${(i * 7919) % 100_003}.${i % 97}</v></c>`; return new TextEncoder().encode(text.slice(0, length)); },
  length => new Uint8Array(length),
  length => new Uint8Array(length).fill(0x41),
  length => Uint8Array.from({ length }, (_, i) => i % 3 === 0 ? 0x61 : 0x62),
  length => Uint8Array.from({ length }, (_, i) => (i * 31 + (i >> 8)) & 0xff),
  length => Uint8Array.from({ length }, (_, i) => (i >> 12) % 2 ? Math.floor(random() * 256) : 0x20 + (i % 64)),
];

test("inflateRaw reproduces 200 random and structured buffers compressed by zlib at levels 0, 1, 6 and 9", () => {
  let checked = 0, bytes = 0;
  for (let index = 0; index < 200; index++) {
    const length = index < 16 ? index : Math.round(Math.exp(Math.log(4 * 1024 * 1024) * ((index - 16) / 183) ** 2));
    const data = generators[index % generators.length](length);
    for (const level of [0, 1, 6, 9]) {
      const compressed = new Uint8Array(deflateRawSync(data, { level }));
      assert.deepEqual(inflateRaw(compressed, data.length), data, `buffer ${index} (${length} bytes, generator ${index % generators.length}), level ${level}`);
      checked++; bytes += data.length;
    }
  }
  assert.equal(checked, 800);
  console.log(`# inflateRaw differential: ${checked} streams, ${(bytes / 1048576).toFixed(1)} MiB decoded`);
});

test("inflateRaw reproduces zlib output for every strategy and small windows", () => {
  const strategies = { default: constants.Z_DEFAULT_STRATEGY, filtered: constants.Z_FILTERED, huffmanOnly: constants.Z_HUFFMAN_ONLY, rle: constants.Z_RLE, fixed: constants.Z_FIXED };
  for (const [name, strategy] of Object.entries(strategies)) {
    for (const windowBits of [9, 12, 15]) {
      for (const [index, generate] of generators.entries()) {
        const data = generate(20_000 + index * 3_000);
        const compressed = new Uint8Array(deflateRawSync(data, { strategy, windowBits, level: 6 }));
        assert.deepEqual(inflateRaw(compressed, data.length), data, `${name}, windowBits ${windowBits}, generator ${index}`);
      }
    }
  }
});

// -----------------------------------------------------------------------------------------------
// Hand-built streams for every decoder path and error
// -----------------------------------------------------------------------------------------------

test("hand-built fixed and stored blocks decode", () => {
  const writer = fixedBlock();
  for (const symbol of [65, 66, 200]) fixedLiteral(writer, symbol);
  fixedLiteral(writer, 258).code(0, 5); // length 4 (symbol 258), distance 1 (code 0): overlapping copy
  fixedLiteral(writer, 256);
  assert.deepEqual([...inflateRaw(writer.done(), 7)], [65, 66, 200, 200, 200, 200, 200]);
  assert.deepEqual([...inflateRaw(Uint8Array.of(0x01, 0x03, 0x00, 0xfc, 0xff, 7, 8, 9), 3)], [7, 8, 9]);
  assert.deepEqual([...inflateRaw(Uint8Array.of(0x01, 0x00, 0x00, 0xff, 0xff), 0)], []);
  // A non-final fixed block followed by a final stored block that starts mid-byte: the decoder
  // must align to the next byte boundary before reading LEN/NLEN.
  const stream = new BitWriter().bits(0, 1).bits(1, 2);
  fixedLiteral(stream, 0x5a); fixedLiteral(stream, 256);
  stream.bits(1, 1).bits(0, 2).raw(0x01, 0x00, 0xfe, 0xff, 0x7a);
  assert.deepEqual([...inflateRaw(stream.done(), 2)], [0x5a, 0x7a]);
});

test("block-level errors keep their messages", () => {
  fails(() => inflateRaw(new BitWriter().bits(1, 1).bits(3, 2).done(), 1), /^Reserved DEFLATE block type\.$/);
  fails(() => inflateRaw(Uint8Array.of(0x01, 0x03, 0x00, 0x00, 0x00, 1, 2, 3), 3), /^Invalid uncompressed DEFLATE block length\.$/);
  fails(() => inflateRaw(Uint8Array.of(0x01, 0x05, 0x00, 0xfa, 0xff, 1, 2), 5), /^Truncated DEFLATE stream\.$/);
  fails(() => inflateRaw(Uint8Array.of(0x01, 0x05, 0x00, 0xfa, 0xff, 1, 2, 3, 4, 5), 3), /^DEFLATE output exceeds its declared size\.$/);
  fails(() => inflateRaw(new Uint8Array(deflateRawSync(Buffer.from("abc"))), 5), /^DEFLATE output size does not match the ZIP declaration\.$/);
  fails(() => inflateRaw(new Uint8Array(deflateRawSync(Buffer.from("abcdef"))), 3), /^DEFLATE output exceeds its declared size\.$/);
  fails(() => inflateRaw(new Uint8Array(0), 0), /^Truncated DEFLATE stream\.$/);
  assert.throws(() => inflateRaw(new Uint8Array(1), 33 * 1024 * 1024), error => error.code === "resource-limit" && /^DEFLATE output exceeds the configured limit\.$/.test(error.message));
  assert.throws(() => inflateRaw(new Uint8Array(1), -1), error => error.code === "resource-limit");
});

test("compressed-data errors keep their messages", () => {
  fails(() => inflateRaw(fixedLiteral(fixedBlock(), 257).code(0, 5).done(), 3), /^DEFLATE distance points before output\.$/);
  fails(() => inflateRaw(fixedLiteral(fixedLiteral(fixedBlock(), 65), 257).code(30, 5).done(), 4), /^Invalid DEFLATE distance code\.$/);
  fails(() => inflateRaw(fixedLiteral(fixedBlock(), 286).done(), 3), /^Invalid DEFLATE length code\.$/);
  fails(() => inflateRaw(dynamicHeader(1, 0, 0, 1).code(1, 1).done(), 1), /^Invalid DEFLATE repeat code\.$/);
  fails(() => inflateRaw(dynamicHeader(0, 0, 1, 1).code(1, 1).bits(127, 7).code(1, 1).bits(127, 7).done(), 1), /^DEFLATE code lengths exceed their declaration\.$/);
  // All-zero code lengths remain invalid, as before.
  fails(() => inflateRaw(dynamicHeader(0, 0, 0, 0).done(), 1), /^Invalid DEFLATE Huffman code length\.$/);
});

test("over-subscribed Huffman codes are rejected (RFC 1951; the previous decoder accepted them)", () => {
  // Three one-bit codes cannot exist.
  fails(() => inflateRaw(dynamicHeader(1, 1, 1, 0).code(0, 1).done(), 1), /^Invalid DEFLATE Huffman code length\.$/);
});

test("every truncation of a valid stream reports a truncated stream", () => {
  const data = generators[1](30_000);
  for (const level of [1, 9]) {
    const compressed = new Uint8Array(deflateRawSync(data, { level }));
    for (let cut = 0; cut < compressed.length; cut += Math.max(1, Math.floor(compressed.length / 400))) fails(() => inflateRaw(compressed.subarray(0, cut), data.length), /^Truncated DEFLATE stream\.$/);
    fails(() => inflateRaw(compressed.subarray(0, compressed.length - 1), data.length), /^Truncated DEFLATE stream\.$/);
  }
});

// -----------------------------------------------------------------------------------------------
// Performance guard (logged; catches order-of-magnitude regressions only)
// -----------------------------------------------------------------------------------------------

test("inflating a 32 MiB text entry stays well under two seconds", () => {
  const data = generators[2](32 * 1024 * 1024 - 1), compressed = new Uint8Array(deflateRawSync(data, { level: 6 }));
  const started = performance.now(), result = inflateRaw(compressed, data.length), elapsed = performance.now() - started;
  assert.equal(Buffer.compare(Buffer.from(result.buffer), Buffer.from(data.buffer, data.byteOffset, data.byteLength)), 0);
  console.log(`# inflateRaw: 32 MiB from ${(compressed.length / 1048576).toFixed(1)} MiB in ${elapsed.toFixed(0)} ms`);
  assert.ok(elapsed < 2000, `took ${elapsed.toFixed(0)} ms`);
});
