import { PreviewError } from "../errors.js";

/**
 * A canonical Huffman code as a lookup table indexed by the next `bits` input bits
 * (least-significant bit first, as DEFLATE stores codes). Each entry is
 * `symbol << 4 | codeLength`; 0 marks a bit pattern that no code uses.
 */
type Table = { readonly entries: Uint16Array; readonly bits: number; readonly mask: number };

function buildTable(lengths: ArrayLike<number>): Table {
  let maxLength = 0;
  const counts = new Array<number>(16).fill(0);
  for (let symbol = 0; symbol < lengths.length; symbol++) {
    const length = lengths[symbol]!;
    if (length > maxLength) maxLength = length;
    counts[length] = counts[length]! + 1;
  }
  if (!maxLength || maxLength > 15) throw new PreviewError("Invalid DEFLATE Huffman code length.", "invalid-file");
  counts[0] = 0;
  // An over-subscribed code (more codes than the lengths allow) is invalid (RFC 1951 §3.2.2).
  let available = 1;
  for (let length = 1; length <= maxLength; length++) {
    available = (available << 1) - counts[length]!;
    if (available < 0) throw new PreviewError("Invalid DEFLATE Huffman code length.", "invalid-file");
  }
  // Canonical first code per length (RFC 1951 §3.2.2, step 2).
  const next = new Array<number>(16).fill(0);
  for (let length = 1, code = 0; length <= maxLength; length++) { code = (code + counts[length - 1]!) << 1; next[length] = code; }
  const size = 1 << maxLength, entries = new Uint16Array(size);
  for (let symbol = 0; symbol < lengths.length; symbol++) {
    const length = lengths[symbol]!;
    if (!length) continue;
    const code = next[length]!;
    next[length] = code + 1;
    let reversed = 0;
    for (let bit = 0; bit < length; bit++) reversed |= ((code >>> bit) & 1) << (length - 1 - bit);
    for (let index = reversed; index < size; index += 1 << length) entries[index] = (symbol << 4) | length;
  }
  return { entries, bits: maxLength, mask: size - 1 };
}

let fixedTables: { literal: Table; distance: Table } | undefined;
function fixed(): { literal: Table; distance: Table } {
  if (!fixedTables) {
    const literal = new Uint8Array(288).fill(8);
    literal.fill(9, 144, 256); literal.fill(7, 256, 280); literal.fill(8, 280, 288);
    fixedTables = { literal: buildTable(literal), distance: buildTable(new Uint8Array(32).fill(5)) };
  }
  return fixedTables;
}

const lengthBase = Uint16Array.of(3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258);
const lengthExtra = Uint8Array.of(0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0);
const distanceBase = Uint16Array.of(1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577);
const distanceExtra = Uint8Array.of(0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13);
const codeOrder = Uint8Array.of(16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15);

const truncated = () => new PreviewError("Truncated DEFLATE stream.", "invalid-file");
const overflow = () => new PreviewError("DEFLATE output exceeds its declared size.", "invalid-file");
const invalidCode = () => new PreviewError("Invalid DEFLATE Huffman code.", "invalid-file");

/** Bit reader for block headers and stored blocks. The compressed-data loop copies its state into locals. */
class BitReader {
  position = 0;
  /** Unconsumed bits are the low `bitCount` bits of `bitBuffer`. */
  bitBuffer = 0;
  bitCount = 0;
  constructor(readonly input: Uint8Array) {}

  /** Reads `count` (0-16) bits, least-significant first. */
  read(count: number): number {
    while (this.bitCount < count) {
      if (this.position >= this.input.length) throw truncated();
      this.bitBuffer |= this.input[this.position++]! << this.bitCount;
      this.bitCount += 8;
    }
    const value = this.bitBuffer & ((1 << count) - 1);
    this.bitBuffer >>>= count;
    this.bitCount -= count;
    return value;
  }

  /** Decodes one symbol. Missing trailing bits read as 0, accepted only when the matched code does not need them. */
  decode(table: Table): number {
    while (this.bitCount < table.bits && this.position < this.input.length) {
      this.bitBuffer |= this.input[this.position++]! << this.bitCount;
      this.bitCount += 8;
    }
    const entry = table.entries[this.bitBuffer & table.mask]!, length = entry & 15;
    if (!length) throw this.bitCount < table.bits ? truncated() : invalidCode();
    if (length > this.bitCount) throw truncated();
    this.bitBuffer >>>= length;
    this.bitCount -= length;
    return entry >>> 4;
  }

  align(): void { this.bitBuffer >>>= this.bitCount & 7; this.bitCount -= this.bitCount & 7; }
}

/** RFC 1951 raw DEFLATE decoder. Output is bounded by the declared size and caller limit. */
export function inflateRaw(input: Uint8Array, expectedSize: number, maxOutputSize = 32 * 1024 * 1024): Uint8Array {
  if (!Number.isSafeInteger(expectedSize) || expectedSize < 0 || expectedSize > maxOutputSize) throw new PreviewError("DEFLATE output exceeds the configured limit.", "resource-limit");
  const output = new Uint8Array(expectedSize), reader = new BitReader(input), inputLength = input.length;
  let out = 0, final = false;
  while (!final) {
    final = reader.read(1) === 1;
    const type = reader.read(2);
    if (type === 0) {
      reader.align();
      const length = reader.read(16), complement = reader.read(16);
      if (((length ^ 0xffff) & 0xffff) !== complement) throw new PreviewError("Invalid uncompressed DEFLATE block length.", "invalid-file");
      // Fail exactly where byte-by-byte reading would: truncation before overflow at the same byte.
      const available = (reader.bitCount >>> 3) + inputLength - reader.position, room = expectedSize - out;
      if (length > available || length > room) throw available <= room ? truncated() : overflow();
      let remaining = length;
      while (remaining && reader.bitCount) { output[out++] = reader.bitBuffer & 0xff; reader.bitBuffer >>>= 8; reader.bitCount -= 8; remaining--; }
      output.set(input.subarray(reader.position, reader.position + remaining), out);
      out += remaining; reader.position += remaining;
      continue;
    }
    if (type === 3) throw new PreviewError("Reserved DEFLATE block type.", "invalid-file");
    let literalTable: Table, distanceTable: Table;
    if (type === 1) ({ literal: literalTable, distance: distanceTable } = fixed());
    else {
      const literalCount = reader.read(5) + 257, distanceCount = reader.read(5) + 1, codeCount = reader.read(4) + 4, total = literalCount + distanceCount;
      const codeLengths = new Uint8Array(19);
      for (let i = 0; i < codeCount; i++) codeLengths[codeOrder[i]!] = reader.read(3);
      const codeTable = buildTable(codeLengths), allLengths = new Uint8Array(total);
      let filled = 0;
      while (filled < total) {
        const symbol = reader.decode(codeTable);
        if (symbol <= 15) { allLengths[filled++] = symbol; continue; }
        let count: number, value = 0;
        if (symbol === 16) {
          if (!filled) throw new PreviewError("Invalid DEFLATE repeat code.", "invalid-file");
          count = reader.read(2) + 3; value = allLengths[filled - 1]!;
        } else if (symbol === 17 || symbol === 18) count = symbol === 17 ? reader.read(3) + 3 : reader.read(7) + 11;
        else throw new PreviewError("Invalid DEFLATE code-length symbol.", "invalid-file");
        if (filled + count > total) throw new PreviewError("DEFLATE code lengths exceed their declaration.", "invalid-file");
        allLengths.fill(value, filled, filled + count);
        filled += count;
      }
      literalTable = buildTable(allLengths.subarray(0, literalCount));
      distanceTable = buildTable(allLengths.subarray(literalCount));
    }
    out = inflateBlock(reader, output, out, literalTable, distanceTable);
  }
  if (out !== expectedSize) throw new PreviewError("DEFLATE output size does not match the ZIP declaration.", "invalid-file");
  return output;
}

/**
 * Decodes one compressed block until its end-of-block symbol and returns the new output
 * position. Reader state is held in local variables for speed and written back on return.
 */
function inflateBlock(reader: BitReader, output: Uint8Array, out: number, literalTable: Table, distanceTable: Table): number {
  const input = reader.input, inputLength = input.length, expectedSize = output.length;
  const literals = literalTable.entries, literalMask = literalTable.mask, literalBits = literalTable.bits;
  const distances = distanceTable.entries, distanceMask = distanceTable.mask, distanceBits = distanceTable.bits;
  let position = reader.position, buffer = reader.bitBuffer, count = reader.bitCount;
  try {
    while (true) {
      while (count < literalBits && position < inputLength) { buffer |= input[position++]! << count; count += 8; }
      let entry = literals[buffer & literalMask]!, length = entry & 15;
      if (!length) throw count < literalBits ? truncated() : invalidCode();
      if (length > count) throw truncated();
      buffer >>>= length; count -= length;
      const symbol = entry >>> 4;
      if (symbol < 256) {
        if (out >= expectedSize) throw overflow();
        output[out++] = symbol;
        continue;
      }
      if (symbol === 256) return out;
      if (symbol > 285) throw new PreviewError("Invalid DEFLATE length code.", "invalid-file");
      const li = symbol - 257, lengthBits = lengthExtra[li]!;
      let copyLength = lengthBase[li]!;
      if (lengthBits) {
        while (count < lengthBits) { if (position >= inputLength) throw truncated(); buffer |= input[position++]! << count; count += 8; }
        copyLength += buffer & ((1 << lengthBits) - 1); buffer >>>= lengthBits; count -= lengthBits;
      }
      while (count < distanceBits && position < inputLength) { buffer |= input[position++]! << count; count += 8; }
      entry = distances[buffer & distanceMask]!; length = entry & 15;
      if (!length) throw count < distanceBits ? truncated() : invalidCode();
      if (length > count) throw truncated();
      buffer >>>= length; count -= length;
      const ds = entry >>> 4;
      if (ds > 29) throw new PreviewError("Invalid DEFLATE distance code.", "invalid-file");
      const distBits = distanceExtra[ds]!;
      let distance = distanceBase[ds]!;
      if (distBits) {
        while (count < distBits) { if (position >= inputLength) throw truncated(); buffer |= input[position++]! << count; count += 8; }
        distance += buffer & ((1 << distBits) - 1); buffer >>>= distBits; count -= distBits;
      }
      if (distance > out) throw new PreviewError("DEFLATE distance points before output.", "invalid-file");
      if (out + copyLength > expectedSize) throw overflow();
      let from = out - distance;
      if (distance >= copyLength && copyLength > 64) { output.copyWithin(out, from, from + copyLength); out += copyLength; }
      else { const end = out + copyLength; while (out < end) output[out++] = output[from++]!; }
    }
  } finally {
    reader.position = position; reader.bitBuffer = buffer; reader.bitCount = count;
  }
}
