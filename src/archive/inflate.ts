import { PreviewError } from "../errors.js";

class Bits {
  #bit = 0;
  constructor(private readonly data: Uint8Array) {}
  read(count: number): number {
    if (count < 0 || count > 24 || this.#bit + count > this.data.length * 8) throw new PreviewError("Truncated DEFLATE stream.", "invalid-file");
    let value = 0;
    for (let i = 0; i < count; i++, this.#bit++) value |= ((this.data[this.#bit >>> 3]! >>> (this.#bit & 7)) & 1) << i;
    return value;
  }
  align(): void { this.#bit = (this.#bit + 7) & ~7; }
}

type Huffman = { byLength: Array<Map<number, number>>; maxLength: number };
function makeTree(lengths: number[]): Huffman {
  const maxLength = Math.max(...lengths);
  if (!maxLength || maxLength > 15) throw new PreviewError("Invalid DEFLATE Huffman code length.", "invalid-file");
  const counts = new Array<number>(maxLength + 1).fill(0);
  for (const length of lengths) if (length > 0) counts[length] = counts[length]! + 1;
  let code = 0;
  const next = new Array<number>(maxLength + 1).fill(0);
  for (let bits = 1; bits <= maxLength; bits++) { code = (code + (counts[bits - 1] ?? 0)) << 1; next[bits] = code; }
  const byLength = Array.from({ length: maxLength + 1 }, () => new Map<number, number>());
  lengths.forEach((length, symbol) => { if (length) { byLength[length]!.set(next[length]!, symbol); next[length] = next[length]! + 1; } });
  return { byLength, maxLength };
}
function decode(bits: Bits, tree: Huffman): number {
  let code = 0;
  for (let length = 1; length <= tree.maxLength; length++) {
    code = (code << 1) | bits.read(1);
    const symbol = tree.byLength[length]!.get(code);
    if (symbol !== undefined) return symbol;
  }
  throw new PreviewError("Invalid DEFLATE Huffman code.", "invalid-file");
}

const lengthBase = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258];
const lengthExtra = [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];
const distanceBase = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577];
const distanceExtra = [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];
const codeOrder = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];

/** RFC 1951 raw DEFLATE decoder. Output is bounded by the declared size and caller limit. */
export function inflateRaw(input: Uint8Array, expectedSize: number, maxOutputSize = 32 * 1024 * 1024): Uint8Array {
  if (!Number.isSafeInteger(expectedSize) || expectedSize < 0 || expectedSize > maxOutputSize) throw new PreviewError("DEFLATE output exceeds the configured limit.", "resource-limit");
  const bits = new Bits(input), output = new Uint8Array(expectedSize);
  let out = 0, final = false;
  const emit = (value: number) => { if (out >= expectedSize) throw new PreviewError("DEFLATE output exceeds its declared size.", "invalid-file"); output[out++] = value; };
  while (!final) {
    final = bits.read(1) === 1;
    const type = bits.read(2);
    if (type === 0) {
      bits.align(); const length = bits.read(16), complement = bits.read(16);
      if (((length ^ 0xffff) & 0xffff) !== complement) throw new PreviewError("Invalid uncompressed DEFLATE block length.", "invalid-file");
      for (let i = 0; i < length; i++) emit(bits.read(8));
      continue;
    }
    if (type === 3) throw new PreviewError("Reserved DEFLATE block type.", "invalid-file");
    let literalTree: Huffman, distanceTree: Huffman;
    if (type === 1) {
      const literalLengths = new Array<number>(288).fill(8);
      literalLengths.fill(9, 144, 256); literalLengths.fill(7, 256, 280); literalLengths.fill(8, 280, 288);
      literalTree = makeTree(literalLengths); distanceTree = makeTree(new Array<number>(32).fill(5));
    } else {
      const literalCount = bits.read(5) + 257, distanceCount = bits.read(5) + 1, codeCount = bits.read(4) + 4;
      const codeLengths = new Array<number>(19).fill(0);
      for (let i = 0; i < codeCount; i++) codeLengths[codeOrder[i]!] = bits.read(3);
      const codeTree = makeTree(codeLengths), allLengths: number[] = [];
      while (allLengths.length < literalCount + distanceCount) {
        const symbol = decode(bits, codeTree);
        if (symbol <= 15) allLengths.push(symbol);
        else if (symbol === 16) {
          if (!allLengths.length) throw new PreviewError("Invalid DEFLATE repeat code.", "invalid-file");
          const count = bits.read(2) + 3, previous = allLengths[allLengths.length - 1]!;
          for (let i = 0; i < count; i++) allLengths.push(previous);
        } else if (symbol === 17 || symbol === 18) {
          const count = bits.read(symbol === 17 ? 3 : 7) + (symbol === 17 ? 3 : 11);
          for (let i = 0; i < count; i++) allLengths.push(0);
        } else throw new PreviewError("Invalid DEFLATE code-length symbol.", "invalid-file");
        if (allLengths.length > literalCount + distanceCount) throw new PreviewError("DEFLATE code lengths exceed their declaration.", "invalid-file");
      }
      literalTree = makeTree(allLengths.slice(0, literalCount)); distanceTree = makeTree(allLengths.slice(literalCount));
    }
    while (true) {
      const symbol = decode(bits, literalTree);
      if (symbol < 256) { emit(symbol); continue; }
      if (symbol === 256) break;
      if (symbol < 257 || symbol > 285) throw new PreviewError("Invalid DEFLATE length code.", "invalid-file");
      const li = symbol - 257, length = lengthBase[li]! + bits.read(lengthExtra[li]!);
      const ds = decode(bits, distanceTree);
      if (ds > 29) throw new PreviewError("Invalid DEFLATE distance code.", "invalid-file");
      const distance = distanceBase[ds]! + bits.read(distanceExtra[ds]!);
      if (distance > out) throw new PreviewError("DEFLATE distance points before output.", "invalid-file");
      for (let i = 0; i < length; i++) emit(output[out - distance]!);
    }
  }
  if (out !== expectedSize) throw new PreviewError("DEFLATE output size does not match the ZIP declaration.", "invalid-file");
  return output;
}
