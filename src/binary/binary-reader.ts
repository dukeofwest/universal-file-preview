import { PreviewError } from "../errors.js";

export class BinaryReader {
  readonly #view: DataView;
  #offset = 0;

  constructor(readonly bytes: Uint8Array) {
    this.#view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  get length(): number { return this.bytes.length; }
  get offset(): number { return this.#offset; }

  seek(position: number): void {
    if (!Number.isSafeInteger(position) || position < 0 || position > this.length) {
      throw new PreviewError(`Invalid byte offset ${position}.`, "invalid-file");
    }
    this.#offset = position;
  }

  skip(count: number): void { this.#require(count); this.#offset += count; }
  readUint8(): number { this.#require(1); return this.#view.getUint8(this.#offset++); }
  readInt8(): number { this.#require(1); return this.#view.getInt8(this.#offset++); }
  readUint16LE(): number { this.#require(2); const n = this.#view.getUint16(this.#offset, true); this.#offset += 2; return n; }
  readUint16BE(): number { this.#require(2); const n = this.#view.getUint16(this.#offset, false); this.#offset += 2; return n; }
  readInt16LE(): number { this.#require(2); const n = this.#view.getInt16(this.#offset, true); this.#offset += 2; return n; }
  readUint32LE(): number { this.#require(4); const n = this.#view.getUint32(this.#offset, true); this.#offset += 4; return n; }
  readUint32BE(): number { this.#require(4); const n = this.#view.getUint32(this.#offset, false); this.#offset += 4; return n; }
  readInt32LE(): number { this.#require(4); const n = this.#view.getInt32(this.#offset, true); this.#offset += 4; return n; }

  readBytes(count: number): Uint8Array {
    this.#require(count);
    const result = this.bytes.subarray(this.#offset, this.#offset + count);
    this.#offset += count;
    return result;
  }

  readString(count: number, encoding = "utf-8"): string {
    try { return new TextDecoder(encoding, { fatal: true }).decode(this.readBytes(count)); }
    catch (cause) { throw new PreviewError(`Invalid ${encoding} text.`, "invalid-file", { cause }); }
  }

  #require(count: number): void {
    if (!Number.isSafeInteger(count) || count < 0 || count > this.length - this.#offset) {
      throw new PreviewError("Unexpected end of file.", "invalid-file");
    }
  }
}
