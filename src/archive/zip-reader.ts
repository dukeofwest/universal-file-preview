import { PreviewError } from "../errors.js";
import { BinaryReader } from "../binary/binary-reader.js";
import { inflateRaw } from "./inflate.js";

export interface ZipEntryInfo { readonly name: string; readonly compressedSize: number; readonly uncompressedSize: number; readonly compressionMethod: number; }
const EOCD = 0x06054b50, CENTRAL = 0x02014b50, LOCAL = 0x04034b50, DESCRIPTOR = 0x08074b50;
/** General-purpose flag bits that change how an entry is read: 0 = encrypted, 3 = sizes and CRC in a data descriptor. */
const READ_FLAGS = 0b1001;

export class ZipArchive {
  readonly #entries = new Map<string, { info: ZipEntryInfo; crc: number; flags: number; localOffset: number }>();
  private constructor(private readonly bytes: Uint8Array, private readonly maxEntrySize: number) {}

  static open(bytes: Uint8Array, options: { maxEntrySize?: number; maxEntries?: number } = {}): ZipArchive {
    const archive = new ZipArchive(bytes, options.maxEntrySize ?? 32 * 1024 * 1024);
    archive.#readDirectory(options.maxEntries ?? 4096);
    return archive;
  }

  get entries(): ZipEntryInfo[] { return [...this.#entries.values()].map(item => item.info); }
  has(path: string): boolean { return this.#entries.has(normalizePath(path)); }
  read(path: string): Uint8Array {
    const item = this.#entries.get(normalizePath(path));
    if (!item) throw new PreviewError(`ZIP entry not found: ${path}`, "invalid-file");
    const reader = new BinaryReader(this.bytes);
    reader.seek(item.localOffset);
    if (reader.readUint32LE() !== LOCAL) throw new PreviewError("Invalid ZIP local file header.", "invalid-file");
    reader.skip(2); const localFlags = reader.readUint16LE(), localMethod = reader.readUint16LE();
    reader.skip(16);
    const nameLength = reader.readUint16LE(), extraLength = reader.readUint16LE();
    if (localFlags & 1) throw new PreviewError("Encrypted ZIP entries are unsupported.", "unsupported-feature");
    if (localMethod !== item.info.compressionMethod) throw new PreviewError("ZIP local and central compression methods differ.", "invalid-file");
    reader.skip(nameLength + extraLength);
    const data = reader.readBytes(item.info.compressedSize);
    const result = item.info.compressionMethod === 0 ? data.slice() : inflateRaw(data, item.info.uncompressedSize, this.maxEntrySize);
    if (result.length !== item.info.uncompressedSize) throw new PreviewError("ZIP entry size does not match its declaration.", "invalid-file");
    if (crc32(result) !== item.crc) throw new PreviewError(`ZIP checksum mismatch for ${item.info.name}.`, "invalid-file");
    return result;
  }

  #readDirectory(maxEntries: number): void {
    if (this.bytes.length < 22) throw new PreviewError("ZIP file is too short.", "invalid-file");
    const min = Math.max(0, this.bytes.length - 65557);
    let eocd = -1;
    for (let at = this.bytes.length - 22; at >= min; at--) {
      if (readU32(this.bytes, at) === EOCD) {
        const commentLength = readU16(this.bytes, at + 20);
        if (at + 22 + commentLength === this.bytes.length) { eocd = at; break; }
      }
    }
    if (eocd < 0) throw new PreviewError("ZIP end-of-central-directory record was not found.", "invalid-file");
    const end = new BinaryReader(this.bytes); end.seek(eocd + 4);
    const disk = end.readUint16LE(), directoryDisk = end.readUint16LE();
    const diskEntries = end.readUint16LE(), totalEntries = end.readUint16LE();
    const directorySize = end.readUint32LE(), directoryOffset = end.readUint32LE();
    if (disk || directoryDisk || diskEntries !== totalEntries) throw new PreviewError("Multi-disk ZIP archives are unsupported.", "unsupported-feature");
    if (totalEntries === 0xffff || directorySize === 0xffffffff || directoryOffset === 0xffffffff) throw new PreviewError("ZIP64 archives are unsupported.", "unsupported-feature");
    if (totalEntries > maxEntries || directoryOffset + directorySize > eocd) throw new PreviewError("Invalid or oversized ZIP central directory.", "resource-limit");
    const reader = new BinaryReader(this.bytes); reader.seek(directoryOffset);
    for (let i = 0; i < totalEntries; i++) {
      if (reader.readUint32LE() !== CENTRAL) throw new PreviewError("Invalid ZIP central-directory entry.", "invalid-file");
      reader.skip(4);
      const flags = reader.readUint16LE(), method = reader.readUint16LE();
      reader.skip(4);
      const crc = reader.readUint32LE(), compressedSize = reader.readUint32LE(), uncompressedSize = reader.readUint32LE();
      const nameLength = reader.readUint16LE(), extraLength = reader.readUint16LE(), commentLength = reader.readUint16LE();
      const diskStart = reader.readUint16LE(); reader.skip(6);
      const localOffset = reader.readUint32LE();
      if (diskStart !== 0 || compressedSize === 0xffffffff || uncompressedSize === 0xffffffff || localOffset === 0xffffffff) throw new PreviewError("ZIP64 or multi-disk entry is unsupported.", "unsupported-feature");
      if (flags & 1) throw new PreviewError("Encrypted ZIP entries are unsupported.", "unsupported-feature");
      if (method !== 0 && method !== 8) throw new PreviewError(`ZIP compression method ${method} is unsupported.`, "unsupported-feature");
      if (uncompressedSize > this.maxEntrySize || compressedSize > this.maxEntrySize) throw new PreviewError("ZIP entry exceeds the configured size limit.", "resource-limit");
      const nameBytes = reader.readBytes(nameLength);
      let name: string;
      try { name = new TextDecoder("utf-8", { fatal: true }).decode(nameBytes); }
      catch { throw new PreviewError("ZIP entry name is not valid UTF-8.", "invalid-file"); }
      reader.skip(extraLength + commentLength);
      if (!name || name.includes("\\") || name.split("/").includes("..")) throw new PreviewError("Unsafe ZIP package path.", "invalid-file");
      if (this.#entries.has(name)) throw new PreviewError(`Duplicate ZIP entry: ${name}`, "invalid-file");
      this.#entries.set(name, { info: { name, compressedSize, uncompressedSize, compressionMethod: method }, crc, flags, localOffset });
    }
    if (reader.offset !== directoryOffset + directorySize) throw new PreviewError("ZIP central-directory size is inconsistent.", "invalid-file");
    this.#checkLocalRecords(directoryOffset);
  }

  /**
   * Checked for every entry when the archive is opened, whether or not it is ever read: the local
   * header must agree with the central-directory record (name, method, the encryption and
   * data-descriptor flags, and the CRC and sizes, taken from the data descriptor when flag bit 3 is
   * set), and no entry's record (local header, data and data descriptor) may overlap another or the
   * central directory. Overlapping entries let a small file declare far more data than it contains,
   * and records that disagree show different files to different readers.
   */
  #checkLocalRecords(directoryOffset: number): void {
    const ranges: Array<[number, number]> = [];
    const u32 = (offset: number) => readU32(this.bytes, offset);
    for (const { info, crc, flags: centralFlags, localOffset } of this.#entries.values()) {
      if (localOffset + 30 > directoryOffset || u32(localOffset) !== LOCAL) throw new PreviewError("Invalid ZIP local file header.", "invalid-file");
      const flags = readU16(this.bytes, localOffset + 6), method = readU16(this.bytes, localOffset + 8);
      if ((flags & READ_FLAGS) !== (centralFlags & READ_FLAGS)) throw new PreviewError("ZIP local and central entry flags differ.", "invalid-file");
      const nameLength = readU16(this.bytes, localOffset + 26), extraLength = readU16(this.bytes, localOffset + 28);
      const dataStart = localOffset + 30 + nameLength + extraLength, dataEnd = dataStart + info.compressedSize;
      if (dataEnd > directoryOffset) throw new PreviewError("ZIP entry data overlaps the central directory.", "invalid-file");
      if (method !== info.compressionMethod) throw new PreviewError("ZIP local and central compression methods differ.", "invalid-file");
      let localName: string;
      try { localName = new TextDecoder("utf-8", { fatal: true }).decode(this.bytes.subarray(localOffset + 30, localOffset + 30 + nameLength)); }
      catch { throw new PreviewError("ZIP entry name is not valid UTF-8.", "invalid-file"); }
      if (localName !== info.name) throw new PreviewError("ZIP local and central entry names differ.", "invalid-file");
      const matches = (at: number) => u32(at) === crc && u32(at + 4) === info.compressedSize && u32(at + 8) === info.uncompressedSize;
      let recordEnd = dataEnd;
      if (flags & 8) {
        // The sizes and CRC follow the data, with or without the optional descriptor signature; the
        // local header holds zeros or the same values.
        const signed = dataEnd + 16 <= directoryOffset && u32(dataEnd) === DESCRIPTOR && matches(dataEnd + 4);
        if (signed) recordEnd = dataEnd + 16;
        else if (dataEnd + 12 <= directoryOffset && matches(dataEnd)) recordEnd = dataEnd + 12;
        else throw new PreviewError("ZIP data descriptor is missing or differs from the central directory.", "invalid-file");
        const local = [u32(localOffset + 14), u32(localOffset + 18), u32(localOffset + 22)];
        if (!(local.every(value => value === 0) || matches(localOffset + 14))) throw new PreviewError("ZIP local and central entry sizes differ.", "invalid-file");
      } else if (!matches(localOffset + 14)) throw new PreviewError("ZIP local and central entry sizes differ.", "invalid-file");
      ranges.push([localOffset, recordEnd]);
    }
    ranges.sort((a, b) => a[0] - b[0]);
    for (let index = 1; index < ranges.length; index++) if (ranges[index]![0] < ranges[index - 1]![1]) throw new PreviewError("ZIP entries overlap.", "invalid-file");
  }
}

function normalizePath(path: string): string { return path.replace(/^\/+/, "").replace(/\/+/g, "/"); }
function readU16(bytes: Uint8Array, offset: number): number { return bytes[offset]! | (bytes[offset + 1]! << 8); }
function readU32(bytes: Uint8Array, offset: number): number { return (bytes[offset]! | (bytes[offset + 1]! << 8) | (bytes[offset + 2]! << 16) | (bytes[offset + 3]! << 24)) >>> 0; }
/** CRC-32 (ISO-HDLC, as used by ZIP), one table lookup per byte. */
const crcTable = (() => {
  const table = new Int32Array(256);
  for (let byte = 0; byte < 256; byte++) {
    let crc = byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    table[byte] = crc;
  }
  return table;
})();
export function crc32(bytes: Uint8Array): number {
  let crc = -1;
  for (let index = 0; index < bytes.length; index++) crc = crcTable[(crc ^ bytes[index]!) & 0xff]! ^ (crc >>> 8);
  return (crc ^ -1) >>> 0;
}
