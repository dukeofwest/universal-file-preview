import { PreviewError } from "../errors.js";
import { parseXml, type XmlNode } from "../xml/xml-parser.js";
import { PackageReader } from "./package-reader.js";

/**
 * DOCX, PPTX and XLSX package access with a total decompression budget and per-part caching.
 *
 * Internal module (not exported from the package entry point). Office documents may refer to the
 * same part many times (an image from many relationships, a slide from many slide ids). Every part
 * is decompressed at most once and every XML part parsed at most once, so repeated references
 * cost nothing; the budget counts each part once.
 */
export const OOXML_LIMITS = {
  /** Maximum decompressed size of one part (unchanged from the earlier per-entry limit). */
  maxPartBytes: 32 * 1024 * 1024,
  /** Maximum decompressed bytes across the whole package (same as OpenDocument packages). */
  maxTotalBytes: 256 * 1024 * 1024,
  maxEntries: 4096,
} as const;
/**
 * XML parts larger than 1 MiB may expand at most 100 times their compressed size. Office XML
 * typically compresses 10-20:1; highly repetitive XML that expands further is a decompression bomb
 * whose parsing would keep the page busy even within the byte budget. Media is not checked (a blank
 * bitmap legitimately compresses very well; it is bounded by the byte budget and the pixel limits).
 */
export const MAX_XML_EXPANSION = 100, XML_EXPANSION_GRACE_BYTES = 1024 * 1024;

export class OoxmlPackage {
  readonly #reader: PackageReader;
  readonly #bytes = new Map<string, Uint8Array>();
  readonly #xml = new Map<string, XmlNode>();
  readonly #derived = new Map<string, unknown>();

  private constructor(reader: PackageReader) { this.#reader = reader; }

  static open(bytes: Uint8Array): OoxmlPackage { return new OoxmlPackage(PackageReader.openZip(bytes, OOXML_LIMITS)); }

  has(path: string): boolean { return this.#reader.has(path); }

  /** Decompressed bytes of a part; the same array is returned for every later read of that part. */
  read(path: string): Uint8Array {
    const name = normalize(path);
    let bytes = this.#bytes.get(name);
    if (!bytes) { bytes = this.#reader.readBytes(name); this.#bytes.set(name, bytes); }
    return bytes;
  }

  /** A parsed XML part (UTF-8). The tree is shared between callers and must not be modified. */
  xml(path: string): XmlNode {
    const name = normalize(path);
    let node = this.#xml.get(name);
    const entry = this.#reader.entry(name);
    if (!node && entry && entry.uncompressedSize > XML_EXPANSION_GRACE_BYTES && entry.uncompressedSize > entry.compressedSize * MAX_XML_EXPANSION) {
      throw new PreviewError(`Package part ${name} expands more than ${MAX_XML_EXPANSION} times its compressed size.`, "resource-limit");
    }
    if (!node) { node = parseXml(new TextDecoder("utf-8", { fatal: true }).decode(this.#reader.readBytes(name))); this.#xml.set(name, node); }
    return node;
  }

  /** Computes a value derived from parts (for example a converted image) once per key. */
  derived<T>(key: string, compute: () => T): T {
    if (this.#derived.has(key)) return this.#derived.get(key) as T;
    const value = compute();
    this.#derived.set(key, value);
    return value;
  }

  /** Decompressed bytes counted against the budget so far. */
  get totalBytes(): number { return this.#reader.totalBytes; }
}

function normalize(path: string): string { return path.replace(/^\/+/, "").replace(/\/+/g, "/"); }
