import { PreviewError } from "../errors.js";
import { ZipArchive, type ZipEntryInfo } from "../archive/zip-reader.js";

/**
 * Budgeted reader for package parts (implementation specification §4.2, milestone M6).
 *
 * Internal module: not exported from the package entry point (decision D2). ZipArchive is
 * opened without a per-entry size cap so that large parts a parser never reads (an embedded
 * video, for example) do not make the whole package unreadable. Instead, every read is
 * checked against its part limit and a running total BEFORE anything is decompressed.
 */

export interface PackageLimits {
  /** Maximum number of entries. Default 4,096 (parity with ZipArchive). */
  readonly maxEntries?: number;
  /** Default maximum decompressed size of one part. Default 32 MiB. */
  readonly maxPartBytes?: number;
  /** Per-path overrides of maxPartBytes, for example { "content.xml": 128 MiB }. */
  readonly partBytes?: Readonly<Record<string, number>>;
  /** Maximum decompressed bytes across all reads; repeated reads count each time. Default 256 MiB. */
  readonly maxTotalBytes?: number;
}

/** ZIP stores sizes in 32 bits; ZIP64 is rejected by ZipArchive. */
const ZIP32_MAX = 0xffffffff;
const supportedEncodings = new Set(["utf-8", "utf8", "utf-16", "utf16", "utf-16le", "utf-16be"]);
const encodingDeclaration = /^<\?xml\s[^>]*?\bencoding\s*=\s*(["'])([^"']*)\1/;

export class PackageReader {
  readonly #archive: ZipArchive;
  readonly #entries: ReadonlyMap<string, ZipEntryInfo>;
  readonly #maxPartBytes: number;
  /** A Map, not an object: part names are untrusted and must not resolve inherited properties such as "constructor". */
  readonly #partBytes: ReadonlyMap<string, number>;
  readonly #maxTotalBytes: number;
  #totalBytes = 0;

  private constructor(archive: ZipArchive, limits: PackageLimits) {
    this.#archive = archive;
    this.#entries = new Map(archive.entries.map(entry => [entry.name, entry]));
    this.#maxPartBytes = limit(limits.maxPartBytes, 32 * 1024 * 1024, "maxPartBytes");
    this.#maxTotalBytes = limit(limits.maxTotalBytes, 256 * 1024 * 1024, "maxTotalBytes");
    const partBytes = new Map<string, number>();
    for (const [path, value] of Object.entries(limits.partBytes ?? {})) partBytes.set(normalize(path), limit(value, 1, `partBytes["${path}"]`));
    this.#partBytes = partBytes;
  }

  /** Opens a ZIP package. ZIP structure errors are reported as by ZipArchive. */
  static openZip(bytes: Uint8Array, limits: PackageLimits = {}): PackageReader {
    const maxEntries = limit(limits.maxEntries, 4096, "maxEntries");
    return new PackageReader(ZipArchive.open(bytes, { maxEntries, maxEntrySize: ZIP32_MAX }), limits);
  }

  /** Entries in central-directory order. */
  get entries(): readonly ZipEntryInfo[] { return [...this.#entries.values()]; }

  /** Decompressed bytes read so far. */
  get totalBytes(): number { return this.#totalBytes; }

  has(path: string): boolean { return this.#entries.has(normalize(path)); }

  entry(path: string): ZipEntryInfo | undefined { return this.#entries.get(normalize(path)); }

  /** Reads a part, enforcing its size limit and the total budget before decompressing it. */
  readBytes(path: string): Uint8Array {
    const name = normalize(path), entry = this.#entries.get(name);
    if (!entry) throw new PreviewError(`Package part not found: ${name}`, "invalid-file");
    const partLimit = this.#partBytes.get(name) ?? this.#maxPartBytes;
    if (entry.uncompressedSize > partLimit) throw new PreviewError(`Package part ${name} exceeds the configured size limit.`, "resource-limit");
    if (this.#totalBytes + entry.uncompressedSize > this.#maxTotalBytes) throw new PreviewError("Package exceeds the configured total size limit.", "resource-limit");
    this.#totalBytes += entry.uncompressedSize;
    return this.#archive.read(name);
  }

  /**
   * Reads a part as text. UTF-8 is the default; a UTF-16 byte-order mark selects UTF-16LE or
   * UTF-16BE. A declared XML encoding other than UTF-8 or UTF-16 is unsupported-feature;
   * bytes invalid in the detected encoding are invalid-file. The byte-order mark is removed.
   */
  readText(path: string): string {
    const name = normalize(path), bytes = this.readBytes(name);
    const encoding = bytes[0] === 0xff && bytes[1] === 0xfe ? "utf-16le" : bytes[0] === 0xfe && bytes[1] === 0xff ? "utf-16be" : "utf-8";
    let text: string;
    try { text = new TextDecoder(encoding, { fatal: true }).decode(bytes); }
    catch (cause) { throw new PreviewError(`Package part ${name} is not valid ${encoding.toUpperCase()} text.`, "invalid-file", { cause }); }
    const declared = encodingDeclaration.exec(text)?.[2];
    if (declared !== undefined && !supportedEncodings.has(declared.trim().toLowerCase())) throw new PreviewError(`Package part ${name} uses the unsupported encoding ${declared}.`, "unsupported-feature");
    return text;
  }
}

/** Same normalisation as ZipArchive: no leading slash, single separators. */
function normalize(path: string): string { return path.replace(/^\/+/, "").replace(/\/+/g, "/"); }

function limit(value: number | undefined, fallback: number, name: string): number {
  if (value === undefined) return fallback;
  if (!Number.isSafeInteger(value) || value < 1) throw new RangeError(`${name} must be a positive safe integer.`);
  return value;
}
