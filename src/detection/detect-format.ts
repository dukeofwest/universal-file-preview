export type FileFormat = "pdf" | "docx" | "xlsx" | "pptx" | "doc" | "xls" | "ppt" | "gif" | "png" | "jpeg" | "webp" | "svg" | "ods" | "ots" | "fods" | "unknown";

const extensions: Readonly<Record<string, FileFormat>> = {
  pdf: "pdf", docx: "docx", xlsx: "xlsx", pptx: "pptx", doc: "doc", xls: "xls", ppt: "ppt",
  gif: "gif", png: "png", jpg: "jpeg", jpeg: "jpeg", webp: "webp", svg: "svg",
  ods: "ods", ots: "ots", fods: "fods",
};

const ODS_MEDIA_TYPE = "application/vnd.oasis.opendocument.spreadsheet";
const OTS_MEDIA_TYPE = "application/vnd.oasis.opendocument.spreadsheet-template";
const ODF_OFFICE_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

const byMime: Readonly<Record<string, FileFormat>> = {
  "application/pdf": "pdf", "image/gif": "gif", "image/png": "png", "image/jpeg": "jpeg",
  "image/webp": "webp", "image/svg+xml": "svg", "application/msword": "doc",
  "application/vnd.ms-excel": "xls", "application/vnd.ms-powerpoint": "ppt",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
  [ODS_MEDIA_TYPE]: "ods", [OTS_MEDIA_TYPE]: "ots",
};

/** OpenDocument media types this library can identify inside a package's mimetype entry. */
const odfPackageFormats: Readonly<Record<string, FileFormat>> = { [ODS_MEDIA_TYPE]: "ods", [OTS_MEDIA_TYPE]: "ots" };

export function detectFormat(name: string, mimeType = "", bytes?: Uint8Array): FileFormat {
  const mime = mimeType.toLowerCase().split(";")[0]?.trim() ?? "";
  if (bytes) {
    // Signatures at offset 0 decide first; a "%PDF-" marker later in the first KiB (which PDF readers tolerate)
    // counts only when nothing else matched, so a ZIP entry name or image data containing it cannot change the format.
    if (starts(bytes, [0x25, 0x50, 0x44, 0x46, 0x2d])) return "pdf";
    if (starts(bytes, [0x47, 0x49, 0x46, 0x38])) return "gif";
    if (starts(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return "png";
    if (starts(bytes, [0xff, 0xd8, 0xff])) return "jpeg";
    if (starts(bytes, [0x52, 0x49, 0x46, 0x46]) && ascii(bytes, 8, 4) === "WEBP") return "webp";
    if (starts(bytes, [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])) return fromZipName(name) === "unknown" ? lookup(byMime, mime) ?? "unknown" : fromZipName(name);
    if (starts(bytes, [0x50, 0x4b, 0x03, 0x04])) {
      const packageFormat = lookup(odfPackageFormats, sniffOdfMimetype(bytes) ?? "");
      if (packageFormat) return packageFormat;
      return fromZipName(name) === "unknown" ? lookup(byMime, mime) ?? "unknown" : fromZipName(name);
    }
    if ((extensionOf(name) === "fods" || mime === ODS_MEDIA_TYPE) && isFlatSpreadsheet(bytes)) return "fods";
    const prefix = new TextDecoder().decode(bytes.subarray(0, Math.min(bytes.length, 256))).trimStart();
    if (/^(<\?xml[^>]*>\s*)?<svg[\s>]/i.test(prefix)) return "svg";
    if (ascii(bytes, 0, Math.min(bytes.length, 1024)).includes("%PDF-")) return "pdf";
  }
  return lookup(byMime, mime) ?? lookup(extensions, extensionOf(name)) ?? "unknown";
}

/**
 * Reads the media type from an OpenDocument package's first ZIP entry (ODF 1.3 Part 2 §3.3:
 * a STORED file named "mimetype" without a data descriptor). Returns undefined for anything
 * else; never throws and never decompresses.
 */
function sniffOdfMimetype(bytes: Uint8Array): string | undefined {
  if (bytes.length < 38 || !starts(bytes, [0x50, 0x4b, 0x03, 0x04])) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const flags = view.getUint16(6, true), method = view.getUint16(8, true);
  const compressedSize = view.getUint32(18, true), uncompressedSize = view.getUint32(22, true);
  const nameLength = view.getUint16(26, true), extraLength = view.getUint16(28, true);
  if (method !== 0 || flags & 0x8 || nameLength !== 8 || extraLength > 1024) return undefined;
  if (compressedSize < 1 || compressedSize > 256 || uncompressedSize !== compressedSize) return undefined;
  if (ascii(bytes, 30, 8) !== "mimetype") return undefined;
  const start = 38 + extraLength, end = start + compressedSize;
  if (end > bytes.length) return undefined;
  for (let index = start; index < end; index++) {
    const byte = bytes[index]!;
    if (byte > 0x7e || (byte < 0x20 && byte !== 0x09 && byte !== 0x0a && byte !== 0x0d)) return undefined;
  }
  return ascii(bytes, start, compressedSize).trim();
}

/** A flat OpenDocument spreadsheet declares the office namespace and the spreadsheet media type near its start. */
function isFlatSpreadsheet(bytes: Uint8Array): boolean {
  const text = new TextDecoder().decode(bytes.subarray(0, Math.min(bytes.length, 4096)));
  return text.includes(ODF_OFFICE_NAMESPACE) && text.includes(ODS_MEDIA_TYPE);
}

/** Own-property lookup: names and MIME types are untrusted and must not resolve inherited properties such as "constructor". */
function lookup(table: Readonly<Record<string, FileFormat>>, key: string): FileFormat | undefined { return Object.hasOwn(table, key) ? table[key] : undefined; }
function extensionOf(name: string): string { return name.toLowerCase().split(".").pop() ?? ""; }
function fromZipName(name: string): FileFormat { return lookup(extensions, extensionOf(name)) ?? "unknown"; }
function starts(bytes: Uint8Array, sig: number[]): boolean { return sig.every((value, i) => bytes[i] === value); }
function ascii(bytes: Uint8Array, offset: number, length: number): string { return String.fromCharCode(...bytes.subarray(offset, offset + length)); }
