export type FileFormat = "pdf" | "docx" | "xlsx" | "pptx" | "doc" | "xls" | "ppt" | "gif" | "png" | "jpeg" | "webp" | "svg" | "unknown";

const extensions: Record<string, FileFormat> = {
  pdf: "pdf", docx: "docx", xlsx: "xlsx", pptx: "pptx", doc: "doc", xls: "xls", ppt: "ppt",
  gif: "gif", png: "png", jpg: "jpeg", jpeg: "jpeg", webp: "webp", svg: "svg",
};

export function detectFormat(name: string, mimeType = "", bytes?: Uint8Array): FileFormat {
  const mime = mimeType.toLowerCase().split(";")[0]?.trim();
  const byMime: Record<string, FileFormat> = {
    "application/pdf": "pdf", "image/gif": "gif", "image/png": "png", "image/jpeg": "jpeg",
    "image/webp": "webp", "image/svg+xml": "svg", "application/msword": "doc",
    "application/vnd.ms-excel": "xls", "application/vnd.ms-powerpoint": "ppt",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
  };
  if (bytes) {
    const header = ascii(bytes, 0, Math.min(bytes.length, 1024));
    if (header.includes("%PDF-")) return "pdf";
    if (starts(bytes, [0x47, 0x49, 0x46, 0x38])) return "gif";
    if (starts(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return "png";
    if (starts(bytes, [0xff, 0xd8, 0xff])) return "jpeg";
    if (starts(bytes, [0x52, 0x49, 0x46, 0x46]) && ascii(bytes, 8, 4) === "WEBP") return "webp";
    if (starts(bytes, [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])) return fromZipName(name) === "unknown" ? byMime[mime ?? ""] ?? "unknown" : fromZipName(name);
    if (starts(bytes, [0x50, 0x4b, 0x03, 0x04])) return fromZipName(name) === "unknown" ? byMime[mime ?? ""] ?? "unknown" : fromZipName(name);
    const prefix = new TextDecoder().decode(bytes.subarray(0, Math.min(bytes.length, 256))).trimStart();
    if (/^(<\?xml[^>]*>\s*)?<svg[\s>]/i.test(prefix)) return "svg";
  }
  return byMime[mime ?? ""] ?? extensions[name.toLowerCase().split(".").pop() ?? ""] ?? "unknown";
}

function fromZipName(name: string): FileFormat { return extensions[name.toLowerCase().split(".").pop() ?? ""] ?? "unknown"; }
function starts(bytes: Uint8Array, sig: number[]): boolean { return sig.every((value, i) => bytes[i] === value); }
function ascii(bytes: Uint8Array, offset: number, length: number): string { return String.fromCharCode(...bytes.subarray(offset, offset + length)); }
