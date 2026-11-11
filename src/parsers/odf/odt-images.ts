import type { OdfPackageInfo } from "../../package/odf-package.js";
import type { PackageReader } from "../../package/package-reader.js";

/**
 * Resolving the images a text document refers to (ODF 1.3 Part 3 §10.4).
 *
 * Internal module: not exported from the package entry point (decision D2). Only entries of the
 * document's own package are read. A reference with a scheme, an absolute path or a parent-directory
 * step points outside the document, and this library never follows or fetches one.
 */

/**
 * Image media types this library shows. SVG is included because every SVG is reduced to the
 * sanitizer's allowlist before a blob URL exists for it. WMF, EMF and SVM cannot be read here, so a
 * document using them is told that the image is not shown rather than being left with a gap.
 */
const IMAGE_TYPES: Readonly<Record<string, string>> = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", jpe: "image/jpeg",
  gif: "image/gif", webp: "image/webp", bmp: "image/bmp", svg: "image/svg+xml",
};
const KNOWN_TYPES = new Set(Object.values(IMAGE_TYPES));

export interface PackageImage { readonly bytes: Uint8Array; readonly mimeType: string }

/** The bytes and media type of a package entry a draw:image refers to, or undefined. */
export function resolvePackageImage(reader: PackageReader, info: OdfPackageInfo, href: string): PackageImage | undefined {
  const path = safePath(href);
  if (path === undefined || !reader.has(path)) return undefined;
  const entry = info.manifest.get(path);
  if (entry?.encrypted) return undefined;
  const declared = entry?.mediaType ?? "";
  const extension = path.toLowerCase().split(".").pop() ?? "";
  // The manifest's media type is used only when this library shows that type; otherwise the extension
  // decides. Either way createImageUrl checks the bytes themselves before any URL exists.
  const mimeType = KNOWN_TYPES.has(declared) ? declared : Object.hasOwn(IMAGE_TYPES, extension) ? IMAGE_TYPES[extension] : undefined;
  if (mimeType === undefined) return undefined;
  return { bytes: reader.readBytes(path), mimeType };
}

/** The entry path a reference denotes, or undefined when it does not denote one in this package. */
export function safePath(href: string): string | undefined {
  if (!href || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(href)) return undefined;
  if (href.startsWith("/") || href.startsWith("\\") || href.includes("#") || href.includes("?")) return undefined;
  const path = href.replace(/^\.\//, "");
  if (!path || path.split(/[/\\]/).some(part => part === "" || part === "." || part === "..")) return undefined;
  return path;
}


/** The media type the bytes themselves declare, so an inline image carries an honest type. */
export function sniffImageType(bytes: Uint8Array): string | undefined {
  const starts = (...signature: number[]) => signature.every((value, index) => bytes[index] === value);
  if (starts(0x89, 0x50, 0x4e, 0x47)) return "image/png";
  if (starts(0xff, 0xd8, 0xff)) return "image/jpeg";
  if (starts(0x47, 0x49, 0x46, 0x38)) return "image/gif";
  if (starts(0x52, 0x49, 0x46, 0x46)) return "image/webp";
  if (starts(0x42, 0x4d)) return "image/bmp";
  const head = new TextDecoder().decode(bytes.subarray(0, Math.min(bytes.length, 256))).trimStart();
  if (/^(<\?xml[^>]*>\s*)?<svg[\s>]/i.test(head)) return "image/svg+xml";
  return undefined;
}

/**
 * Decodes base64 image data a flat document stores inline. The encoded length is checked before
 * anything is decoded, because decoding is where the memory is spent.
 */
export function decodeBase64(text: string, maxBytes: number): Uint8Array | undefined {
  if (!text || text.length / 4 * 3 > maxBytes) return undefined;
  const cleaned = text.replace(/[^A-Za-z0-9+/=]/g, "");
  if (!cleaned || cleaned.length % 4 !== 0 || cleaned.length / 4 * 3 > maxBytes) return undefined;
  try {
    const binary = atob(cleaned);
    const out = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index++) out[index] = binary.charCodeAt(index);
    return out;
  } catch { return undefined; }
}
