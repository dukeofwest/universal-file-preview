import { PreviewError } from "../errors.js";
import type { DiagnosticCollector } from "../diagnostics.js";
import { nsAttr, nsChildren, nsFirst, nsText, parseXmlNs, type NsXmlNode } from "../xml/ns-tree.js";
import { ODF } from "../parsers/odf/namespaces.js";
import type { PackageReader } from "./package-reader.js";

/**
 * OpenDocument package inspection (ODF 1.3 Part 2; implementation specification §4.2 and
 * milestone M6, rules 1-8).
 *
 * Internal module: not exported from the package entry point (decision D2). Inspection
 * validates the package contract and detects encryption, signatures and macros before any
 * document part is parsed. It never reads macro or signature contents.
 */

export interface OdfManifestEntry { readonly path: string; readonly mediaType: string; readonly encrypted: boolean; readonly version?: string; }
export interface OdfPackageInfo {
  /** From the mimetype entry, else the manifest's root ("/") entry. */
  readonly mediaType: string;
  readonly manifestVersion?: string;
  /** Keyed by manifest:full-path exactly as written. Empty when the package has no manifest. */
  readonly manifest: ReadonlyMap<string, OdfManifestEntry>;
  readonly signed: boolean;
  readonly hasMacros: boolean;
  /** meta:generator from meta.xml, when readable. */
  readonly generator?: string;
  /** "Thumbnails/thumbnail.png" when present (not used in Phase 1). */
  readonly thumbnailPath?: string;
}

const { manifest: MANIFEST, office: OFFICE, meta: META } = ODF;
const MANIFEST_PATH = "META-INF/manifest.xml";
const THUMBNAIL_PATH = "Thumbnails/thumbnail.png";
const MAX_MIMETYPE_BYTES = 256, MAX_MANIFEST_BYTES = 4 * 1024 * 1024, MAX_META_BYTES = 1024 * 1024;
/** The newest OpenDocument version this library knows (ODF 1.4, OASIS Standard 2025). */
const NEWEST_KNOWN = [1, 4] as const;

export function inspectOdfPackage(reader: PackageReader, collector: DiagnosticCollector): OdfPackageInfo {
  // Rule 1: the mimetype entry should be first and stored (Part 2 §3.3).
  let mimetype: string | undefined;
  const mimetypeEntry = reader.entry("mimetype");
  if (mimetypeEntry) {
    if (mimetypeEntry.uncompressedSize > MAX_MIMETYPE_BYTES) throw new PreviewError("The package mimetype entry is too large.", "invalid-file");
    const bytes = reader.readBytes("mimetype");
    if (bytes.some(byte => byte > 0x7e || (byte < 0x20 && byte !== 0x09 && byte !== 0x0a && byte !== 0x0d))) throw new PreviewError("The package mimetype entry is not ASCII text.", "invalid-file");
    mimetype = String.fromCharCode(...bytes).trim();
    if (!mimetype) throw new PreviewError("The package mimetype entry is empty.", "invalid-file");
    if (reader.entries[0]?.name !== "mimetype" || mimetypeEntry.compressionMethod !== 0) collector.report("odf.package.mimetype-position", { part: "mimetype" });
  }

  // Rules 2 and 3: the manifest (Part 2 §3.2).
  const manifest = new Map<string, OdfManifestEntry>();
  let manifestVersion: string | undefined;
  if (reader.has(MANIFEST_PATH)) {
    if ((reader.entry(MANIFEST_PATH)?.uncompressedSize ?? 0) > MAX_MANIFEST_BYTES) throw new PreviewError(`Package part ${MANIFEST_PATH} exceeds the configured size limit.`, "resource-limit");
    const root = parsePart(reader, MANIFEST_PATH);
    if (root.uri !== MANIFEST || root.local !== "manifest") throw new PreviewError(`${MANIFEST_PATH} is not an OpenDocument manifest.`, "invalid-file");
    manifestVersion = nsAttr(root, MANIFEST, "version");
    for (const fileEntry of nsChildren(root, MANIFEST, "file-entry")) {
      const path = nsAttr(fileEntry, MANIFEST, "full-path");
      if (path === undefined || manifest.has(path)) continue;
      const version = nsAttr(fileEntry, MANIFEST, "version");
      manifest.set(path, { path, mediaType: nsAttr(fileEntry, MANIFEST, "media-type") ?? "", encrypted: nsFirst(fileEntry, MANIFEST, "encryption-data") !== undefined, ...(version !== undefined ? { version } : {}) });
    }
  } else {
    if (!reader.has("content.xml")) throw new PreviewError("Not an OpenDocument package.", "invalid-file");
    collector.report("odf.package.no-manifest", { part: MANIFEST_PATH });
  }

  // Rule 4: mimetype and the manifest root entry must agree (decision D4).
  const rootMediaType = manifest.get("/")?.mediaType || undefined;
  if (mimetype !== undefined && rootMediaType !== undefined && mimetype !== rootMediaType) throw new PreviewError("Package media type is inconsistent.", "invalid-file");
  const mediaType = mimetype ?? rootMediaType;
  if (!mediaType) throw new PreviewError("The package does not declare its media type.", "invalid-file");

  // Rule 5: refuse encrypted documents before parsing any document part (Part 2 §3.4).
  const encrypted = (path: string) => manifest.get(path)?.encrypted === true;
  if (encrypted("content.xml") || encrypted("styles.xml") || reader.has("encrypted-package") || manifest.has("encrypted-package")) {
    throw new PreviewError("Password-protected OpenDocument files are not supported.", "unsupported-feature");
  }

  // Rules 6 and 7: signatures and macros are detected by name only; their contents are never read.
  const entries = reader.entries;
  const signed = entries.some(entry => entry.name.startsWith("META-INF/") && entry.name.slice(9).includes("signatures"));
  if (signed) collector.report("odf.package.signed", { part: "META-INF" });
  const hasMacros = entries.some(entry => entry.name.startsWith("Basic/") || entry.name.startsWith("Scripts/"));
  if (hasMacros) collector.report("odf.package.macros");

  if (manifestVersion !== undefined && isNewerThanKnown(manifestVersion)) collector.report("odf.version.newer", { part: MANIFEST_PATH }, { version: manifestVersion });

  // Rule 8: optional metadata; never fails the preview.
  const generator = readGenerator(reader, collector, encrypted("meta.xml"));
  return {
    mediaType, manifest, signed, hasMacros,
    ...(manifestVersion !== undefined ? { manifestVersion } : {}),
    ...(generator !== undefined ? { generator } : {}),
    ...(reader.has(THUMBNAIL_PATH) ? { thumbnailPath: THUMBNAIL_PATH } : {}),
  };
}

/** Parses an XML part, prefixing parse errors with the part name. */
function parsePart(reader: PackageReader, path: string): NsXmlNode {
  try { return parseXmlNs(reader.readText(path)); }
  catch (error) {
    if (error instanceof PreviewError && error.code === "parse-error") throw new PreviewError(`${path}: ${error.message}`, "parse-error", { cause: error });
    throw error;
  }
}

function readGenerator(reader: PackageReader, collector: DiagnosticCollector, encrypted: boolean): string | undefined {
  const entry = reader.entry("meta.xml");
  if (!entry) return undefined;
  if (encrypted || entry.uncompressedSize > MAX_META_BYTES) { collector.report("odf.meta.ignored", { part: "meta.xml" }); return undefined; }
  try {
    const root = parsePart(reader, "meta.xml");
    if (root.uri !== OFFICE || root.local !== "document-meta") { collector.report("odf.meta.ignored", { part: "meta.xml" }); return undefined; }
    const generator = nsFirst(nsFirst(root, OFFICE, "meta"), META, "generator");
    const text = generator ? nsText(generator).trim() : "";
    return text || undefined;
  } catch (error) {
    // Budgets and cancellation still apply; anything else about meta.xml is only reported.
    if (error instanceof PreviewError && (error.code === "resource-limit" || error.code === "aborted")) throw error;
    collector.report("odf.meta.ignored", { part: "meta.xml" });
    return undefined;
  }
}

function isNewerThanKnown(version: string): boolean {
  const match = /^(\d+)\.(\d+)$/.exec(version.trim());
  if (!match) return false;
  const major = Number(match[1]), minor = Number(match[2]);
  return major > NEWEST_KNOWN[0] || (major === NEWEST_KNOWN[0] && minor > NEWEST_KNOWN[1]);
}
