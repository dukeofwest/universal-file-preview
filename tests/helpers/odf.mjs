// Synthetic OpenDocument packages for tests (test-only; Node built-ins only).
import { makeZip } from "./zip.mjs";

export const ODF_NAMESPACES = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  draw: "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
  number: "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
  meta: "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
  config: "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
  chart: "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
  manifest: "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
  xlink: "http://www.w3.org/1999/xlink",
  dc: "http://purl.org/dc/elements/1.1/",
  of: "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
  calcext: "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
  loext: "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
};
export const ODS_MEDIA_TYPE = "application/vnd.oasis.opendocument.spreadsheet";

/**
 * Renders XML written with the canonical ODF prefixes (and no namespace declarations) using
 * the prefixes in `prefixes`, e.g. { office: "o", table: "t" }. Every namespace used is
 * declared on the root element. Formula values such as of:=... are renamed too.
 */
export function odfXml(xml, prefixes = {}) {
  const used = new Set(), rename = prefix => { used.add(prefix); return prefixes[prefix] ?? prefix; };
  let output = xml.replace(/(<\/?|\s)([a-z]+[0-9]*):(?=[A-Za-z_])/g, (match, lead, prefix) => Object.hasOwn(ODF_NAMESPACES, prefix) ? `${lead}${rename(prefix)}:` : match);
  output = output.replace(/="([a-z]+):=/g, (match, prefix) => Object.hasOwn(ODF_NAMESPACES, prefix) ? `="${rename(prefix)}:=` : match);
  const declarations = [...used].map(prefix => ` xmlns:${prefixes[prefix] ?? prefix}="${ODF_NAMESPACES[prefix]}"`).join("");
  return output.replace(/^(\s*(?:<\?xml[^>]*\?>\s*)?<[^\s/>]+)/, `$1${declarations}`);
}

/** A manifest listing `paths` (other than mimetype and the manifest itself) under a root entry. */
export function autoManifest(mediaType, paths, version = "1.3") {
  const typeOf = path => path.endsWith(".xml") ? "text/xml" : path.endsWith(".png") ? "image/png" : "";
  const entries = paths.filter(path => path !== "mimetype" && path !== "META-INF/manifest.xml").map(path => `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="${typeOf(path)}"/>`).join("");
  return odfXml(`<?xml version="1.0" encoding="UTF-8"?><manifest:manifest manifest:version="${version}"><manifest:file-entry manifest:full-path="/" manifest:version="${version}" manifest:media-type="${mediaType}"/>${entries}</manifest:manifest>`);
}

/**
 * Builds an ODF package. `mimetype: null` and `manifest: null` omit those entries;
 * `manifest: "auto"` (default) generates one. `order` lists entry names in the desired
 * order (default: mimetype first). The mimetype entry is STORED unless `storeMimetype` is false.
 */
export function makeOdfPackage({ mimetype = ODS_MEDIA_TYPE, manifest = "auto", content, styles, meta, settings, extra = {}, order, deflate = true, storeMimetype = true } = {}) {
  const parts = {};
  if (mimetype !== null) parts.mimetype = mimetype;
  if (content !== undefined) parts["content.xml"] = content;
  if (styles !== undefined) parts["styles.xml"] = styles;
  if (meta !== undefined) parts["meta.xml"] = meta;
  if (settings !== undefined) parts["settings.xml"] = settings;
  Object.assign(parts, extra);
  if (manifest === "auto") parts["META-INF/manifest.xml"] = autoManifest(mimetype ?? ODS_MEDIA_TYPE, Object.keys(parts));
  else if (manifest !== null) parts["META-INF/manifest.xml"] = manifest;
  const names = order ?? Object.keys(parts);
  const ordered = Object.fromEntries(names.map(name => [name, parts[name]]));
  return new Uint8Array(makeZip(ordered, { deflate, stored: storeMimetype ? ["mimetype"] : [] }));
}

/** Minimal spreadsheet content.xml. */
export const minimalContent = odfXml('<?xml version="1.0" encoding="UTF-8"?><office:document-content office:version="1.3"><office:body><office:spreadsheet/></office:body></office:document-content>');
