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

/** A spreadsheet content.xml; `body` is the office:spreadsheet content (tables, settings), written with canonical prefixes. */
export function spreadsheetContent({ body, automatic = "", fontFaces = "", version = "1.3" }, prefixes = {}) {
  return odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document-content office:version="${version}"><office:font-face-decls>${fontFaces}</office:font-face-decls><office:automatic-styles>${automatic}</office:automatic-styles><office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document-content>`, prefixes);
}

/** A LibreOffice-like styles.xml with a default cell style (Liberation Sans 10 pt, en-US) plus `styles`. */
export function spreadsheetStyles({ styles = "", language = "en", country = "US" } = {}, prefixes = {}) {
  return odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document-styles office:version="1.3"><office:font-face-decls><style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'"/></office:font-face-decls><office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt" fo:language="${language}" fo:country="${country}"/></style:default-style><style:style style:name="Default" style:family="table-cell"/>${styles}</office:styles></office:document-styles>`, prefixes);
}

/** A complete .ods package from spreadsheetContent/spreadsheetStyles options. */
export function makeOds({ body, automatic, fontFaces, styles, generator = "LibreOffice/24.8.4.2$Windows_X86_64 LibreOffice_project/test", extra, mimetype, manifest } = {}, prefixes = {}) {
  const meta = odfXml(`<office:document-meta office:version="1.3"><office:meta><meta:generator>${generator}</meta:generator></office:meta></office:document-meta>`, prefixes);
  return makeOdfPackage({ content: spreadsheetContent({ body, automatic, fontFaces }, prefixes), styles: spreadsheetStyles({ styles }, prefixes), meta, ...(extra ? { extra } : {}), ...(mimetype ? { mimetype } : {}), ...(manifest !== undefined ? { manifest } : {}) });
}

/**
 * The flat (.fods) equivalent of makeOds: the same meta, font faces, common and automatic styles,
 * settings and body in one office:document (ODF 1.3 Part 3 §2.2.1), as a UTF-8 string.
 * `settings` is the inner content of office:settings.
 */
export function flatOdsXml({ body, automatic = "", fontFaces = "", styles = "", generator = "LibreOffice/24.8.4.2$Windows_X86_64 LibreOffice_project/test", settings, mimetype = ODS_MEDIA_TYPE, language = "en", country = "US" } = {}, prefixes = {}) {
  return odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document office:version="1.3" office:mimetype="${mimetype}"><office:meta><meta:generator>${generator}</meta:generator></office:meta>${settings !== undefined ? `<office:settings>${settings}</office:settings>` : ""}<office:font-face-decls><style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'"/>${fontFaces}</office:font-face-decls><office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt" fo:language="${language}" fo:country="${country}"/></style:default-style><style:style style:name="Default" style:family="table-cell"/>${styles}</office:styles><office:automatic-styles>${automatic}</office:automatic-styles><office:master-styles/><office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document>`, prefixes);
}
export const makeFods = (options, prefixes) => new TextEncoder().encode(flatOdsXml(options, prefixes));

// ---------------------------------------------------------------------------------------------
// OpenDocument text (milestones 3.1 to 3.5). Same shape as the spreadsheet helpers above, so a
// text fixture is written the same way and carries the same meta, styles and automatic styles.
// ---------------------------------------------------------------------------------------------

export const ODT_MEDIA_TYPE = "application/vnd.oasis.opendocument.text";
export const OTT_MEDIA_TYPE = "application/vnd.oasis.opendocument.text-template";

/** content.xml of a text document: `body` is the inner content of office:text. */
export function textContent({ body = "", automatic = "", fontFaces = "", version = "1.3" } = {}, prefixes = {}) {
  // `automatic: null` leaves out office:automatic-styles altogether, which the format allows: the
  // element is optional, and styles.xml must still be read without it.
  const automaticStyles = automatic === null ? "" : `<office:automatic-styles>${automatic}</office:automatic-styles>`;
  return odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document-content office:version="${version}"><office:font-face-decls>${fontFaces}</office:font-face-decls>${automaticStyles}<office:body><office:text>${body}</office:text></office:body></office:document-content>`, prefixes);
}

/** styles.xml of a text document: `styles` is the inner content of office:styles. */
export function textStyles({ styles = "", master = "" } = {}, prefixes = {}) {
  return odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document-styles office:version="1.3"><office:styles>${styles}</office:styles><office:automatic-styles/><office:master-styles>${master}</office:master-styles></office:document-styles>`, prefixes);
}

/** A complete .odt (or .ott) package from textContent/textStyles options. */
export function makeOdt({ body, automatic, fontFaces, styles, master, generator = "LibreOffice/24.8.4.2$Windows_X86_64 LibreOffice_project/test", extra, mimetype = ODT_MEDIA_TYPE, manifest } = {}, prefixes = {}) {
  const meta = odfXml(`<office:document-meta office:version="1.3"><office:meta><meta:generator>${generator}</meta:generator></office:meta></office:document-meta>`, prefixes);
  return makeOdfPackage({
    mimetype,
    content: textContent({ body, automatic, fontFaces }, prefixes),
    styles: textStyles({ styles, master }, prefixes),
    meta,
    ...(extra ? { extra } : {}),
    ...(manifest !== undefined ? { manifest } : {}),
  });
}

/** The flat (.fodt) equivalent of makeOdt, as a UTF-8 string. */
export function flatOdtXml({ body = "", automatic = "", fontFaces = "", styles = "", master = "", generator = "LibreOffice/24.8.4.2$Windows_X86_64 LibreOffice_project/test", mimetype = ODT_MEDIA_TYPE } = {}, prefixes = {}) {
  return odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document office:version="1.3" office:mimetype="${mimetype}"><office:meta><meta:generator>${generator}</meta:generator></office:meta><office:font-face-decls>${fontFaces}</office:font-face-decls><office:styles>${styles}</office:styles><office:automatic-styles>${automatic}</office:automatic-styles><office:master-styles>${master}</office:master-styles><office:body><office:text>${body}</office:text></office:body></office:document>`, prefixes);
}
export const makeFodt = (options, prefixes) => new TextEncoder().encode(flatOdtXml(options, prefixes));
