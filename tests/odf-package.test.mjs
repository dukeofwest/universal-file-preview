import test from "node:test";
import assert from "node:assert/strict";
import { PackageReader } from "../dist/package/package-reader.js";
import { inspectOdfPackage } from "../dist/package/odf-package.js";
import { DiagnosticCollector } from "../dist/diagnostics.js";
import { ZipArchive } from "../dist/archive/zip-reader.js";
import { PreviewError } from "../dist/errors.js";
import { makeZip } from "./helpers/zip.mjs";
import { makeOdfPackage, odfXml, autoManifest, minimalContent, ODS_MEDIA_TYPE, ODF_NAMESPACES } from "./helpers/odf.mjs";

const fails = (run, code, message) => assert.throws(run, error => error instanceof PreviewError && error.code === code && message.test(error.message), String(message));
const inspect = (bytes, limits) => {
  const collector = new DiagnosticCollector(), info = inspectOdfPackage(PackageReader.openZip(bytes, limits), collector);
  return { info, codes: collector.list().map(entry => entry.code), diagnostics: collector.list() };
};
const meta = generator => odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document-meta office:version="1.3"><office:meta><meta:generator>${generator}</meta:generator></office:meta></office:document-meta>`);
const centralOffset = archive => archive.readUInt32LE(archive.length - 6);
const encryptionData = '<manifest:encryption-data manifest:checksum-type="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k" manifest:checksum="AAAA"><manifest:algorithm manifest:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes256-cbc" manifest:initialisation-vector="AAAA"/><manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:key-size="32" manifest:iteration-count="100000" manifest:salt="AAAA"/></manifest:encryption-data>';
const manifestWith = entries => odfXml(`<manifest:manifest manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="${ODS_MEDIA_TYPE}"/>${entries}</manifest:manifest>`);

// -----------------------------------------------------------------------------------------------
// PackageReader
// -----------------------------------------------------------------------------------------------

test("PackageReader reads parts, normalises paths and tracks the total", () => {
  const reader = PackageReader.openZip(new Uint8Array(makeZip({ "a.xml": "<a/>", "dir/b.bin": "12345" }, { deflate: true })));
  assert.deepEqual(reader.entries.map(entry => entry.name), ["a.xml", "dir/b.bin"]);
  assert.equal(reader.has("/a.xml"), true);
  assert.equal(reader.has("dir//b.bin"), true);
  assert.equal(reader.entry("dir/b.bin").uncompressedSize, 5);
  assert.equal(reader.entry("missing"), undefined);
  assert.equal(new TextDecoder().decode(reader.readBytes("dir/b.bin")), "12345");
  assert.equal(reader.readText("/a.xml"), "<a/>");
  assert.equal(reader.totalBytes, 9);
  fails(() => reader.readBytes("missing.xml"), "invalid-file", /^Package part not found: missing\.xml$/);
});

test("PackageReader enforces part limits, per-path overrides and the total budget before decompressing", () => {
  const bytes = new Uint8Array(makeZip({ "content.xml": "x".repeat(20), "small.xml": "y".repeat(5) }, { deflate: true }));
  fails(() => PackageReader.openZip(bytes, { maxPartBytes: 10 }).readBytes("content.xml"), "resource-limit", /^Package part content\.xml exceeds the configured size limit\.$/);
  assert.equal(PackageReader.openZip(bytes, { maxPartBytes: 10, partBytes: { "/content.xml": 20 } }).readBytes("content.xml").length, 20);
  assert.equal(PackageReader.openZip(bytes, { maxTotalBytes: 25 }).readBytes("content.xml").length, 20, "reaching the budget exactly is allowed");
  const budget = PackageReader.openZip(bytes, { maxTotalBytes: 29 });
  budget.readBytes("content.xml"); budget.readBytes("small.xml");
  fails(() => budget.readBytes("small.xml"), "resource-limit", /^Package exceeds the configured total size limit\.$/);
  assert.equal(budget.totalBytes, 25, "a rejected read does not count");
  // Part names are untrusted: inherited object properties must not act as limits.
  const tricky = new Uint8Array(makeZip({ constructor: "z".repeat(20), toString: "z".repeat(20) }));
  for (const name of ["constructor", "toString"]) fails(() => PackageReader.openZip(tricky, { maxPartBytes: 10 }).readBytes(name), "resource-limit", /exceeds the configured size limit/);
});

test("PackageReader opens packages with large unused parts that ZipArchive's default limit rejects", () => {
  const archive = makeZip({ "content.xml": "<a/>", "Media/video.bin": "x" });
  const second = centralOffset(archive) + 46 + "content.xml".length;
  archive.writeUInt32LE(40 * 1024 * 1024, second + 24);
  fails(() => ZipArchive.open(archive), "resource-limit", /^ZIP entry exceeds the configured size limit\.$/);
  const reader = PackageReader.openZip(new Uint8Array(archive));
  assert.equal(reader.readText("content.xml"), "<a/>");
  fails(() => reader.readBytes("Media/video.bin"), "resource-limit", /^Package part Media\/video\.bin exceeds/);
  assert.equal(reader.totalBytes, 4);
});

test("PackageReader decodes UTF-8 and UTF-16 by byte-order mark and rejects other encodings", () => {
  const bom = String.fromCharCode(0xfeff), utf16le = text => Buffer.from(bom + text, "utf16le"), utf16be = text => { const bytes = utf16le(text); for (let i = 0; i < bytes.length; i += 2) [bytes[i], bytes[i + 1]] = [bytes[i + 1], bytes[i]]; return bytes; };
  const reader = PackageReader.openZip(new Uint8Array(makeZip({
    "utf8.xml": '<?xml version="1.0" encoding="UTF-8"?><a/>',
    "bom.xml": Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from("<a/>")]),
    "le.xml": utf16le('<?xml version="1.0" encoding="UTF-16"?><a>x</a>'),
    "be.xml": utf16be("<a>y</a>"),
    "latin1.xml": "<?xml version='1.0' encoding='ISO-8859-1'?><a/>",
    "bad.xml": Buffer.from([0x3c, 0x61, 0xff, 0x2f, 0x3e]),
    "bad16.xml": Buffer.from([0xff, 0xfe, 0x00, 0xd8]),
  })));
  assert.equal(reader.readText("utf8.xml"), '<?xml version="1.0" encoding="UTF-8"?><a/>');
  assert.equal(reader.readText("bom.xml"), "<a/>");
  assert.equal(reader.readText("le.xml"), '<?xml version="1.0" encoding="UTF-16"?><a>x</a>');
  assert.equal(reader.readText("be.xml"), "<a>y</a>");
  fails(() => reader.readText("latin1.xml"), "unsupported-feature", /^Package part latin1\.xml uses the unsupported encoding ISO-8859-1\.$/);
  fails(() => reader.readText("bad.xml"), "invalid-file", /^Package part bad\.xml is not valid UTF-8 text\.$/);
  fails(() => reader.readText("bad16.xml"), "invalid-file", /is not valid UTF-16LE text/);
});

test("PackageReader validates its limits", () => {
  const bytes = new Uint8Array(makeZip({ "a.xml": "<a/>" }));
  for (const key of ["maxEntries", "maxPartBytes", "maxTotalBytes"]) for (const value of [0, -1, 1.5, Number.NaN]) assert.throws(() => PackageReader.openZip(bytes, { [key]: value }), RangeError, `${key}=${value}`);
  assert.throws(() => PackageReader.openZip(bytes, { partBytes: { "a.xml": 0 } }), RangeError);
  fails(() => PackageReader.openZip(new Uint8Array(makeZip({ a: "1", b: "2", c: "3" })), { maxEntries: 2 }), "resource-limit", /central directory/);
});

// -----------------------------------------------------------------------------------------------
// inspectOdfPackage
// -----------------------------------------------------------------------------------------------

test("a well-formed spreadsheet package is described without diagnostics", () => {
  const { info, codes } = inspect(makeOdfPackage({ content: minimalContent, styles: odfXml("<office:document-styles/>"), meta: meta("LibreOffice/24.8.4.2$Windows_X86_64 LibreOffice_project/abc"), extra: { "Thumbnails/thumbnail.png": "png" } }));
  assert.equal(info.mediaType, ODS_MEDIA_TYPE);
  assert.equal(info.manifestVersion, "1.3");
  assert.equal(info.generator, "LibreOffice/24.8.4.2$Windows_X86_64 LibreOffice_project/abc");
  assert.equal(info.thumbnailPath, "Thumbnails/thumbnail.png");
  assert.deepEqual([info.signed, info.hasMacros], [false, false]);
  assert.deepEqual([...info.manifest.keys()], ["/", "content.xml", "styles.xml", "meta.xml", "Thumbnails/thumbnail.png"]);
  assert.deepEqual(info.manifest.get("/"), { path: "/", mediaType: ODS_MEDIA_TYPE, encrypted: false, version: "1.3" });
  assert.deepEqual(info.manifest.get("content.xml"), { path: "content.xml", mediaType: "text/xml", encrypted: false });
  assert.deepEqual(codes, []);
});

test("rule 1: the mimetype entry is trimmed, and position or compression problems are reported", () => {
  assert.equal(inspect(makeOdfPackage({ mimetype: `${ODS_MEDIA_TYPE}\r\n`, content: minimalContent, manifest: autoManifest(ODS_MEDIA_TYPE, ["content.xml"]) })).info.mediaType, ODS_MEDIA_TYPE);
  assert.deepEqual(inspect(makeOdfPackage({ content: minimalContent, order: ["content.xml", "mimetype", "META-INF/manifest.xml"] })).codes, ["odf.package.mimetype-position"]);
  assert.deepEqual(inspect(makeOdfPackage({ content: minimalContent, storeMimetype: false })).codes, ["odf.package.mimetype-position"]);
  fails(() => inspect(makeOdfPackage({ mimetype: "application/\u00ffods", manifest: null, content: minimalContent })), "invalid-file", /^The package mimetype entry is not ASCII text\.$/);
  fails(() => inspect(makeOdfPackage({ mimetype: "a".repeat(300), manifest: null, content: minimalContent })), "invalid-file", /^The package mimetype entry is too large\.$/);
  fails(() => inspect(makeOdfPackage({ mimetype: "  ", manifest: null, content: minimalContent })), "invalid-file", /^The package mimetype entry is empty\.$/);
});

test("rule 2: a missing manifest is tolerated only with content.xml and a mimetype", () => {
  const { info, codes } = inspect(makeOdfPackage({ manifest: null, content: minimalContent }));
  assert.deepEqual([info.mediaType, info.manifest.size, codes], [ODS_MEDIA_TYPE, 0, ["odf.package.no-manifest"]]);
  fails(() => inspect(makeOdfPackage({ manifest: null, styles: "<x/>" })), "invalid-file", /^Not an OpenDocument package\.$/);
  fails(() => inspect(makeOdfPackage({ mimetype: null, manifest: null, content: minimalContent })), "invalid-file", /^The package does not declare its media type\.$/);
  // An OOXML package is not mistaken for ODF.
  fails(() => inspect(new Uint8Array(makeZip({ "[Content_Types].xml": "<Types/>", "xl/workbook.xml": "<workbook/>" }))), "invalid-file", /^Not an OpenDocument package\.$/);
});

test("rule 3: the manifest is resolved by namespace, and malformed manifests are rejected", () => {
  const renamed = `<manifest xmlns="${ODF_NAMESPACES.manifest}" xmlns:m="${ODF_NAMESPACES.manifest}" m:version="1.3"><file-entry m:full-path="/" m:media-type="${ODS_MEDIA_TYPE}"/><file-entry m:full-path="content.xml" m:media-type="text/xml"/></manifest>`;
  const prefixed = odfXml(`<manifest:manifest manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="${ODS_MEDIA_TYPE}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`, { manifest: "zz" });
  for (const manifest of [renamed, prefixed]) {
    const { info } = inspect(makeOdfPackage({ manifest, content: minimalContent }));
    assert.deepEqual([info.manifestVersion, [...info.manifest.keys()]], ["1.3", ["/", "content.xml"]]);
  }
  fails(() => inspect(makeOdfPackage({ manifest: '<x:manifest xmlns:x="urn:other"/>', content: minimalContent })), "invalid-file", /^META-INF\/manifest\.xml is not an OpenDocument manifest\.$/);
  fails(() => inspect(makeOdfPackage({ manifest: "<manifest:manifest>", content: minimalContent })), "parse-error", /^META-INF\/manifest\.xml: Unbound namespace prefix manifest/);
  fails(() => inspect(makeOdfPackage({ manifest: '<!DOCTYPE manifest:manifest PUBLIC "-//OpenOffice.org//DTD Manifest 1.0//EN" "Manifest.dtd"><manifest:manifest/>', content: minimalContent })), "parse-error", /^META-INF\/manifest\.xml: DOCTYPE/);
  const entries = manifestWith('<manifest:file-entry manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="ignored"/>');
  assert.equal(inspect(makeOdfPackage({ manifest: entries, content: minimalContent })).info.manifest.get("content.xml").mediaType, "text/xml", "entries without full-path are skipped; the first duplicate wins");
});

test("rule 4: mimetype and manifest root must agree; the manifest root alone also works", () => {
  fails(() => inspect(makeOdfPackage({ mimetype: "application/vnd.oasis.opendocument.text", manifest: manifestWith(""), content: minimalContent })), "invalid-file", /^Package media type is inconsistent\.$/);
  assert.equal(inspect(makeOdfPackage({ mimetype: null, manifest: manifestWith(""), content: minimalContent })).info.mediaType, ODS_MEDIA_TYPE);
  fails(() => inspect(makeOdfPackage({ mimetype: null, manifest: odfXml('<manifest:manifest manifest:version="1.3"/>'), content: minimalContent })), "invalid-file", /^The package does not declare its media type\.$/);
});

test("rule 5: encrypted documents are refused before any document part is parsed", () => {
  const encryptedEntry = path => `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="text/xml">${encryptionData}</manifest:file-entry>`;
  // The encrypted content is not XML; reaching a parser would fail differently.
  for (const path of ["content.xml", "styles.xml"]) fails(() => inspect(makeOdfPackage({ manifest: manifestWith(encryptedEntry(path)), content: "\u0001\u0002 not xml", styles: "\u0003 not xml", meta: "not xml" })), "unsupported-feature", /^Password-protected OpenDocument files are not supported\.$/);
  fails(() => inspect(makeOdfPackage({ content: minimalContent, extra: { "encrypted-package": "x" } })), "unsupported-feature", /^Password-protected/);
  fails(() => inspect(makeOdfPackage({ manifest: manifestWith('<manifest:file-entry manifest:full-path="encrypted-package" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'), content: minimalContent })), "unsupported-feature", /^Password-protected/);
  // Encrypted images alone do not block the preview; an encrypted meta.xml is only reported.
  const partial = inspect(makeOdfPackage({ manifest: manifestWith(encryptedEntry("Pictures/a.png") + encryptedEntry("meta.xml")), content: minimalContent, meta: "\u0004 not xml", extra: { "Pictures/a.png": "x" } }));
  assert.equal(partial.info.manifest.get("Pictures/a.png").encrypted, true);
  assert.equal(partial.info.generator, undefined);
  assert.deepEqual(partial.codes, ["odf.meta.ignored"]);
});

test("rules 6 and 7: signatures and macros are detected by name and never read", () => {
  const bytes = makeOdfPackage({ content: minimalContent, extra: { "META-INF/documentsignatures.xml": "<sig/>", "META-INF/macrosignatures.xml": "<sig/>", "Basic/Standard/Module1.xml": "<script/>", "Scripts/python/a.py": "print(1)" } });
  const reader = PackageReader.openZip(bytes), reads = [], read = reader.readBytes.bind(reader);
  reader.readBytes = path => { reads.push(path); return read(path); };
  const collector = new DiagnosticCollector(), info = inspectOdfPackage(reader, collector);
  assert.deepEqual([info.signed, info.hasMacros], [true, true]);
  assert.deepEqual(collector.list().map(entry => [entry.code, entry.count]), [["odf.package.signed", 1], ["odf.package.macros", 1]]);
  assert.deepEqual(reads, ["mimetype", "META-INF/manifest.xml"]);
  assert.equal(inspect(makeOdfPackage({ content: minimalContent, extra: { "META-INF/other.xml": "<x/>", "NotBasic/x.xml": "<x/>" } })).codes.length, 0);
});

test("rule 8: meta.xml is optional, resolved by namespace, and never fails the preview", () => {
  assert.equal(inspect(makeOdfPackage({ content: minimalContent, meta: odfXml(`<office:document-meta><office:meta><meta:generator>  Gen/1  </meta:generator></office:meta></office:document-meta>`, { office: "o", meta: "m" }) })).info.generator, "Gen/1");
  assert.equal(inspect(makeOdfPackage({ content: minimalContent, meta: odfXml("<office:document-meta><office:meta/></office:document-meta>") })).info.generator, undefined);
  for (const broken of ["<office:document-meta>", '<x:other xmlns:x="urn:x"/>', "x".repeat(1024 * 1024 + 1)]) {
    const result = inspect(makeOdfPackage({ content: minimalContent, meta: broken }));
    assert.deepEqual([result.info.generator, result.codes], [undefined, ["odf.meta.ignored"]], broken.slice(0, 30));
  }
  // Budgets still apply to meta.xml.
  const big = makeOdfPackage({ content: minimalContent, meta: meta("G".repeat(5000)) });
  fails(() => inspect(big, { maxTotalBytes: 3000 }), "resource-limit", /total size limit/);
});

test("manifest versions newer than ODF 1.4 are reported", () => {
  const withVersion = version => inspect(makeOdfPackage({ content: minimalContent, manifest: autoManifest(ODS_MEDIA_TYPE, ["content.xml"], version) }));
  const newer = withVersion("1.5");
  assert.deepEqual(newer.codes, ["odf.version.newer"]);
  assert.equal(newer.diagnostics[0].message, "The document uses OpenDocument version 1.5; newer features may be missing.");
  for (const version of ["1.2", "1.3", "1.4", "x"]) assert.deepEqual(withVersion(version).codes, [], version);
  assert.deepEqual(withVersion("2.0").codes, ["odf.version.newer"]);
});
