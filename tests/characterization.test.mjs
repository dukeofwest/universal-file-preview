// M0 characterization baseline (docs/research/odf-implementation-specification.md, §7 M0 and §8).
//
// These tests record what the library does TODAY, including surprising behaviour, so later
// milestones can prove they changed nothing unintentionally. They are not a statement of what
// the behaviour should be. Cases that changed on purpose later are tagged "changed in M7".
import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { deflateRawSync } from "node:zlib";
import { makeZip } from "./helpers/zip.mjs";
import { normalize, matchSnapshot } from "./helpers/snapshot.mjs";
import { ooxmlFixtures, buildFixture } from "./helpers/ooxml-fixtures.mjs";
import { writeDemoSamples } from "../scripts/demo-samples.mjs";
import { detectFormat } from "../dist/detection/detect-format.js";
import { parseXml, elements, firstElement, textContent } from "../dist/xml/xml-parser.js";
import { ZipArchive } from "../dist/archive/zip-reader.js";
import { inflateRaw } from "../dist/archive/inflate.js";
import { PreviewError } from "../dist/errors.js";
import { parseDocx, parseXlsx, parsePptx } from "../dist/parsers/ooxml.js";

// -----------------------------------------------------------------------------------------------
// Byte builders
// -----------------------------------------------------------------------------------------------

const bytes = (...values) => Uint8Array.from(values);
const ascii = text => Uint8Array.from(text, character => character.charCodeAt(0));
const utf8 = text => new Uint8Array(Buffer.from(text, "utf8"));
const concat = (...parts) => new Uint8Array(Buffer.concat(parts.map(part => Buffer.from(part))));
const PNG = bytes(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a);
const JPEG = bytes(0xff, 0xd8, 0xff, 0xe0);
const CFB = bytes(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);
const PK = bytes(0x50, 0x4b, 0x03, 0x04);
/** "%PDF-1.7" preceded by `offset` spaces. */
const pdfAt = offset => concat(new Uint8Array(offset).fill(0x20), ascii("%PDF-1.7"));
const ooxmlZip = new Uint8Array(makeZip({ "[Content_Types].xml": "<Types/>", "word/document.xml": "<w:document/>" }));
/** A minimal ODF-shaped package: STORED `mimetype` as the first entry (makeZip stores by default). */
const odsZip = new Uint8Array(makeZip({
  mimetype: "application/vnd.oasis.opendocument.spreadsheet",
  "META-INF/manifest.xml": "<manifest:manifest/>",
  "content.xml": "<office:document-content/>",
}));

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const ODS = "application/vnd.oasis.opendocument.spreadsheet";

// -----------------------------------------------------------------------------------------------
// 1. Detection matrix (spec §8.1)
// -----------------------------------------------------------------------------------------------

/** [label, name, mimeType, bytes, expected] — `bytes` undefined means detectFormat(name, mime). */
const detectionCases = [
  // Signatures take precedence over MIME type and extension.
  ["PNG signature beats PDF name and PNG MIME", "misleading.pdf", "image/png", PNG, "png"],
  ["PDF marker at offset 0", "x.bin", "", ascii("%PDF-1.7"), "pdf"],
  ["PDF marker inside the first 1 KiB", "x.bin", "", pdfAt(500), "pdf"],
  ["PDF marker ending exactly at byte 1024", "x.bin", "", pdfAt(1019), "pdf"],
  ["PDF marker starting at byte 1020 is outside the window", "x.bin", "", pdfAt(1020), "unknown"],
  ["GIF89a signature", "x.bin", "", ascii("GIF89a"), "gif"],
  ["GIF87a signature", "x.bin", "", ascii("GIF87a"), "gif"],
  ["JPEG signature", "x.bin", "", JPEG, "jpeg"],
  ["WebP RIFF container", "x.bin", "", concat(ascii("RIFF"), bytes(0, 0, 0, 0), ascii("WEBPVP8 ")), "webp"],
  ["SVG root element", "x.bin", "", utf8('<svg xmlns="http://www.w3.org/2000/svg"></svg>'), "svg"],
  ["SVG after an XML declaration", "x.bin", "", utf8('<?xml version="1.0"?>\n<svg xmlns="http://www.w3.org/2000/svg"></svg>'), "svg"],
  ["SVG after a UTF-8 BOM and whitespace", "x.bin", "", utf8("\uFEFF  <svg></svg>"), "svg"],
  ["PNG signature beats .svg extension", "a.svg", "", PNG, "png"],
  ["JPEG signature beats .png extension", "image.png", "", JPEG, "jpeg"],

  // Compound File Binary: resolved by extension, then MIME type.
  ["CFB with .xls extension", "legacy.xls", "application/octet-stream", CFB, "xls"],
  ["CFB with .doc extension", "x.doc", "", CFB, "doc"],
  ["CFB with .ppt extension", "x.ppt", "", CFB, "ppt"],
  ["CFB with PowerPoint MIME", "x.bin", "application/vnd.ms-powerpoint", CFB, "ppt"],
  ["CFB with Word MIME", "x.bin", "application/msword", CFB, "doc"],
  ["CFB without metadata", "x.bin", "", CFB, "unknown"],

  // ZIP: resolved by extension, then MIME type.
  ["ZIP header with DOCX MIME", "renamed.bin", DOCX, PK, "docx"],
  ["OOXML ZIP with .xlsx", "book.xlsx", "", ooxmlZip, "xlsx"],
  ["OOXML ZIP with .pptx", "deck.pptx", "", ooxmlZip, "pptx"],
  ["OOXML ZIP with .docx", "doc.docx", "", ooxmlZip, "docx"],
  ["OOXML ZIP with upper-case extension", "Report.DOCX", "", ooxmlZip, "docx"],
  ["OOXML ZIP with .zip extension", "unknown.zip", "", ooxmlZip, "unknown"],
  ["OOXML ZIP with .zip extension and XLSX MIME", "unknown.zip", XLSX, ooxmlZip, "xlsx"],
  ["OOXML ZIP without extension", "x.bin", "", ooxmlZip, "unknown"],

  // Metadata only (no bytes): MIME type first, then extension.
  ["extension only", "report.docx", "", undefined, "docx"],
  ["MIME only", "data.bin", "application/pdf", undefined, "pdf"],
  ["MIME beats extension without bytes", "a.png", "application/pdf", undefined, "pdf"],
  ["MIME parameters are ignored", "x.webp", "image/webp; charset=binary", undefined, "webp"],
  ["MIME type is case-insensitive", "x.bin", "IMAGE/PNG", undefined, "png"],
  ["unrecognised MIME falls back to extension", "a.docx", "text/plain", undefined, "docx"],
  ["Excel MIME", "x.bin", "application/vnd.ms-excel", undefined, "xls"],
  ["extension is case-insensitive", "photo.JPG", "", undefined, "jpeg"],
  [".jpeg extension", "x.jpeg", "", undefined, "jpeg"],
  [".gif extension", "x.gif", "", undefined, "gif"],
  [".png extension", "x.png", "", undefined, "png"],
  [".svg extension", "x.svg", "", undefined, "svg"],
  [".pdf extension", "x.pdf", "", undefined, "pdf"],
  [".xlsx extension", "x.xlsx", "", undefined, "xlsx"],
  [".pptx extension", "x.pptx", "", undefined, "pptx"],
  [".doc extension", "x.doc", "", undefined, "doc"],
  [".xls extension", "x.xls", "", undefined, "xls"],
  [".ppt extension", "x.ppt", "", undefined, "ppt"],
  ["no extension", "noextension", "", undefined, "unknown"],
  ["empty name", "", "", undefined, "unknown"],
  ["only the last extension counts", "archive.tar.gz", "", undefined, "unknown"],
  ["generic binary MIME", "x.bin", "application/octet-stream", undefined, "unknown"],
  ["empty byte array falls back to extension", "a.png", "", new Uint8Array(0), "png"],
  ["text bytes with unmapped extension", "notes.txt", "", ascii("hello"), "unknown"],

  // OpenDocument: these rows were "unknown" at the M0 baseline and changed intentionally in M7
  // (spec §8.3). They are the only baseline rows allowed to change.
  ["ODS extension only (changed in M7)", "sheet.ods", "", undefined, "ods"],
  ["OTS extension only (changed in M7)", "sheet.ots", "", undefined, "ots"],
  ["FODS extension only (changed in M7)", "sheet.fods", "", undefined, "fods"],
  ["ODS MIME without bytes (changed in M7)", "sheet.ods", ODS, undefined, "ods"],
  ["ODS package with .ods (changed in M7)", "sheet.ods", "", odsZip, "ods"],
  ["ODS package with ODS MIME (changed in M7)", "sheet.ods", ODS, odsZip, "ods"],
  ["ODS package with .zip (changed in M7)", "data.zip", "", odsZip, "ods"],
];

/**
 * Current behaviour that is surprising but real. Recorded so that any change to it is a
 * deliberate, reviewed decision rather than a side effect.
 */
const detectionQuirks = [
  ["a %PDF- marker anywhere in the first 1 KiB beats every other signature", "x.bin", "", concat(PNG, ascii("%PDF-")), "pdf"],
  ["a RIFF file that is not WebP falls back to the extension", "sound.webp", "", concat(ascii("RIFF"), bytes(0, 0, 0, 0), ascii("WAVE")), "webp"],
  ["a self-closing <svg/> root is not sniffed", "x.bin", "", utf8("<svg/>"), "unknown"],
  ["a comment before <svg> is not sniffed", "x.bin", "", utf8("<!-- c --><svg></svg>"), "unknown"],
  ["a CFB container named .docx (e.g. a password-protected OOXML file) detects as docx", "secret.docx", "", CFB, "docx"],
  ["a ZIP named with an image extension detects as that image format", "photo.png", "", ooxmlZip, "png"],
  ["a ZIP named .doc detects as doc", "x.doc", "", ooxmlZip, "doc"],
  ["bytes without a known signature fall back to the extension", "report.docx", "", ascii("hello"), "docx"],
];

// "ods", "ots" and "fods" were added to FileFormat in M7 (decision D1).
const knownFormats = new Set(["pdf", "docx", "xlsx", "pptx", "doc", "xls", "ppt", "gif", "png", "jpeg", "webp", "svg", "ods", "ots", "fods", "unknown"]);
const detect = ([, name, mime, input]) => input === undefined ? detectFormat(name, mime) : detectFormat(name, mime, input);

test("M0 detection matrix records current detectFormat behaviour", () => {
  for (const testCase of detectionCases) assert.equal(detect(testCase), testCase[4], testCase[0]);
});

test("M0 detection quirks record current detectFormat behaviour", () => {
  for (const testCase of detectionQuirks) assert.equal(detect(testCase), testCase[4], testCase[0]);
});

test("M0 detection only returns members of the current FileFormat union", () => {
  for (const testCase of [...detectionCases, ...detectionQuirks]) assert.ok(knownFormats.has(detect(testCase)), testCase[0]);
  for (let length = 0; length <= 16; length++) assert.ok(knownFormats.has(detectFormat("x.bin", "", new Uint8Array(length).fill(0x50))));
});

// -----------------------------------------------------------------------------------------------
// 2. parseXml behaviour (protects the OOXML XML layer; spec AD-1 keeps it unchanged)
// -----------------------------------------------------------------------------------------------

const previewError = (code, message) => error => error instanceof PreviewError && error.name === "PreviewError" && error.code === code && message.test(error.message);

test("M0 parseXml limits: elements, depth, and source length", () => {
  assert.equal(parseXml(`<r>${"<a/>".repeat(249_999)}</r>`).children.length, 249_999);
  assert.throws(() => parseXml(`<r>${"<a/>".repeat(250_000)}</r>`), previewError("resource-limit", /^XML element count exceeds the configured limit\.$/));
  assert.throws(() => parseXml("<r><a/><a/></r>", { maxNodes: 2 }), previewError("resource-limit", /element count/));
  assert.equal(parseXml(`${"<a>".repeat(256)}${"</a>".repeat(256)}`).localName, "a");
  assert.throws(() => parseXml(`${"<a>".repeat(257)}${"</a>".repeat(257)}`), previewError("resource-limit", /^XML nesting exceeds the configured limit\.$/));
  assert.throws(() => parseXml("<a><b/></a>", { maxDepth: 1 }), previewError("resource-limit", /nesting/));
  assert.throws(() => parseXml("x".repeat(16 * 1024 * 1024 + 1)), previewError("resource-limit", /^XML document exceeds the configured size limit\.$/));
  assert.throws(() => parseXml("<a>12345</a>", { maxTextLength: 10 }), previewError("resource-limit", /size limit/));
});

test("M0 parseXml rejects DTDs and entity declarations with a fixed message", () => {
  assert.throws(() => parseXml("<!DOCTYPE a><a/>"), previewError("parse-error", /^DOCTYPE and entity declarations are not supported \(at character 0\)\.$/));
  assert.throws(() => parseXml("<!doctype a><a/>"), previewError("parse-error", /^DOCTYPE and entity declarations/));
  assert.throws(() => parseXml('<!ENTITY x "y"><a/>'), previewError("parse-error", /^DOCTYPE and entity declarations/));
  assert.throws(() => parseXml('<?xml version="1.0"?><!DOCTYPE office:document-content PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "office.dtd"><a/>'), previewError("parse-error", /DOCTYPE/));
});

test("M0 parseXml entities and character references", () => {
  const root = parseXml('<a b="&lt;&#x41;&#X42;">&amp;&#67;&quot;&apos;&gt;</a>');
  assert.equal(root.attributes.b, "<AB");
  assert.equal(textContent(root), "&C\"'>");
  assert.throws(() => parseXml("<a>&#0;</a>"), previewError("parse-error", /Invalid XML character reference/));
  assert.throws(() => parseXml("<a>&#xD800;</a>"), previewError("parse-error", /Invalid XML character reference/));
  assert.throws(() => parseXml("<a>&#x110000;</a>"), previewError("parse-error", /Invalid XML character reference/));
  assert.throws(() => parseXml("<a>&nbsp;</a>"), previewError("parse-error", /Unescaped or unsupported ampersand/));
  assert.throws(() => parseXml('<a b="x & y"/>'), previewError("parse-error", /Unescaped or unsupported ampersand/));
});

test("M0 parseXml names: prefixed attribute keys, prefix-blind localName and element helpers", () => {
  const root = parseXml('<p:a xmlns:p="urn:p" xmlns:q="urn:q" p:x="1" y="2"><p:b/><q:b>t</q:b><c/></p:a>');
  assert.equal(root.name, "p:a");
  assert.equal(root.localName, "a");
  assert.deepEqual(root.attributes, { "xmlns:p": "urn:p", "xmlns:q": "urn:q", "p:x": "1", y: "2" });
  // Namespace URIs are not tracked: helpers match by local name regardless of prefix or namespace.
  assert.deepEqual(elements(root, "b").map(node => node.name), ["p:b", "q:b"]);
  assert.equal(firstElement(root, "b").name, "p:b");
  assert.equal(firstElement(root, "missing"), undefined);
  assert.equal(textContent(root), "t");
});

test("M0 parseXml whitespace, comments, processing instructions, BOM and CDATA", () => {
  assert.deepEqual(parseXml("<a> <b/> </a>").children.map(child => typeof child === "string" ? child : child.name), [" ", "b", " "]);
  assert.deepEqual(parseXml("<?xml version='1.0'?><!-- c --><a><?pi x?><!--d--></a>").children, []);
  assert.equal(parseXml("\uFEFF<a/>").name, "a");
  assert.equal(parseXml("  <a/>  \n").name, "a");
  assert.equal(textContent(parseXml("<a><![CDATA[<raw>]]></a>")), "<raw>");
  // Attribute values are not whitespace-normalised (XML 1.0 §3.3.3 is not applied).
  assert.equal(parseXml('<a x="1\n2\t3"/>').attributes.x, "1\n2\t3");
});

test("M0 parseXml quirk: CDATA content is entity-decoded (XML 1.0 says it is literal)", () => {
  // Recorded as current behaviour, not endorsed: a bare "&" inside CDATA is rejected and
  // "&amp;" inside CDATA becomes "&". Changing this is a deliberate decision (spec D11).
  assert.equal(textContent(parseXml("<a><![CDATA[a &amp; b]]></a>")), "a & b");
  assert.throws(() => parseXml("<a><![CDATA[a & b]]></a>"), previewError("parse-error", /Unescaped or unsupported ampersand/));
});

test("M0 parseXml well-formedness errors", () => {
  const cases = [
    ["<a/><b/>", /^Multiple root elements \(at character \d+\)\.$/],
    ["x<a/>", /^Text outside the root element/],
    ["<a>", /^Unclosed element a/],
    ["<a></b>", /^Mismatched closing tag b/],
    ["<a b='1' b='2'/>", /^Duplicate attribute b/],
    ["<a b=1/>", /^Attribute value must be quoted/],
    ["<a b/>", /^Expected '=' after attribute name/],
    ["<a", /^Unclosed start tag/],
    ["<a><!-- x</a>", /^Unclosed XML comment/],
    ["<a><![CDATA[x</a>", /^Unclosed CDATA section/],
    ["<?xml version='1.0'", /^Unclosed processing instruction/],
    ["", /^XML document has no root element/],
  ];
  for (const [source, message] of cases) assert.throws(() => parseXml(source), previewError("parse-error", message), JSON.stringify(source));
});

// -----------------------------------------------------------------------------------------------
// 3. ZIP and DEFLATE defaults (protects spec M5)
// -----------------------------------------------------------------------------------------------

/** Offset of the first central-directory record in a makeZip archive. */
const centralOffset = archive => archive.readUInt32LE(archive.length - 6);

test("M0 ZipArchive default limits and rejections", () => {
  assert.equal(ZipArchive.open(makeZip(Object.fromEntries(Array.from({ length: 4096 }, (_, index) => [`f${index}`, "x"])))).entries.length, 4096);
  assert.throws(() => ZipArchive.open(makeZip(Object.fromEntries(Array.from({ length: 4097 }, (_, index) => [`f${index}`, "x"])))), previewError("resource-limit", /^Invalid or oversized ZIP central directory\.$/));

  // The local header (at offset 0) declares the same size, so the archive is consistent apart from its contents.
  const declaredSize = size => { const archive = makeZip({ "big.bin": "x" }); archive.writeUInt32LE(size, centralOffset(archive) + 24); archive.writeUInt32LE(size, 22); return archive; };
  assert.equal(ZipArchive.open(declaredSize(32 * 1024 * 1024)).entries[0].uncompressedSize, 32 * 1024 * 1024);
  assert.throws(() => ZipArchive.open(declaredSize(32 * 1024 * 1024 + 1)), previewError("resource-limit", /^ZIP entry exceeds the configured size limit\.$/));

  const withMethod = method => { const archive = makeZip({ "a.txt": "x" }); archive.writeUInt16LE(method, centralOffset(archive) + 10); return archive; };
  assert.throws(() => ZipArchive.open(withMethod(12)), previewError("unsupported-feature", /^ZIP compression method 12 is unsupported\.$/));
  const encrypted = makeZip({ "a.txt": "x" }); encrypted.writeUInt16LE(0x801, centralOffset(encrypted) + 8);
  assert.throws(() => ZipArchive.open(encrypted), previewError("unsupported-feature", /^Encrypted ZIP entries are unsupported\.$/));
  const zip64 = makeZip({ "a.txt": "x" }); zip64.writeUInt32LE(0xffffffff, centralOffset(zip64) + 20);
  assert.throws(() => ZipArchive.open(zip64), previewError("unsupported-feature", /^ZIP64 or multi-disk entry is unsupported\.$/));
  assert.throws(() => ZipArchive.open(makeZip({ "dir\\file": "x" })), previewError("invalid-file", /^Unsafe ZIP package path\.$/));
  assert.throws(() => ZipArchive.open(makeZip({ "a.txt": "x" })).read("missing.txt"), previewError("invalid-file", /^ZIP entry not found: missing\.txt$/));
});

test("M0 inflateRaw output matches zlib for varied payloads and levels", () => {
  let seed = 0x2545f491;
  const random = () => (seed = (seed * 1103515245 + 12345) >>> 0) >>> 24;
  const payloads = [
    new Uint8Array(0),
    utf8("a"),
    utf8("repeated text ".repeat(5000)),
    Uint8Array.from({ length: 70_000 }, random),
    Uint8Array.from({ length: 200_000 }, (_, index) => (index % 251) ^ (index >> 9)),
    utf8(`<row>${'<c r="A1" s="1"><v>1234.5</v></c>'.repeat(3000)}</row>`),
  ];
  for (const [index, payload] of payloads.entries()) {
    for (const level of [0, 1, 6, 9]) {
      const compressed = new Uint8Array(deflateRawSync(payload, { level }));
      assert.deepEqual(inflateRaw(compressed, payload.length), payload, `payload ${index}, level ${level}`);
    }
  }
  assert.throws(() => inflateRaw(new Uint8Array(deflateRawSync(utf8("abc"))), 33 * 1024 * 1024), previewError("resource-limit", /^DEFLATE output exceeds the configured limit\.$/));
});

// -----------------------------------------------------------------------------------------------
// 4. Parser model snapshots (spec §8.2)
// -----------------------------------------------------------------------------------------------

const parsers = {
  docx: (input, name) => parseDocx(input, name),
  pptx: (input, name) => parsePptx(input, name),
  xlsx: (input, name, options) => options === undefined ? parseXlsx(input, name) : parseXlsx(input, name, options),
};

for (const fixture of ooxmlFixtures) {
  test(`M0 snapshot: ${fixture.id} (${fixture.source})`, () => {
    const parse = () => parsers[fixture.kind](buildFixture(fixture), fixture.file, fixture.options);
    const first = parse();
    // The parser must be deterministic for a snapshot to be meaningful.
    assert.deepStrictEqual(normalize(first), normalize(parse()));
    matchSnapshot(`model-${fixture.id}`, first);
  });
}

test("M0 snapshot: demo application samples (scripts/demo-samples.mjs)", async () => {
  const directory = mkdtempSync(join(tmpdir(), "ufp-m0-"));
  try {
    await writeDemoSamples(directory);
    const sample = name => new Uint8Array(readFileSync(join(directory, name)));
    const docx = sample("sample-report.docx"), xlsx = sample("sample-workbook.xlsx"), pptx = sample("sample-deck.pptx");
    assert.equal(detectFormat("sample-report.docx", "", docx), "docx");
    assert.equal(detectFormat("sample-workbook.xlsx", "", xlsx), "xlsx");
    assert.equal(detectFormat("sample-deck.pptx", "", pptx), "pptx");
    assert.equal(detectFormat("sample-illustration.svg", "", sample("sample-illustration.svg")), "svg");
    const models = {
      "demo-docx": () => parseDocx(docx, "sample-report.docx"),
      "demo-xlsx": () => parseXlsx(xlsx, "sample-workbook.xlsx"),
      "demo-xlsx-de": () => parseXlsx(xlsx, "sample-workbook.xlsx", { locale: "de-DE" }),
      "demo-pptx": () => parsePptx(pptx, "sample-deck.pptx"),
    };
    for (const [id, parse] of Object.entries(models)) {
      const first = parse();
      assert.deepStrictEqual(normalize(first), normalize(parse()), `${id} is deterministic`);
      matchSnapshot(`model-${id}`, first);
    }
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});
