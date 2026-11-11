import test from "node:test";
import assert from "node:assert/strict";
import { detectFormat } from "../dist/detection/detect-format.js";
import { PackageReader } from "../dist/package/package-reader.js";
import { DiagnosticCollector } from "../dist/diagnostics.js";
import { readOdtPackage } from "../dist/parsers/odf/odt-package.js";
import { createDefaultRendererRegistry } from "../dist/renderers/registry.js";
import { makeZip } from "./helpers/zip.mjs";
import { makeOdfPackage, odfXml, minimalContent, ODS_MEDIA_TYPE } from "./helpers/odf.mjs";

// Milestone 3.1: OpenDocument text detection and package validation. No parser or renderer exists
// yet, so a detected document is refused with "unsupported-format" exactly as PDF is.

const ODT = "application/vnd.oasis.opendocument.text";
const OTT = "application/vnd.oasis.opendocument.text-template";
const ODM = "application/vnd.oasis.opendocument.text-master";
const OTH = "application/vnd.oasis.opendocument.text-web";
const ODP = "application/vnd.oasis.opendocument.presentation";
const textContent = '<office:document-content office:version="1.3"><office:body><office:text><text:p>Hello</text:p></office:text></office:body></office:document-content>';

const odt = makeOdfPackage({ mimetype: ODT, content: textContent });
const ott = makeOdfPackage({ mimetype: OTT, content: textContent });
const odm = makeOdfPackage({ mimetype: ODM, content: textContent });
const oth = makeOdfPackage({ mimetype: OTH, content: textContent });
const ods = makeOdfPackage({ content: minimalContent });

/** A flat text document, with the media type given so the precedence rules can be exercised. */
const flatText = (mimetype = ODT, extra = "") => new TextEncoder().encode(odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document office:version="1.3" office:mimetype="${mimetype}">${extra}<office:body><office:text><text:p>Hello</text:p></office:text></office:body></office:document>`));

test("text packages are detected from their mimetype entry, whatever the file is called", () => {
  assert.equal(detectFormat("letter.odt", "", odt), "odt");
  assert.equal(detectFormat("Letter.ODT", "", odt), "odt");
  assert.equal(detectFormat("letter.ott", "", ott), "ott");
  assert.equal(detectFormat("letter.zip", "", odt), "odt");
  assert.equal(detectFormat("upload.bin", "application/octet-stream", odt), "odt");
  assert.equal(detectFormat("letter.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", odt), "odt");
  assert.equal(detectFormat("template.odt", "", ott), "ott", "the package decides, not the extension");
});

test("the text kinds this library does not read stay unknown", () => {
  // Master documents and web templates have their own media types, which are prefixed by the text
  // one: matching by prefix would claim them. They are left alone until there is a reader for them.
  assert.equal(detectFormat("book.odm", "", odm), "unknown");
  assert.equal(detectFormat("page.oth", "", oth), "unknown");
  assert.equal(detectFormat("book.odm", "", undefined), "unknown");
  assert.equal(detectFormat("x.zip", "", makeOdfPackage({ mimetype: `${ODT}-master-template`, content: textContent })), "unknown");
  assert.equal(detectFormat("deck.odp", "", makeOdfPackage({ mimetype: ODP, content: textContent })), "unknown", "presentations arrive in phase 4");
});

test("text documents are detected from the MIME type or the extension alone", () => {
  assert.equal(detectFormat("letter.odt"), "odt");
  assert.equal(detectFormat("letter.ott"), "ott");
  assert.equal(detectFormat("letter.fodt"), "fodt");
  assert.equal(detectFormat("upload.bin", ODT), "odt");
  assert.equal(detectFormat("upload.bin", OTT), "ott");
  assert.equal(detectFormat("upload.bin", `${ODT}; charset=utf-8`), "odt", "parameters are ignored");
  assert.equal(detectFormat("upload.bin", ODM), "unknown");
  // Names and MIME types are untrusted: inherited properties must not resolve.
  assert.equal(detectFormat("x.constructor"), "unknown");
  assert.equal(detectFormat("x.bin", "__proto__"), "unknown");
});

test("a package whose mimetype entry cannot be trusted falls back to the extension", () => {
  // Same matrix as the spreadsheet rules: the entry must be stored, first and without a descriptor.
  const variants = {
    "compressed mimetype": makeOdfPackage({ mimetype: ODT, content: textContent, storeMimetype: false }),
    "mimetype not first": makeOdfPackage({ mimetype: ODT, content: textContent, order: ["content.xml", "mimetype"] }),
  };
  for (const [label, bytes] of Object.entries(variants)) {
    assert.equal(detectFormat("x.zip", "", bytes), "unknown", label);
    assert.equal(detectFormat("x.odt", "", bytes), "odt", `${label}: falls back to the extension`);
  }
});

test("a truncated text package never throws and still resolves", () => {
  for (let length = 0; length <= odt.length; length += 97) {
    const result = detectFormat("letter.odt", "", odt.subarray(0, length));
    assert.ok(result === "odt" || result === "unknown", `length ${length} gave ${result}`);
  }
});

test("flat text documents are detected by their declared media type", () => {
  assert.equal(detectFormat("letter.fodt", "", flatText()), "fodt");
  assert.equal(detectFormat("letter.fodt", "", flatText(OTT)), "fodt", "a flat template is previewed as fodt");
  assert.equal(detectFormat("upload.bin", ODT, flatText()), "fodt", "the MIME type also opens the flat check");
  // The media type is matched as a whole attribute value, so the longer text kinds are not claimed.
  assert.equal(detectFormat("book.fodt", "", flatText(ODM)), "fodt", "the extension still decides when the content is another kind");
  assert.equal(detectFormat("book.bin", ODM, flatText(ODM)), "unknown");
  assert.equal(detectFormat("book.bin", "", flatText(OTH)), "unknown");
});

test("flat spreadsheets and flat text documents do not take each other's files", () => {
  const flatOds = new TextEncoder().encode(odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document office:version="1.3" office:mimetype="${ODS_MEDIA_TYPE}"><office:body><office:spreadsheet/></office:body></office:document>`));
  assert.equal(detectFormat("sheet.fods", "", flatOds), "fods");
  assert.equal(detectFormat("letter.fodt", "", flatText()), "fodt");
  // Each flat check only runs when the name or the MIME type points at that format, so neither ever
  // claims the other's content; a file named for the wrong one resolves by its extension, which is
  // what any unreadable file does. The parser refuses it once it reads the declared media type.
  assert.equal(detectFormat("sheet.fodt", "", flatOds), "fodt", "the text check does not match spreadsheet content");
  assert.equal(detectFormat("letter.fods", "", flatText()), "fods", "the spreadsheet check does not match text content");
  assert.equal(detectFormat("sheet.bin", "", flatOds), "unknown", "with neither name nor MIME type, no flat check runs");
  assert.equal(detectFormat("letter.bin", "", flatText()), "unknown");
});

test("signatures and other formats keep precedence over the flat text check", () => {
  const pdfLead = new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, ...flatText()]);
  assert.equal(detectFormat("letter.fodt", "", pdfLead), "pdf");
  assert.equal(detectFormat("letter.fodt", "", new TextEncoder().encode("<svg xmlns=\"http://www.w3.org/2000/svg\"/>")), "svg");
  // A PDF marker further inside a flat text document must not beat the flat check.
  assert.equal(detectFormat("letter.fodt", "", flatText(ODT, "<office:meta>%PDF-1.7</office:meta>")), "fodt");
  assert.equal(detectFormat("letter.fodt", "", new TextEncoder().encode("plain text, not markup")), "fodt", "an unreadable file still resolves by extension");
});

test("every text format has a renderer, and PDF still has none", () => {
  const registry = createDefaultRendererRegistry();
  assert.ok(registry.get("odt"), "the text renderer is registered");
  assert.equal(registry.get("ott"), registry.get("odt"), "one renderer instance serves them all");
  assert.equal(registry.get("fodt"), registry.get("odt"));
  assert.equal(registry.get("pdf"), undefined);
  for (const format of ["ods", "ots", "fods", "xlsx", "docx", "pptx", "png", "svg"]) assert.ok(registry.get(format), `${format} is still registered`);
});

/** Opens a package the way a parser would, so the validator can be exercised on its own. */
function readPackage(bytes) {
  return readOdtPackage(PackageReader.openZip(bytes), new DiagnosticCollector());
}
const failsWith = (code, pattern) => error => error?.name === "PreviewError" && error.code === code && pattern.test(error.message);

test("the package validator accepts text documents and names the format from the media type", () => {
  assert.equal(readPackage(odt).format, "odt");
  assert.equal(readPackage(ott).format, "ott");
  assert.equal(readPackage(odt).info.mediaType, ODT);
});

test("the package validator refuses every other OpenDocument kind", () => {
  for (const [label, bytes] of [["spreadsheet", ods], ["master document", odm], ["web template", oth]]) {
    assert.throws(() => readPackage(bytes), failsWith("unsupported-format", /not an opendocument text document/i), label);
  }
});

test("the package validator requires the content stream and a consistent package", () => {
  assert.throws(() => readPackage(makeOdfPackage({ mimetype: ODT })), failsWith("invalid-file", /no content\.xml/i));
  assert.throws(() => readPackage(makeOdfPackage({ mimetype: ODT, content: textContent, manifest: `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="${ODS_MEDIA_TYPE}"/></manifest:manifest>` })), failsWith("invalid-file", /media type is inconsistent/i));
  assert.throws(() => readPackage(new Uint8Array(makeZip({ "content.xml": textContent }))), failsWith("invalid-file", /does not declare its media type/i));
});

test("the package validator refuses password-protected documents before reading content", () => {
  const manifest = `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="${ODT}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data/></manifest:file-entry></manifest:manifest>`;
  assert.throws(() => readPackage(makeOdfPackage({ mimetype: ODT, content: textContent, manifest })), failsWith("unsupported-feature", /password-protected/i));
});

test("the validator's messages never quote text taken from the document", () => {
  const spoof = "SPOOF: your session has expired";
  const bytes = makeOdfPackage({ mimetype: `${ODT}-${spoof}`, content: textContent });
  const error = (() => { try { readPackage(bytes); return undefined; } catch (reason) { return reason; } })();
  assert.ok(error, "an unknown media type must be refused");
  assert.equal(error.message.includes(spoof), false, error.message);
});
