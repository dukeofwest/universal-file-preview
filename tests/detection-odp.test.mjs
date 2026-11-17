import test from "node:test";
import assert from "node:assert/strict";
import { detectFormat } from "../dist/detection/detect-format.js";
import { PackageReader } from "../dist/package/package-reader.js";
import { DiagnosticCollector } from "../dist/diagnostics.js";
import { readOdpPackage } from "../dist/parsers/odf/odp-package.js";
import { createDefaultRendererRegistry } from "../dist/renderers/registry.js";
import { makeZip } from "./helpers/zip.mjs";
import { makeOdfPackage, odfXml, minimalContent, ODS_MEDIA_TYPE, ODT_MEDIA_TYPE } from "./helpers/odf.mjs";

// Milestone 4.1: OpenDocument presentation detection and package validation. No parser or renderer
// exists yet, so a detected presentation is refused with "unsupported-format" exactly as PDF is.

const ODP = "application/vnd.oasis.opendocument.presentation";
const OTP = "application/vnd.oasis.opendocument.presentation-template";
const ODG = "application/vnd.oasis.opendocument.graphics";
const OTG = "application/vnd.oasis.opendocument.graphics-template";
const content = '<office:document-content office:version="1.3"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>';

const odp = makeOdfPackage({ mimetype: ODP, content });
const otp = makeOdfPackage({ mimetype: OTP, content });
const odg = makeOdfPackage({ mimetype: ODG, content });
const otg = makeOdfPackage({ mimetype: OTG, content });
const odt = makeOdfPackage({ mimetype: ODT_MEDIA_TYPE, content: '<office:document-content office:version="1.3"><office:body><office:text><text:p>Hello</text:p></office:text></office:body></office:document-content>' });
const ods = makeOdfPackage({ content: minimalContent });

/** A flat presentation, with the media type given so the precedence rules can be exercised. */
const flatPresentation = (mimetype = ODP, extra = "") => new TextEncoder().encode(odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document office:version="1.3" office:mimetype="${mimetype}">${extra}<office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document>`));

test("presentation packages are detected from their mimetype entry, whatever the file is called", () => {
  assert.equal(detectFormat("deck.odp", "", odp), "odp");
  assert.equal(detectFormat("Deck.ODP", "", odp), "odp");
  assert.equal(detectFormat("deck.otp", "", otp), "otp");
  assert.equal(detectFormat("deck.zip", "", odp), "odp");
  assert.equal(detectFormat("upload.bin", "application/octet-stream", odp), "odp");
  assert.equal(detectFormat("deck.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", odp), "odp");
  assert.equal(detectFormat("template.odp", "", otp), "otp", "the package decides, not the extension");
});

test("the drawing kinds this library does not read stay unknown", () => {
  // Drawings and drawing templates have their own media types and their own body element
  // (office:drawing), so they are left alone until there is a reader for them.
  assert.equal(detectFormat("diagram.odg", "", odg), "unknown");
  assert.equal(detectFormat("diagram.otg", "", otg), "unknown");
  assert.equal(detectFormat("diagram.odg", "", undefined), "unknown");
  assert.equal(detectFormat("x.zip", "", makeOdfPackage({ mimetype: "application/vnd.oasis.opendocument.chart", content })), "unknown");
});

test("presentations are detected from the MIME type or the extension alone", () => {
  assert.equal(detectFormat("deck.odp"), "odp");
  assert.equal(detectFormat("deck.otp"), "otp");
  assert.equal(detectFormat("deck.fodp"), "fodp");
  assert.equal(detectFormat("upload.bin", ODP), "odp");
  assert.equal(detectFormat("upload.bin", OTP), "otp");
  assert.equal(detectFormat("upload.bin", `${ODP}; charset=utf-8`), "odp", "parameters are ignored");
  assert.equal(detectFormat("upload.bin", ODG), "unknown");
});

test("a package whose mimetype entry cannot be trusted falls back to the extension", () => {
  const variants = {
    "compressed mimetype": makeOdfPackage({ mimetype: ODP, content, storeMimetype: false }),
    "mimetype not first": makeOdfPackage({ mimetype: ODP, content, order: ["content.xml", "mimetype"] }),
  };
  for (const [label, bytes] of Object.entries(variants)) {
    assert.equal(detectFormat("x.zip", "", bytes), "unknown", label);
    assert.equal(detectFormat("x.odp", "", bytes), "odp", `${label}: falls back to the extension`);
  }
});

test("a truncated presentation package never throws and still resolves", () => {
  for (let length = 0; length <= odp.length; length += 97) {
    const result = detectFormat("deck.odp", "", odp.subarray(0, length));
    assert.ok(result === "odp" || result === "unknown", `length ${length} gave ${result}`);
  }
});

test("flat presentations are detected by their declared media type", () => {
  assert.equal(detectFormat("deck.fodp", "", flatPresentation()), "fodp");
  assert.equal(detectFormat("deck.fodp", "", flatPresentation(OTP)), "fodp", "a flat template is previewed as fodp");
  assert.equal(detectFormat("upload.bin", ODP, flatPresentation()), "fodp", "the MIME type also opens the flat check");
  // The media type is matched as a whole attribute value, so the graphics kinds are not claimed.
  assert.equal(detectFormat("diagram.bin", ODG, flatPresentation(ODG)), "unknown");
  assert.equal(detectFormat("diagram.bin", "", flatPresentation(OTG)), "unknown");
});

test("each flat check keeps to its own format", () => {
  const flatText = new TextEncoder().encode(odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document office:version="1.3" office:mimetype="${ODT_MEDIA_TYPE}"><office:body><office:text><text:p>Hello</text:p></office:text></office:body></office:document>`));
  const flatOds = new TextEncoder().encode(odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document office:version="1.3" office:mimetype="${ODS_MEDIA_TYPE}"><office:body><office:spreadsheet/></office:body></office:document>`));
  assert.equal(detectFormat("deck.fodp", "", flatPresentation()), "fodp");
  // A file named for one flat format but holding another resolves by its extension, which is what
  // any unreadable file does; the parser refuses it once it reads the declared media type.
  assert.equal(detectFormat("deck.fodp", "", flatText), "fodp", "the presentation check does not match text content");
  assert.equal(detectFormat("letter.fodt", "", flatPresentation()), "fodt", "the text check does not match presentation content");
  assert.equal(detectFormat("sheet.fods", "", flatPresentation()), "fods");
  assert.equal(detectFormat("deck.bin", "", flatPresentation()), "unknown", "with neither name nor MIME type, no flat check runs");
  assert.equal(detectFormat("deck.bin", "", flatOds), "unknown");
});

test("signatures and other formats keep precedence over the flat presentation check", () => {
  const pdfLead = new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, ...flatPresentation()]);
  assert.equal(detectFormat("deck.fodp", "", pdfLead), "pdf");
  assert.equal(detectFormat("deck.fodp", "", new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>')), "svg");
  // A PDF marker further inside a flat presentation must not beat the flat check.
  assert.equal(detectFormat("deck.fodp", "", flatPresentation(ODP, "<office:meta>%PDF-1.7</office:meta>")), "fodp");
  assert.equal(detectFormat("deck.fodp", "", new TextEncoder().encode("plain text, not markup")), "fodp", "an unreadable file still resolves by extension");
});

test("a flat presentation is detected whichever quote character delimits its media type", () => {
  const single = new TextEncoder().encode(odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document office:version="1.3" office:mimetype='${ODP}'><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document>`));
  assert.equal(detectFormat("deck.fodp", "", single), "fodp");
  assert.equal(detectFormat("upload.bin", ODP, single), "fodp", "a flat document is not taken for a package");
});

test("every presentation format has a renderer, and PDF still has none", () => {
  const registry = createDefaultRendererRegistry();
  assert.ok(registry.get("odp"), "the presentation renderer is registered");
  assert.equal(registry.get("otp"), registry.get("odp"), "one renderer instance serves them all");
  assert.equal(registry.get("fodp"), registry.get("odp"), "the flat format shares that renderer");
  assert.equal(registry.get("pdf"), undefined);
  for (const format of ["odt", "ott", "fodt", "ods", "ots", "fods", "xlsx", "docx", "pptx", "png", "svg"]) assert.ok(registry.get(format), `${format} is still registered`);
});

/** Opens a package the way a parser would, so the validator can be exercised on its own. */
function readPackage(bytes) {
  return readOdpPackage(PackageReader.openZip(bytes), new DiagnosticCollector());
}
const failsWith = (code, pattern) => error => error?.name === "PreviewError" && error.code === code && pattern.test(error.message);

test("the package validator accepts presentations and names the format from the media type", () => {
  assert.equal(readPackage(odp).format, "odp");
  assert.equal(readPackage(otp).format, "otp");
  assert.equal(readPackage(odp).info.mediaType, ODP);
});

test("the package validator refuses every other OpenDocument kind", () => {
  for (const [label, bytes] of [["spreadsheet", ods], ["text document", odt], ["drawing", odg], ["drawing template", otg]]) {
    assert.throws(() => readPackage(bytes), failsWith("unsupported-format", /not an opendocument presentation/i), label);
  }
});

test("the package validator requires the content stream and a consistent package", () => {
  assert.throws(() => readPackage(makeOdfPackage({ mimetype: ODP })), failsWith("invalid-file", /no content\.xml/i));
  assert.throws(() => readPackage(makeOdfPackage({ mimetype: ODP, content, manifest: `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="${ODS_MEDIA_TYPE}"/></manifest:manifest>` })), failsWith("invalid-file", /media type is inconsistent/i));
  assert.throws(() => readPackage(new Uint8Array(makeZip({ "content.xml": content }))), failsWith("invalid-file", /does not declare its media type/i));
});

test("the package validator refuses password-protected presentations before reading content", () => {
  const manifest = `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="${ODP}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data/></manifest:file-entry></manifest:manifest>`;
  assert.throws(() => readPackage(makeOdfPackage({ mimetype: ODP, content, manifest })), failsWith("unsupported-feature", /password-protected/i));
});

test("the validator's messages never quote text taken from the document", () => {
  const spoof = "SPOOF: your session has expired";
  const bytes = makeOdfPackage({ mimetype: `${ODP}-${spoof}`, content });
  const error = (() => { try { readPackage(bytes); return undefined; } catch (reason) { return reason; } })();
  assert.ok(error, "an unknown media type must be refused");
  assert.equal(error.message.includes(spoof), false, error.message);
});
