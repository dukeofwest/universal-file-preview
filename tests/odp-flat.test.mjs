import test from "node:test";
import assert from "node:assert/strict";
import { parseFodp } from "../dist/parsers/odf/odp-flat.js";
import { parseOdp } from "../dist/parsers/odf/odp.js";
import { makeFodp, makeOdp, flatOdpXml, ODP_MEDIA_TYPE, OTP_MEDIA_TYPE } from "./helpers/odf.mjs";

// Milestone 4.5: a flat presentation (.fodp) is one XML file holding what a package splits across
// parts. The same reader handles both, so a flat document and its packaged twin must agree.

const EMU = 12_700;
const textOf = paragraph => paragraph.runs.map(run => run.text).join("");
const textElements = slide => (slide.elements ?? []).filter(element => element.kind === "text");
const frame = lines => `<draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="3cm"><draw:text-box>${lines.map(line => `<text:p>${line}</text:p>`).join("")}</draw:text-box></draw:frame>`;
const page = (body, name = "page1") => `<draw:page draw:name="${name}">${body}</draw:page>`;
const failsWith = (code, pattern) => error => error?.name === "PreviewError" && error.code === code && pattern.test(error.message);

test("a flat presentation reads its slides, text and styles", () => {
  const styles = '<style:style style:name="Strong" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>';
  const body = page(`<draw:frame svg:x="0cm" svg:y="0cm" svg:width="5cm" svg:height="5cm"><draw:text-box><text:p>plain <text:span text:style-name="Strong">bold</text:span></text:p></draw:text-box></draw:frame>`);
  const model = parseFodp(makeFodp({ styles, body }), "deck.fodp");
  assert.equal(model.kind, "presentation");
  assert.equal(model.slides.length, 1);
  const [paragraph] = textElements(model.slides[0])[0].paragraphs;
  assert.equal(textOf(paragraph), "plain bold");
  assert.equal(paragraph.runs.find(run => run.text === "bold").bold, true, "styles.xml's styles apply in a flat document too");
  assert.equal(model.source.format, "fodp");
});

test("a flat presentation and its packaged twin produce the same slides", () => {
  const styleAutomatic = '<style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="25.4cm" fo:page-height="14.29cm"/></style:page-layout>';
  const body = page(frame(["Title", "Body"])) + page(frame(["Second"]), "page2");
  const flat = parseFodp(makeFodp({ automatic: styleAutomatic, body }), "deck.fodp");
  const packaged = parseOdp(makeOdp({ styleAutomatic, body }), "deck.odp");
  const shape = model => model.slides.map(slide => ({
    width: slide.width, height: slide.height, title: slide.title,
    text: textElements(slide).map(element => element.paragraphs.map(textOf)),
  }));
  assert.deepEqual(shape(flat), shape(packaged), "the same document in both forms reads the same");
  assert.equal(flat.slides[0].width, Math.round(25.4 * (72 / 2.54) * EMU));
});

test("a flat template is reported as shown-as-saved", () => {
  const model = parseFodp(makeFodp({ mimetype: OTP_MEDIA_TYPE, body: page(frame(["Draft"])) }), "deck.fodp");
  assert.ok(model.diagnostics.some(entry => entry.code === "odf.template"));
  assert.equal(model.source.format, "fodp", "the format is how it is previewed, not what it declares");
});

test("a flat document of another kind is refused by its declared media type", () => {
  const graphics = new TextEncoder().encode(flatOdpXml({ mimetype: "application/vnd.oasis.opendocument.graphics", body: page(frame(["x"])) }));
  assert.throws(() => parseFodp(graphics, "diagram.fodg"), failsWith("unsupported-format", /not an opendocument presentation/i));
  // A document that declares no media type at all is read as what the caller asked for.
  const undeclared = new TextEncoder().encode(flatOdpXml({ body: page(frame(["Fine"])) }).replace(` office:mimetype="${ODP_MEDIA_TYPE}"`, ""));
  assert.equal(parseFodp(undeclared, "deck.fodp").slides.length, 1);
});

test("a file that is not a flat OpenDocument document is refused", () => {
  assert.throws(() => parseFodp(new TextEncoder().encode('<?xml version="1.0"?><html><body>no</body></html>'), "x.fodp"),
    failsWith("invalid-file", /not a flat opendocument document/i));
  assert.throws(() => parseFodp(new TextEncoder().encode("not xml at all"), "x.fodp"), error => error?.name === "PreviewError");
});

test("a flat document larger than the limit is refused before it is read", () => {
  const bytes = makeFodp({ body: page(frame(["Big"])) });
  assert.throws(() => parseFodp(bytes, "deck.fodp", { limits: { maxContentBytes: 10 } }), failsWith("resource-limit", /size limit/i));
});

test("an image stored inline in a flat document is shown, typed from its bytes", () => {
  const png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==";
  const body = page(`<draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="3cm"><draw:image><office:binary-data>${png}</office:binary-data></draw:image></draw:frame>`);
  const model = parseFodp(makeFodp({ body }), "deck.fodp");
  const [element] = model.slides[0].elements;
  assert.equal(element.kind, "image");
  assert.equal(element.image.mimeType, "image/png", "inline data is typed from the bytes themselves");
  assert.deepEqual([...element.image.bytes.subarray(0, 4)], [0x89, 0x50, 0x4e, 0x47]);
});

test("inline data that is not an image is not shown", () => {
  const body = page('<draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="3cm"><draw:image><office:binary-data>bm90IGFuIGltYWdlIGF0IGFsbCE=</office:binary-data></draw:image></draw:frame>');
  const model = parseFodp(makeFodp({ body }), "deck.fodp");
  assert.equal((model.slides[0].elements ?? []).length, 0);
  assert.ok(model.diagnostics.some(entry => entry.code === "odf.unsupported.slide-content"));
});
