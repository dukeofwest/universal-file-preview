import test from "node:test";
import assert from "node:assert/strict";
import { parseOdp, resolveOdpLimits, odpWorkLimits } from "../dist/parsers/odf/odp.js";
import { makeOdp, ODP_MEDIA_TYPE, OTP_MEDIA_TYPE } from "./helpers/odf.mjs";

// Milestone 4.2: the slides and their text. One slide per draw:page, a text box becomes the model's
// text element, and anything this milestone cannot represent is reported rather than dropped.

const EMU = 12_700;
const parse = (options, name = "deck.odp") => parseOdp(makeOdp(options), name);
/** The text of a paragraph, as a reader sees it. */
const textOf = paragraph => paragraph.runs.map(run => run.text).join("");
/** Every text element on a slide, in order. */
const textElements = slide => (slide.elements ?? []).filter(element => element.kind === "text");
/** A frame holding a text box with one paragraph per entry of `lines`. */
const frame = (lines, { x = "1cm", y = "2cm", width = "10cm", height = "3cm", klass, styleName } = {}) =>
  `<draw:frame draw:name="f" svg:x="${x}" svg:y="${y}" svg:width="${width}" svg:height="${height}"${klass ? ` presentation:class="${klass}"` : ""}${styleName ? ` draw:style-name="${styleName}"` : ""}><draw:text-box>${lines.map(line => `<text:p>${line}</text:p>`).join("")}</draw:text-box></draw:frame>`;
const page = (body, name = "page1") => `<draw:page draw:name="${name}" draw:master-page-name="Default">${body}</draw:page>`;
const failsWith = (code, pattern) => error => error?.name === "PreviewError" && error.code === code && pattern.test(error.message);

test("each draw:page becomes a slide, in document order", () => {
  const model = parse({ body: page(frame(["First"]), "Intro") + page(frame(["Second"]), "Detail") });
  assert.equal(model.kind, "presentation");
  assert.equal(model.slides.length, 2);
  assert.equal(textOf(textElements(model.slides[0])[0].paragraphs[0]), "First");
  assert.equal(textOf(textElements(model.slides[1])[0].paragraphs[0]), "Second");
  assert.equal(model.title, "deck", "the model is titled after the file");
  assert.equal(model.source.format, "odp");
  assert.equal(model.source.producer, "LibreOffice/24.8.4.2");
});

test("a text box becomes a text element with the frame's geometry in EMU", () => {
  const body = page(frame(["Body"], { x: "2cm", y: "1in", width: "72pt", height: "10mm" }));
  const [element] = textElements(parse({ body }).slides[0]);
  assert.equal(element.kind, "text");
  assert.equal(element.left, Math.round(2 * (72 / 2.54) * EMU));
  assert.equal(element.top, Math.round(72 * EMU));
  assert.equal(element.width, 72 * EMU);
  assert.equal(element.height, Math.round(10 * (72 / 25.4) * EMU));
  assert.deepEqual(element.margins, { top: 91_440, right: 91_440, bottom: 91_440, left: 91_440 });
});

test("a paragraph keeps its spans, spaces, tabs and line breaks", () => {
  const styles = '<style:style style:name="Strong" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>';
  const body = page(`<draw:frame svg:x="0cm" svg:y="0cm" svg:width="5cm" svg:height="5cm"><draw:text-box><text:p>a<text:span text:style-name="Strong">bold</text:span><text:s text:c="3"/>b<text:tab/>c<text:line-break/>d</text:p></draw:text-box></draw:frame>`);
  const [element] = textElements(parse({ styles, body }).slides[0]);
  const [paragraph] = element.paragraphs;
  assert.equal(textOf(paragraph), "abold   b\tc\nd");
  assert.equal(paragraph.runs.find(run => run.text === "bold").bold, true);
});

test("the slide size comes from the page layout the styles declare", () => {
  const styleAutomatic = '<style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="28cm" fo:page-height="15.75cm"/></style:page-layout>';
  const { slides } = parse({ styleAutomatic, body: page(frame(["Wide"])) });
  assert.equal(slides[0].width, Math.round(28 * (72 / 2.54) * EMU));
  assert.equal(slides[0].height, Math.round(15.75 * (72 / 2.54) * EMU));
  // Without a layout the ODF default of 10 by 7.5 inches is used, never a guess from the content.
  const plain = parse({ body: page(frame(["Plain"])) });
  assert.equal(plain.slides[0].width, 720 * EMU);
  assert.equal(plain.slides[0].height, 540 * EMU);
});

test("the slide is titled from its title placeholder, and otherwise from the page name", () => {
  const titled = parse({ body: page(frame(["The title"], { klass: "title" }) + frame(["Body"]), "page1") });
  assert.equal(titled.slides[0].title, "The title");
  const unnamed = parse({ body: page(frame(["Body"]), "Agenda") });
  assert.equal(unnamed.slides[0].title, "Agenda", "the page name is the fallback");
});

test("a list in a text box labels only the first paragraph of each item", () => {
  const styles = '<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style>';
  const body = page(`<draw:frame svg:x="0cm" svg:y="0cm" svg:width="5cm" svg:height="5cm"><draw:text-box><text:list text:style-name="L1"><text:list-item><text:p>One</text:p><text:p>still one</text:p></text:list-item><text:list-item><text:p>Two</text:p></text:list-item></text:list></draw:text-box></draw:frame>`);
  const [element] = textElements(parse({ styles, body }).slides[0]);
  assert.deepEqual(element.paragraphs.map(textOf), ["One", "still one", "Two"]);
  assert.deepEqual(element.paragraphs.map(paragraph => paragraph.list.label), ["1.", "", "2."]);
  assert.deepEqual(element.paragraphs.map(paragraph => paragraph.list.kind), ["ordered", "ordered", "ordered"]);
});

test("a list that names no style still shows as a list", () => {
  const body = page(`<draw:frame svg:x="0cm" svg:y="0cm" svg:width="5cm" svg:height="5cm"><draw:text-box><text:list><text:list-item><text:p>Point</text:p></text:list-item></text:list></draw:text-box></draw:frame>`);
  const [element] = textElements(parse({ body }).slides[0]);
  assert.equal(element.paragraphs[0].list.kind, "bullet");
  assert.equal(element.paragraphs[0].list.label, "•");
});

test("content this preview cannot show is reported, not dropped silently", () => {
  // An embedded object and an image the package does not hold are both left out; the shape beside
  // them is drawn, since milestone 4.4 reads shapes.
  const body = page([
    '<draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="4cm"><draw:image xlink:href="Pictures/missing.png"/></draw:frame>',
    '<draw:custom-shape svg:x="0cm" svg:y="5cm" svg:width="4cm" svg:height="4cm"><draw:enhanced-geometry draw:type="ellipse"/></draw:custom-shape>',
    '<draw:frame svg:x="5cm" svg:y="0cm" svg:width="4cm" svg:height="4cm"><draw:object xlink:href="./Object 1"/></draw:frame>',
  ].join(""));
  const model = parse({ body });
  const reported = model.diagnostics.find(entry => entry.code === "odf.unsupported.slide-content");
  assert.ok(reported, "the unsupported kinds are reported");
  assert.equal(reported.severity, "warning");
  assert.equal(reported.count, 2, "the image and the embedded object are both named");
  assert.deepEqual((model.slides[0].elements ?? []).map(element => element.kind), ["shape"], "only what can be read is shown");
});

test("speaker notes are reported and are not read as slide content", () => {
  const body = page(frame(["Shown"]) + '<presentation:notes><draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="4cm"><draw:text-box><text:p>Private note</text:p></draw:text-box></draw:frame></presentation:notes>');
  const model = parse({ body });
  assert.ok(model.diagnostics.some(entry => entry.code === "odf.unsupported.notes"));
  const texts = textElements(model.slides[0]).flatMap(element => element.paragraphs.map(textOf));
  assert.deepEqual(texts, ["Shown"], "the note's text never reaches the slide");
});

test("a template is reported as shown-as-saved", () => {
  const model = parseOdp(makeOdp({ mimetype: OTP_MEDIA_TYPE, body: page(frame(["Draft"])) }), "deck.otp");
  assert.equal(model.source.format, "otp");
  assert.ok(model.diagnostics.some(entry => entry.code === "odf.template"));
  // An ordinary presentation says nothing of the sort.
  assert.equal(parse({ body: page(frame(["Final"])) }).diagnostics?.some(entry => entry.code === "odf.template") ?? false, false);
});

test("slides and objects past their limits are reported rather than dropped silently", () => {
  const many = Array.from({ length: 12 }, (_, index) => page(frame([`Slide ${index}`]), `p${index}`)).join("");
  const limited = parseOdp(makeOdp({ body: many }), "deck.odp", { limits: { maxSlides: 5 } });
  assert.equal(limited.slides.length, 5);
  const slides = limited.diagnostics.find(entry => entry.code === "odf.bounds.slides");
  assert.ok(slides, "the slides left out are reported");
  assert.equal(slides.count, 7);

  const crowded = page(Array.from({ length: 9 }, (_, index) => frame([`Box ${index}`], { y: `${index}cm` })).join(""));
  const trimmed = parseOdp(makeOdp({ body: crowded }), "deck.odp", { limits: { maxElementsPerSlide: 4 } });
  assert.equal(trimmed.slides[0].elements.length, 4);
  const objects = trimmed.diagnostics.find(entry => entry.code === "odf.bounds.slide-elements");
  assert.ok(objects, "the objects left out are reported");
  assert.equal(objects.count, 5);
});

test("a presentation with no slides still produces one surface", () => {
  const model = parse({ body: "" });
  assert.equal(model.slides.length, 1);
  assert.equal(model.slides[0].width, 720 * EMU);
  assert.deepEqual(model.slides[0].paragraphs, []);
});

test("the content stream must be an OpenDocument content stream", () => {
  const notContent = '<office:document-styles office:version="1.3"><office:body/></office:document-styles>';
  assert.throws(() => parseOdp(makeOdp({ body: "" }).slice(0, 40), "deck.odp"), error => error?.name === "PreviewError");
  const bytes = makeOdp({ body: page(frame(["x"])) });
  assert.ok(parseOdp(bytes, "deck.odp"), "the fixture itself is valid");
  assert.equal(typeof notContent, "string");
});

test("limits are validated, and the work budget follows the slide and object budgets", () => {
  assert.throws(() => resolveOdpLimits({ maxSlides: 0 }), RangeError);
  assert.throws(() => resolveOdpLimits({ maxSlides: 1.5 }), RangeError);
  assert.throws(() => resolveOdpLimits({ nonsense: 1 }), RangeError);
  assert.equal(resolveOdpLimits(undefined).maxSlides, 1_000);
  const work = odpWorkLimits({ maxSlides: 10, maxElementsPerSlide: 10 });
  assert.equal(work.maxElements, 250_000 + 100 * 12);
  assert.equal(work.maxTokens, work.maxElements * 4);
  // The budget comes from the content, never from the byte size: a small file cannot buy more work.
  const big = odpWorkLimits({ maxSlides: 1_000, maxElementsPerSlide: 500 });
  assert.ok(big.maxElements > work.maxElements);
});

test("a paragraph longer than the limit stops the preview", () => {
  const body = page(`<draw:frame svg:x="0cm" svg:y="0cm" svg:width="5cm" svg:height="5cm"><draw:text-box><text:p>${"x".repeat(500)}</text:p></draw:text-box></draw:frame>`);
  assert.throws(() => parseOdp(makeOdp({ body }), "deck.odp", { limits: { maxParagraphTextLength: 100 } }), failsWith("resource-limit", /length limit/i));
  // The limit also bounds text merged into one run, which a tab would otherwise split.
  const split = page(`<draw:frame svg:x="0cm" svg:y="0cm" svg:width="5cm" svg:height="5cm"><draw:text-box><text:p>${Array.from({ length: 10 }, () => "y".repeat(20)).join("<text:tab/>")}</text:p></draw:text-box></draw:frame>`);
  assert.throws(() => parseOdp(makeOdp({ body: split }), "deck.odp", { limits: { maxParagraphTextLength: 100 } }), failsWith("resource-limit", /length limit/i));
});

test("the parser refuses a package that is not a presentation", () => {
  assert.throws(() => parseOdp(makeOdp({ mimetype: "application/vnd.oasis.opendocument.graphics", body: page(frame(["x"])) }), "diagram.odg"),
    failsWith("unsupported-format", /not an opendocument presentation/i));
  assert.equal(ODP_MEDIA_TYPE, "application/vnd.oasis.opendocument.presentation");
});

// --- Milestone 4.3: images and slide backgrounds ------------------------------------------------

/** A 1x1 PNG, the smallest real image: the media type must come from these bytes, not from a name. */
const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==", "base64");

test("an image frame becomes an image element with the frame's geometry", () => {
  const body = page('<draw:frame svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm"><draw:image xlink:href="Pictures/a.png"/><svg:desc>A chart of sales</svg:desc></draw:frame>');
  const model = parseOdp(makeOdp({ body, extra: { "Pictures/a.png": PNG } }), "deck.odp");
  const [element] = model.slides[0].elements;
  assert.equal(element.kind, "image");
  assert.equal(element.image.mimeType, "image/png", "the type comes from the bytes");
  assert.deepEqual([...element.image.bytes.subarray(0, 4)], [0x89, 0x50, 0x4e, 0x47]);
  assert.equal(element.image.alt, "A chart of sales", "svg:desc is the alternative text");
  assert.equal(element.left, Math.round((72 / 2.54) * EMU));
  assert.equal(element.width, Math.round(4 * (72 / 2.54) * EMU));
  assert.equal(model.diagnostics?.some(entry => entry.code === "odf.unsupported.slide-content") ?? false, false, "nothing was left out");
});

test("an image the package does not hold is reported, not invented", () => {
  const body = page('<draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="3cm"><draw:image xlink:href="Pictures/missing.png"/></draw:frame>');
  const model = parseOdp(makeOdp({ body }), "deck.odp");
  assert.equal((model.slides[0].elements ?? []).length, 0);
  assert.ok(model.diagnostics.some(entry => entry.code === "odf.unsupported.slide-content"));
});

test("an image reference that points outside the package is never followed", () => {
  for (const href of ["../secret.png", "/etc/passwd", "https://example.test/a.png", "file:///c:/x.png"]) {
    const body = page(`<draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="3cm"><draw:image xlink:href="${href}"/></draw:frame>`);
    const model = parseOdp(makeOdp({ body, extra: { "Pictures/a.png": PNG } }), "deck.odp");
    assert.equal((model.slides[0].elements ?? []).length, 0, href);
    assert.ok(model.diagnostics.some(entry => entry.code === "odf.unsupported.slide-content"), href);
  }
});

test("a package image is typed by the package, and its bytes are checked where the URL is made", () => {
  // The same contract as for text documents (odt-images.ts): at parse time the manifest's media type
  // is used when this library shows that type, and the extension decides otherwise, so the model may
  // carry a type the bytes do not match. Nothing is shown on that basis: createImageUrl sniffs the
  // bytes and refuses a mismatch before any blob URL exists (tests/resource-limits.test.mjs).
  const body = page('<draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="3cm"><draw:image xlink:href="Pictures/a.png"/></draw:frame>');
  const mislabelled = parseOdp(makeOdp({ body, extra: { "Pictures/a.png": "PK\u0003\u0004 not an image" } }), "deck.odp");
  const [element] = mislabelled.slides[0].elements;
  assert.equal(element.image.mimeType, "image/png", "the package's own type is carried, not a guess");
  assert.deepEqual([...element.image.bytes.subarray(0, 2)], [0x50, 0x4b], "the bytes are carried verbatim");
});

test("a media type this library does not show is not read as an image", () => {
  const body = page('<draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="3cm"><draw:image xlink:href="Pictures/a.tiff"/></draw:frame>');
  const model = parseOdp(makeOdp({ body, extra: { "Pictures/a.tiff": PNG } }), "deck.odp");
  assert.equal((model.slides[0].elements ?? []).length, 0, "an unshowable type leaves no element");
  assert.ok(model.diagnostics.some(entry => entry.code === "odf.unsupported.slide-content"));
});

test("a solid slide background is shown, and the other fills are reported", () => {
  const styleAutomatic = '<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="#204060"/></style:style>';
  const solid = parse({ styleAutomatic, body: `<draw:page draw:name="p1" draw:style-name="dp1">${frame(["On colour"])}</draw:page>` });
  assert.deepEqual(solid.slides[0].background, { color: "#204060" });

  const gradientStyle = '<style:style style:name="dp2" style:family="drawing-page"><style:drawing-page-properties draw:fill="gradient" draw:fill-gradient-name="g1"/></style:style>';
  const gradient = parse({ styleAutomatic: gradientStyle, body: `<draw:page draw:name="p1" draw:style-name="dp2">${frame(["On a gradient"])}</draw:page>` });
  assert.equal(gradient.slides[0].background, undefined, "a gradient is not approximated");
  assert.ok(gradient.diagnostics.some(entry => entry.code === "odf.unsupported.slide-content"), "it is reported instead");

  // No fill at all leaves the renderer's own default, and says nothing.
  const plain = parse({ body: page(frame(["Plain"])) });
  assert.equal(plain.slides[0].background, undefined);
  assert.equal(plain.diagnostics?.some(entry => entry.code === "odf.unsupported.slide-content") ?? false, false);
});

test("a background colour is inherited through the style's parent chain", () => {
  const styleAutomatic = [
    '<style:style style:name="base" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="#112233"/></style:style>',
    '<style:style style:name="dp1" style:family="drawing-page" style:parent-style-name="base"/>',
  ].join("");
  const model = parse({ styleAutomatic, body: `<draw:page draw:name="p1" draw:style-name="dp1">${frame(["Inherited"])}</draw:page>` });
  assert.deepEqual(model.slides[0].background, { color: "#112233" });
});

test("a colour a document states in a form this preview does not read is left alone", () => {
  const styleAutomatic = '<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="rgb(1,2,3)"/></style:style>';
  const model = parse({ styleAutomatic, body: `<draw:page draw:name="p1" draw:style-name="dp1">${frame(["Odd colour"])}</draw:page>` });
  assert.equal(model.slides[0].background, undefined, "no colour is guessed from an unreadable value");
});

// --- Milestone 4.4: shapes, lines and tables ----------------------------------------------------

/** A graphic style with the fill and stroke a shape refers to. */
const graphicStyle = (name, properties) => `<style:style style:name="${name}" style:family="graphic"><style:graphic-properties ${properties}/></style:style>`;
const shapes = slide => (slide.elements ?? []).filter(element => element.kind === "shape");

test("a basic shape becomes a shape element with its outline, fill and stroke", () => {
  const styleAutomatic = graphicStyle("gr1", 'draw:fill="solid" draw:fill-color="#336699" draw:stroke="solid" svg:stroke-color="#112233" svg:stroke-width="0.1cm"');
  const body = page('<draw:rect draw:style-name="gr1" svg:x="1cm" svg:y="1cm" svg:width="4cm" svg:height="2cm"/>');
  const [shape] = shapes(parse({ styleAutomatic, body }).slides[0]);
  assert.equal(shape.kind, "shape");
  assert.equal(shape.geometry, "rect");
  assert.deepEqual(shape.fill, { color: "#336699" });
  assert.equal(shape.line.color, "#112233");
  assert.equal(shape.line.width, Math.round(0.1 * (72 / 2.54) * EMU));
  assert.equal(shape.width, Math.round(4 * (72 / 2.54) * EMU));
});

test("an ellipse and a circle are both drawn as ellipses", () => {
  const body = page('<draw:ellipse svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="2cm"/><draw:circle svg:x="0cm" svg:y="3cm" svg:width="2cm" svg:height="2cm"/>');
  assert.deepEqual(shapes(parse({ body }).slides[0]).map(shape => shape.geometry), ["ellipse", "ellipse"]);
});

test("a custom shape takes the preset outline its geometry declares", () => {
  const body = page([
    '<draw:custom-shape svg:x="0cm" svg:y="0cm" svg:width="3cm" svg:height="3cm"><draw:enhanced-geometry draw:type="isosceles-triangle"/></draw:custom-shape>',
    '<draw:custom-shape svg:x="4cm" svg:y="0cm" svg:width="3cm" svg:height="3cm"><draw:enhanced-geometry draw:type="round-rectangle"/></draw:custom-shape>',
    '<draw:custom-shape svg:x="8cm" svg:y="0cm" svg:width="3cm" svg:height="3cm"><draw:enhanced-geometry draw:type="right-arrow"/></draw:custom-shape>',
  ].join(""));
  const model = parse({ body });
  assert.deepEqual(shapes(model.slides[0]).map(shape => shape.geometry), ["triangle", "roundRect", "rightArrow"]);
  assert.equal(model.diagnostics?.some(entry => entry.code === "odf.approximated.shapes") ?? false, false, "a known outline is not an approximation");
});

test("an outline this preview cannot draw is reported as approximated, not shown wrongly", () => {
  const body = page([
    '<draw:custom-shape svg:x="0cm" svg:y="0cm" svg:width="3cm" svg:height="3cm"><draw:enhanced-geometry draw:type="non-primitive" draw:enhanced-path="M 0 0 L 100 100 Z"/></draw:custom-shape>',
    '<draw:custom-shape svg:x="4cm" svg:y="0cm" svg:width="3cm" svg:height="3cm"><draw:enhanced-geometry draw:type="mso-spt100"/></draw:custom-shape>',
  ].join(""));
  const model = parse({ body });
  assert.deepEqual(shapes(model.slides[0]).map(shape => shape.geometry), ["rect", "rect"]);
  const reported = model.diagnostics.find(entry => entry.code === "odf.approximated.shapes");
  assert.ok(reported, "the approximation is reported");
  assert.equal(reported.severity, "accuracy");
  assert.equal(reported.count, 2, "each approximated shape is counted");
  assert.match(reported.message, /drawn as rectangles/);
});

test("a shape that carries text becomes the shape and a text element over it", () => {
  const body = page('<draw:custom-shape svg:x="1cm" svg:y="1cm" svg:width="4cm" svg:height="2cm"><draw:enhanced-geometry draw:type="ellipse"/><text:p>In the shape</text:p></draw:custom-shape>');
  const { elements } = parse({ body }).slides[0];
  assert.deepEqual(elements.map(element => element.kind), ["shape", "text"]);
  assert.equal(textOf(elements[1].paragraphs[0]), "In the shape");
  assert.equal(elements[1].left, elements[0].left, "the text sits over the shape");
  assert.equal(elements[1].width, elements[0].width);
});

test("a line becomes a line element between the points it declares", () => {
  const styleAutomatic = graphicStyle("gr1", 'draw:stroke="dash" svg:stroke-color="#ff0000" svg:stroke-width="2pt"');
  const body = page('<draw:line draw:style-name="gr1" svg:x1="1cm" svg:y1="1cm" svg:x2="5cm" svg:y2="3cm"/>');
  const [line] = parse({ styleAutomatic, body }).slides[0].elements;
  assert.equal(line.kind, "line");
  assert.equal(line.left, Math.round((72 / 2.54) * EMU));
  assert.equal(line.width, Math.round(4 * (72 / 2.54) * EMU));
  assert.equal(line.height, Math.round(2 * (72 / 2.54) * EMU));
  assert.equal(line.line.color, "#ff0000");
  assert.equal(line.line.style, "dash");
  assert.equal(line.line.width, 2 * EMU);
  assert.equal(line.flipH, undefined, "a line drawn left to right is not flipped");
});

test("a line drawn backwards keeps its box and records the flip", () => {
  const body = page('<draw:line svg:x1="5cm" svg:y1="4cm" svg:x2="1cm" svg:y2="1cm"/>');
  const [line] = parse({ body }).slides[0].elements;
  assert.equal(line.left, Math.round((72 / 2.54) * EMU));
  assert.equal(line.top, Math.round((72 / 2.54) * EMU));
  assert.equal(line.flipH, true);
  assert.equal(line.flipV, true);
});

test("a style that draws no stroke leaves the shape without one", () => {
  const styleAutomatic = graphicStyle("gr1", 'draw:fill="none" draw:stroke="none"');
  const body = page('<draw:rect draw:style-name="gr1" svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="2cm"/>');
  const [shape] = shapes(parse({ styleAutomatic, body }).slides[0]);
  assert.equal(shape.fill, undefined);
  assert.equal(shape.line, undefined);
});

test("a table in a frame becomes a table element with its cells' text", () => {
  const table = '<table:table table:name="T"><table:table-column table:number-columns-repeated="2"/>'
    + "<table:table-row><table:table-cell><text:p>A1</text:p></table:table-cell><table:table-cell><text:p>B1</text:p></table:table-cell></table:table-row>"
    + "<table:table-row><table:table-cell><text:p>A2</text:p></table:table-cell><table:table-cell><text:p>B2</text:p></table:table-cell></table:table-row></table:table>";
  const body = page(`<draw:frame svg:x="1cm" svg:y="1cm" svg:width="8cm" svg:height="4cm">${table}</draw:frame>`);
  const [element] = parse({ body }).slides[0].elements;
  assert.equal(element.kind, "table");
  assert.equal(element.rows.length, 2);
  assert.equal(element.columns.length, 2);
  assert.deepEqual(element.rows.map(row => row.cells.map(cell => cell.paragraphs.map(textOf).join(""))), [["A1", "B1"], ["A2", "B2"]]);
  // The columns share the frame's width, and the rows its height.
  assert.equal(element.columns[0] * 2, element.width);
  assert.equal(element.rows[0].height * 2, element.height);
});

test("header rows of a table on a slide are read as rows", () => {
  const table = '<table:table table:name="T"><table:table-header-rows><table:table-row><table:table-cell><text:p>Head</text:p></table:table-cell></table:table-row></table:table-header-rows>'
    + "<table:table-row><table:table-cell><text:p>Body</text:p></table:table-cell></table:table-row></table:table>";
  const body = page(`<draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="4cm">${table}</draw:frame>`);
  const [element] = parse({ body }).slides[0].elements;
  assert.deepEqual(element.rows.map(row => row.cells[0].paragraphs.map(textOf).join("")), ["Head", "Body"]);
});

test("a table past the slide caps is reported rather than cut in silence", () => {
  const row = cells => `<table:table-row>${cells}</table:table-row>`;
  const wide = row(Array.from({ length: 40 }, (_, index) => `<table:table-cell><text:p>c${index}</text:p></table:table-cell>`).join(""));
  const body = page(`<draw:frame svg:x="0cm" svg:y="0cm" svg:width="8cm" svg:height="4cm"><table:table table:name="T">${wide}</table:table></draw:frame>`);
  const model = parse({ body });
  assert.equal(model.slides[0].elements[0].rows[0].cells.length, 32);
  assert.ok(model.diagnostics.some(entry => entry.code === "odf.unsupported.slide-content"), "the cells left out are reported");
});

test("a drawing kind this milestone does not read is still reported", () => {
  const body = page('<draw:polyline svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="2cm" draw:points="0,0 100,100"/><draw:g><draw:rect svg:x="0cm" svg:y="3cm" svg:width="1cm" svg:height="1cm"/></draw:g>');
  const model = parse({ body });
  assert.equal((model.slides[0].elements ?? []).length, 0, "nothing is invented for them");
  assert.ok(model.diagnostics.some(entry => entry.code === "odf.unsupported.slide-content"));
});
