import test from "node:test";
import assert from "node:assert/strict";
import { parseOdt } from "../dist/parsers/odf/odt.js";
import { makeOdt, ODT_MEDIA_TYPE } from "./helpers/odf.mjs";

// Milestone 3.4: images, headers, footers and page layout.

const parse = (options, prefixes) => parseOdt(makeOdt(options, prefixes), "letter.odt");
const text = paragraph => paragraph.runs.map(run => run.text).join("");
const paragraphs = model => model.blocks.filter(block => block.type === "paragraph").map(block => block.value);
const imageRuns = model => paragraphs(model).flatMap(paragraph => paragraph.runs.filter(run => run.image));

/** A one-pixel PNG, so the bytes a package carries are a real image. */
const PNG = Buffer.from("89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea735a99d0000000049454e44ae426082", "hex");
const frame = (href, extra = "") => `<draw:frame svg:width="2cm" svg:height="1cm"${extra}><draw:image xlink:href="${href}"/></draw:frame>`;
const withImage = (body, parts = { "Pictures/a.png": PNG }) => parse({ body, extra: parts });

test("an image stored in the package becomes an image run at the declared size", () => {
  const model = withImage(`<text:p>before ${frame("Pictures/a.png")} after</text:p>`);
  const [run] = imageRuns(model);
  assert.ok(run, "the image is shown");
  assert.equal(run.image.mimeType, "image/png");
  assert.deepEqual([...run.image.bytes.subarray(0, 4)], [0x89, 0x50, 0x4e, 0x47]);
  assert.ok(Math.abs(run.image.width - 56.69) < 0.1, `width was ${run.image.width}`);
  assert.ok(Math.abs(run.image.height - 28.35) < 0.1, `height was ${run.image.height}`);
  assert.equal(text(paragraphs(model)[0]), "before  after", "the surrounding text is kept");
});

test("an image description becomes the alternative text", () => {
  const body = `<text:p><draw:frame svg:width="2cm" svg:height="1cm"><draw:image xlink:href="Pictures/a.png"/><svg:desc>A red dot</svg:desc></draw:frame></text:p>`;
  assert.equal(imageRuns(withImage(body))[0].image.alt, "A red dot");
});

test("a reference that points outside the package is never followed", () => {
  for (const href of ["../secret.png", "/etc/passwd.png", "http://example.com/a.png", "file:///a.png", "Pictures/../../a.png", "Pictures/a.png#frag", "./../a.png"]) {
    const model = withImage(`<text:p>${frame(href)}</text:p>`);
    assert.equal(imageRuns(model).length, 0, href);
    assert.ok(model.diagnostics?.some(entry => entry.severity === "warning"), `${href}: the omission is reported`);
  }
});

test("an image format this library cannot read is reported, not shown", () => {
  const model = withImage(`<text:p>${frame("Pictures/a.wmf")}</text:p>`, { "Pictures/a.wmf": Buffer.from([1, 2, 3, 4]) });
  assert.equal(imageRuns(model).length, 0);
  assert.ok(model.diagnostics?.some(entry => entry.severity === "warning"));
});

test("a missing entry and an encrypted entry are both left out", () => {
  const missing = withImage(`<text:p>${frame("Pictures/none.png")}</text:p>`);
  assert.equal(imageRuns(missing).length, 0);
  const manifest = `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="${ODT_MEDIA_TYPE}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Pictures/a.png" manifest:media-type="image/png"><manifest:encryption-data/></manifest:file-entry></manifest:manifest>`;
  const encrypted = parse({ body: `<text:p>${frame("Pictures/a.png")}</text:p>`, extra: { "Pictures/a.png": PNG }, manifest });
  assert.equal(imageRuns(encrypted).length, 0, "an encrypted image is not decrypted or shown");
});

test("an embedded object or text box in a frame is reported rather than shown as nothing", () => {
  const body = '<text:p><draw:frame svg:width="2cm" svg:height="1cm"><draw:object xlink:href="./Object 1"/></draw:frame></text:p>';
  const model = withImage(body);
  assert.equal(imageRuns(model).length, 0);
  assert.ok(model.diagnostics?.some(entry => entry.severity === "warning"));
});

test("the page size and margins come from the master page's layout", () => {
  const model = parse({
    master: '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>',
    styles: '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-right="1.5cm" fo:margin-bottom="2cm" fo:margin-left="2.5cm"/></style:page-layout>',
    body: "<text:p>Body</text:p>",
  });
  assert.ok(Math.abs(model.page.width - 595.28) < 0.5, `width was ${model.page.width}`);
  assert.ok(Math.abs(model.page.height - 841.89) < 0.5, `height was ${model.page.height}`);
  assert.ok(Math.abs(model.page.marginTop - 56.69) < 0.1);
  assert.ok(Math.abs(model.page.marginLeft - 70.87) < 0.1);
  assert.ok(Math.abs(model.page.marginRight - 42.52) < 0.1);
});

test("a landscape layout without explicit sizes still gives a landscape page", () => {
  const model = parse({
    master: '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>',
    styles: '<style:page-layout style:name="pm1"><style:page-layout-properties style:print-orientation="landscape" fo:margin="1cm"/></style:page-layout>',
    body: "<text:p>Body</text:p>",
  });
  assert.ok(model.page.width > model.page.height, "landscape is wider than tall");
  assert.ok(Math.abs(model.page.marginTop - 28.35) < 0.1, "one margin value applies to every side");
});

test("a page size beyond any real surface is clamped", () => {
  const model = parse({
    master: '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>',
    styles: '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="100000cm" fo:page-height="1e9pt"/></style:page-layout>',
    body: "<text:p>Body</text:p>",
  });
  assert.ok(model.page.width <= 20_000 && model.page.height <= 20_000, `page was ${model.page.width} by ${model.page.height}`);
});

test("headers and footers become blocks of their own", () => {
  const model = parse({
    master: '<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Top of page</text:p></style:header><style:footer><text:p>Page <text:page-number>1</text:page-number></text:p></style:footer></style:master-page>',
    styles: '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout>',
    body: "<text:p>Body</text:p>",
  });
  assert.deepEqual(model.headers.map(block => text(block.value)), ["Top of page"]);
  assert.deepEqual(model.footers.map(block => text(block.value)), ["Page 1"]);
  assert.deepEqual(paragraphs(model).map(text), ["Body"], "surfaces are not part of the body");
  assert.equal(model.differentFirstPageHeaderFooter, undefined);
});

test("a first-page header is kept separately and marks the document as using one", () => {
  const model = parse({
    master: '<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Every page</text:p></style:header><style:header-first><text:p>First only</text:p></style:header-first></style:master-page>',
    styles: '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout>',
    body: "<text:p>Body</text:p>",
  });
  assert.deepEqual(model.headers.map(block => text(block.value)), ["Every page"]);
  assert.deepEqual(model.firstPageHeaders.map(block => text(block.value)), ["First only"]);
  assert.equal(model.differentFirstPageHeaderFooter, true);
});

test("header paragraphs keep their own formatting", () => {
  const model = parse({
    master: '<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p text:style-name="HP">Centred</text:p></style:header></style:master-page>',
    styles: '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout><style:style style:name="HP" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/><style:text-properties fo:font-weight="bold"/></style:style>',
    body: "<text:p>Body</text:p>",
  });
  assert.equal(model.headers[0].value.alignment, "center");
  assert.equal(model.headers[0].value.runs[0].bold, true);
});

test("a document with no master page has no page or surfaces", () => {
  const model = parse({ body: "<text:p>Body</text:p>" });
  assert.equal(model.page, undefined);
  assert.equal(model.headers, undefined);
  assert.equal(model.footers, undefined);
});

test("images, layout and surfaces work with renamed prefixes", () => {
  const renamed = parseOdt(makeOdt({
    master: '<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Head</text:p></style:header></style:master-page>',
    styles: '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout>',
    body: `<text:p>${frame("Pictures/a.png")}</text:p>`,
    extra: { "Pictures/a.png": PNG },
  }, { office: "o", text: "x", style: "s", fo: "f", draw: "d", svg: "g", xlink: "l", meta: "m", manifest: "mf" }), "letter.odt");
  assert.equal(imageRuns(renamed).length, 1);
  assert.equal(text(renamed.headers[0].value), "Head");
  assert.ok(renamed.page.width > 0);
});
