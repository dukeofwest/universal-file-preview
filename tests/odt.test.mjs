import test from "node:test";
import assert from "node:assert/strict";
import { parseOdt, odtWorkLimits, resolveOdtLimits } from "../dist/parsers/odf/odt.js";
import { makeOdt, OTT_MEDIA_TYPE } from "./helpers/odf.mjs";

// Milestone 3.2: the text body of an OpenDocument text document, read into the same WordDocument
// model the DOCX parser produces. Paragraphs, headings, spans and the formatting their styles carry.

const parse = (options, prefixes) => parseOdt(makeOdt(options, prefixes), "letter.odt");
const text = paragraph => paragraph.runs.map(run => run.text).join("");
const paragraphs = model => model.blocks.filter(block => block.type === "paragraph").map(block => block.value);

test("paragraphs and their text are read in document order", () => {
  const model = parse({ body: "<text:p>First</text:p><text:p>Second</text:p><text:p/>" });
  assert.equal(model.kind, "word");
  assert.equal(model.title, "letter");
  assert.equal(model.source.format, "odt");
  assert.equal(model.source.producer, "LibreOffice/24.8.4.2");
  assert.equal(model.source.formatVersion, "1.3");
  const found = paragraphs(model);
  assert.equal(found.length, 3);
  assert.deepEqual(found.map(text), ["First", "Second", ""]);
});

test("a template is recorded as its own format", () => {
  const model = parseOdt(makeOdt({ mimetype: OTT_MEDIA_TYPE, body: "<text:p>Letter</text:p>" }), "letter.ott");
  assert.equal(model.source.format, "ott");
});

test("headings carry their outline level, and a paragraph style can supply it", () => {
  const model = parse({
    styles: '<style:style style:name="Heading_20_3" style:family="paragraph" style:default-outline-level="3"/>',
    body: '<text:h text:outline-level="1">Title</text:h><text:h text:style-name="Heading_20_3">From the style</text:h><text:h>No level</text:h>',
  });
  const [first, second, third] = paragraphs(model);
  assert.equal(first.kind, "heading");
  assert.equal(first.headingLevel, 1);
  assert.equal(second.headingLevel, 3);
  assert.equal(third.kind, "heading");
  assert.equal(third.headingLevel, undefined, "a heading without a level is still a heading");
});

test("an outline level outside the range is ignored rather than trusted", () => {
  const model = parse({ body: '<text:h text:outline-level="0">A</text:h><text:h text:outline-level="99">B</text:h><text:h text:outline-level="NaN">C</text:h><text:h text:outline-level="2.7">D</text:h>' });
  const found = paragraphs(model);
  assert.equal(found[0].headingLevel, undefined);
  assert.equal(found[1].headingLevel, undefined);
  assert.equal(found[2].headingLevel, undefined);
  assert.equal(found[3].headingLevel, 2, "a fractional level is truncated");
});

test("character formatting comes from text styles, and spans nest", () => {
  const model = parse({
    automatic: [
      '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>',
      '<style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic" fo:color="#ff0000"/></style:style>',
      '<style:style style:name="T3" style:family="text"><style:text-properties style:text-underline-style="wave" style:text-line-through-style="solid" fo:font-size="14pt" style:font-name="Arial"/></style:style>',
    ].join(""),
    body: '<text:p>plain <text:span text:style-name="T1">bold <text:span text:style-name="T2">and italic</text:span></text:span><text:span text:style-name="T3">marked</text:span></text:p>',
  });
  const runs = paragraphs(model)[0].runs;
  assert.deepEqual(runs.map(run => run.text), ["plain ", "bold ", "and italic", "marked"]);
  assert.equal(runs[0].bold, undefined);
  assert.equal(runs[1].bold, true);
  assert.equal(runs[2].bold, true, "the enclosing span's formatting is carried in");
  assert.equal(runs[2].italic, true);
  assert.equal(runs[2].color, "#ff0000");
  assert.equal(runs[3].underline, true);
  assert.equal(runs[3].underlineStyle, "wavy");
  assert.equal(runs[3].strike, true);
  assert.equal(runs[3].fontSize, 14);
  assert.equal(runs[3].fontFamily, "Arial");
});

test("a paragraph style's own text properties apply to its runs, and inherit through parents", () => {
  const model = parse({
    styles: '<style:style style:name="Base" style:family="paragraph"><style:text-properties fo:font-weight="bold" fo:color="#112233"/><style:paragraph-properties fo:text-align="center"/></style:style>',
    automatic: '<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Base"><style:text-properties fo:color="#445566"/></style:style>',
    body: '<text:p text:style-name="P1">Inherited</text:p>',
  });
  const paragraph = paragraphs(model)[0];
  assert.equal(paragraph.alignment, "center", "the parent's paragraph properties apply");
  assert.equal(paragraph.style, "P1");
  assert.equal(paragraph.runs[0].bold, true, "inherited from the parent style");
  assert.equal(paragraph.runs[0].color, "#445566", "the nearer declaration wins");
});

test("paragraph formatting is read from the paragraph properties", () => {
  const model = parse({
    automatic: '<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="end" fo:margin-left="1cm" fo:margin-right="0.5cm" fo:text-indent="0.25cm" fo:margin-top="6pt" fo:margin-bottom="3pt" fo:background-color="#eeeeee" fo:break-before="page" fo:line-height="150%"/></style:style>',
    body: '<text:p text:style-name="P1">Formatted</text:p>',
  });
  const paragraph = paragraphs(model)[0];
  assert.equal(paragraph.alignment, "right");
  assert.ok(Math.abs(paragraph.indentLeft - 28.35) < 0.1, `indentLeft was ${paragraph.indentLeft}`);
  assert.ok(Math.abs(paragraph.spacingBefore - 6) < 0.001);
  assert.ok(Math.abs(paragraph.spacingAfter - 3) < 0.001);
  assert.equal(paragraph.background, "#eeeeee");
  assert.equal(paragraph.pageBreakBefore, true);
  assert.equal(paragraph.lineHeight, 1.5);
  assert.equal(paragraph.lineHeightRule, "auto");
});

test("spaces, tabs and line breaks become the text they stand for", () => {
  const model = parse({ body: '<text:p>a<text:s text:c="3"/>b<text:tab/>c<text:line-break/>d</text:p>' });
  const runs = paragraphs(model)[0].runs;
  assert.equal(runs.map(run => run.text).join(""), "a   b\tc\nd");
  // A line break is a newline in the text, as the DOCX parser emits for a plain w:br. It must not be
  // the model's pageBreak run: the renderer splits pages at those, so a line break became a new page.
  assert.equal(runs.some(run => run.pageBreak), false, "a line break is not carried as a page break");
  // A repeat count is clamped, never trusted: it decides how long a string becomes.
  const huge = parse({ body: '<text:p>a<text:s text:c="99999999"/>b</text:p>' });
  assert.ok(text(paragraphs(huge)[0]).length <= 10_002, "the space count is clamped");
  const broken = parse({ body: '<text:p>a<text:s text:c="NaN"/>b</text:p>' });
  assert.equal(text(paragraphs(broken)[0]), "a b");
});

test("the text of a hyperlink and of a field is kept as the producer saved it", () => {
  const model = parse({ body: '<text:p>See <text:a xlink:href="https://example.com/x">the site</text:a> on <text:date text:date-value="2026-01-01">1 January 2026</text:date></text:p>' });
  assert.equal(text(paragraphs(model)[0]), "See the site on 1 January 2026");
});

test("content that cannot be shown is reported rather than dropped silently", () => {
  const model = parse({ body: '<text:p>Before</text:p><text:p>With <draw:frame><draw:image xlink:href="Pictures/a.png"/></draw:frame> a picture</text:p><text:p>After <text:note text:note-class="footnote"><text:note-body><text:p>Note</text:p></text:note-body></text:note></text:p>' });
  assert.deepEqual(paragraphs(model).map(text), ["Before", "With  a picture", "After "]);
  assert.ok(model.diagnostics?.length, "a diagnostic is reported");
  assert.ok(model.diagnostics.every(entry => typeof entry.message === "string" && entry.message.length > 0));
  assert.ok(model.diagnostics.some(entry => entry.severity === "warning"), "it must be visible, not info-only");
});

test("the document default style becomes the model defaults", () => {
  const model = parse({
    styles: '<style:default-style style:family="paragraph"><style:text-properties fo:font-size="12pt" style:font-name="Liberation Serif"/><style:paragraph-properties fo:text-align="justify"/></style:default-style>',
    body: "<text:p>Body</text:p>",
  });
  assert.equal(model.defaultRun.fontSize, 12);
  assert.equal(model.defaultRun.fontFamily, "Liberation Serif");
  assert.equal(model.defaultParagraph.alignment, "justify");
});

test("a style inheritance cycle is reported and does not hang", () => {
  const model = parse({
    automatic: '<style:style style:name="A" style:family="paragraph" style:parent-style-name="B"/><style:style style:name="B" style:family="paragraph" style:parent-style-name="A"/>',
    body: '<text:p text:style-name="A">Text</text:p>',
  });
  assert.equal(text(paragraphs(model)[0]), "Text");
  assert.ok(model.diagnostics?.some(entry => /cycle/i.test(entry.message)));
});

test("any namespace prefix works, as producers may choose their own", () => {
  const renamed = parse({ body: "<text:h text:outline-level=\"1\">Title</text:h><text:p>Body</text:p>" }, { office: "o", text: "x", style: "s", fo: "f", meta: "m" });
  assert.deepEqual(paragraphs(renamed).map(text), ["Title", "Body"]);
  assert.equal(paragraphs(renamed)[0].headingLevel, 1);
});

test("the work limits follow the block budget, never a byte size", () => {
  assert.deepEqual(odtWorkLimits({ maxBlocks: 100_000 }), { maxElements: 2_650_000, maxTokens: 10_600_000 });
  assert.equal(odtWorkLimits({ maxBlocks: 1 }).maxElements, 250_024);
  assert.equal(odtWorkLimits({ maxBlocks: Number.MAX_SAFE_INTEGER }).maxElements, 2 ** 32);
  assert.throws(() => resolveOdtLimits({ maxBlocks: 0 }), /positive safe integer/);
  assert.throws(() => resolveOdtLimits({ nonsense: 1 }), /Unknown ODT limit/);
});

test("a document with more blocks than the limit is refused", () => {
  const body = "<text:p>x</text:p>".repeat(40);
  assert.throws(() => parseOdt(makeOdt({ body }), "big.odt", { limits: { maxBlocks: 10 } }), error => error?.name === "PreviewError" && error.code === "resource-limit");
});

test("a package that is not a text document is refused before anything is read", () => {
  assert.throws(() => parseOdt(makeOdt({ mimetype: "application/vnd.oasis.opendocument.spreadsheet", body: "<text:p>x</text:p>" }), "x.odt"), error => error?.code === "unsupported-format");
});
