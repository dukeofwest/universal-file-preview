import test from "node:test";
import assert from "node:assert/strict";
import { parseFodt } from "../dist/parsers/odf/odt-flat.js";
import { parseOdt } from "../dist/parsers/odf/odt.js";
import { parseModel } from "../dist/parsers/parse-model.js";
import { createDefaultRendererRegistry } from "../dist/renderers/registry.js";
import { WORKER_FORMATS, expectedKind } from "../dist/worker/protocol.js";
import { makeFodt, makeOdt, OTT_MEDIA_TYPE } from "./helpers/odf.mjs";

// Milestone 3.5: flat documents, template behaviour and the worker and registry wiring.

const parse = (options, prefixes) => parseFodt(makeFodt(options, prefixes), "letter.fodt");
const text = paragraph => paragraph.runs.map(run => run.text).join("");
const paragraphs = model => model.blocks.filter(block => block.type === "paragraph").map(block => block.value);
const PNG_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP4//8/AAX+Av6nNamdAAAAAElFTkSuQmCC";

test("a flat document is read into the same model as a package", () => {
  const body = '<text:h text:outline-level="1">Title</text:h><text:p>Body <text:span text:style-name="T1">bold</text:span></text:p>';
  const automatic = '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>';
  const flat = parse({ body, automatic });
  assert.equal(flat.kind, "word");
  assert.equal(flat.source.format, "fodt");
  assert.equal(flat.source.producer, "LibreOffice/24.8.4.2");
  assert.deepEqual(paragraphs(flat).map(text), ["Title", "Body bold"]);
  assert.equal(paragraphs(flat)[0].headingLevel, 1);
  // The same content in a package gives the same blocks, so one reader really does serve both forms.
  const packaged = parseOdt(makeOdt({ body, automatic }), "letter.odt");
  assert.deepEqual(paragraphs(packaged).map(text), paragraphs(flat).map(text));
  assert.deepEqual(paragraphs(packaged)[1].runs.map(run => run.bold), paragraphs(flat)[1].runs.map(run => run.bold));
});

test("a flat document reads its common styles, lists and tables", () => {
  const styles = '<style:style style:name="Base" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style><text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style>';
  const body = '<text:p text:style-name="Base">Centred</text:p><text:list text:style-name="L1"><text:list-item><text:p>One</text:p></text:list-item></text:list><table:table table:name="T"><table:table-row><table:table-cell><text:p>Cell</text:p></table:table-cell></table:table-row></table:table>';
  const model = parse({ styles, body });
  assert.equal(paragraphs(model)[0].alignment, "center");
  assert.equal(paragraphs(model)[1].list.label, "1.");
  assert.equal(model.blocks.filter(block => block.type === "table").length, 1);
});

test("a flat document takes its page and surfaces from its own master styles", () => {
  const model = parse({
    master: '<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Head</text:p></style:header></style:master-page>',
    styles: '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin="2cm"/></style:page-layout>',
    body: "<text:p>Body</text:p>",
  });
  assert.ok(Math.abs(model.page.width - 595.28) < 0.5);
  assert.ok(Math.abs(model.page.marginTop - 56.69) < 0.1);
  assert.deepEqual(model.headers.map(block => text(block.value)), ["Head"]);
});

test("an image stored inline is shown, and its type comes from the bytes", () => {
  const body = `<text:p><draw:frame svg:width="1cm" svg:height="1cm"><draw:image><office:binary-data>${PNG_BASE64}</office:binary-data></draw:image></draw:frame></text:p>`;
  const runs = paragraphs(parse({ body })).flatMap(paragraph => paragraph.runs.filter(run => run.image));
  assert.equal(runs.length, 1);
  assert.equal(runs[0].image.mimeType, "image/png");
  assert.deepEqual([...runs[0].image.bytes.subarray(0, 4)], [0x89, 0x50, 0x4e, 0x47]);
});

test("inline data that is not a readable image is reported rather than shown", () => {
  const body = '<text:p><draw:frame svg:width="1cm" svg:height="1cm"><draw:image><office:binary-data>bm90IGFuIGltYWdl</office:binary-data></draw:image></draw:frame></text:p>';
  const model = parse({ body });
  assert.equal(paragraphs(model).flatMap(paragraph => paragraph.runs.filter(run => run.image)).length, 0);
  assert.ok(model.diagnostics?.some(entry => entry.severity === "warning"));
});

test("a flat document of another kind is refused", () => {
  assert.throws(() => parse({ body: "<text:p>x</text:p>", mimetype: "application/vnd.oasis.opendocument.spreadsheet" }), error => error?.code === "unsupported-format");
  assert.throws(() => parse({ body: "<text:p>x</text:p>", mimetype: "application/vnd.oasis.opendocument.text-master" }), error => error?.code === "unsupported-format");
});

test("a document larger than the limit is refused before it is decoded", () => {
  const bytes = makeFodt({ body: "<text:p>x</text:p>" });
  assert.throws(() => parseFodt(bytes, "letter.fodt", { limits: { maxContentBytes: 10 } }), error => error?.code === "resource-limit");
});

test("a template says so, in both forms, and is shown as saved", () => {
  const packaged = parseOdt(makeOdt({ mimetype: OTT_MEDIA_TYPE, body: "<text:p>Dear &lt;name&gt;</text:p>" }), "letter.ott");
  assert.equal(packaged.source.format, "ott");
  assert.ok(packaged.diagnostics?.some(entry => /template/i.test(entry.message) && entry.severity === "warning"), "a packaged template is labelled");
  assert.equal(text(paragraphs(packaged)[0]), "Dear <name>", "placeholders are shown as saved, never filled in");
  const flat = parse({ body: "<text:p>Dear</text:p>", mimetype: OTT_MEDIA_TYPE });
  assert.equal(flat.source.format, "fodt");
  assert.ok(flat.diagnostics?.some(entry => /template/i.test(entry.message)), "a flat template is labelled too");
  const ordinary = parse({ body: "<text:p>Dear</text:p>" });
  assert.equal(ordinary.diagnostics?.some(entry => /template/i.test(entry.message)) ?? false, false);
});

test("every text format is wired through the parser, the worker and the registry", () => {
  const registry = createDefaultRendererRegistry();
  for (const format of ["odt", "ott", "fodt"]) {
    assert.ok(WORKER_FORMATS.includes(format), `${format} is a worker format`);
    assert.equal(expectedKind(format), "word", `${format} produces a word model`);
    assert.ok(registry.get(format), `${format} has a renderer`);
  }
  assert.equal(registry.get("odt"), registry.get("fodt"), "one renderer instance serves them all");
  // Every worker format must have a kind: a missing entry would only surface as a validation error.
  for (const format of WORKER_FORMATS) assert.ok(["word", "workbook", "presentation"].includes(expectedKind(format)), format);
});

test("parseModel reads each text format and keeps the format it came from", () => {
  const body = "<text:p>Body</text:p>";
  assert.equal(parseModel("odt", makeOdt({ body }), "a.odt", { locale: "en-US" }).source.format, "odt");
  assert.equal(parseModel("ott", makeOdt({ body, mimetype: OTT_MEDIA_TYPE }), "a.ott", { locale: "en-US" }).source.format, "ott");
  assert.equal(parseModel("fodt", makeFodt({ body }), "a.fodt", { locale: "en-US" }).source.format, "fodt");
});

test("a flat document works with renamed prefixes", () => {
  const renamed = parse({ body: '<text:h text:outline-level="2">T</text:h><text:p>B</text:p>' }, { office: "o", text: "x", style: "s", fo: "f", meta: "m" });
  assert.deepEqual(paragraphs(renamed).map(text), ["T", "B"]);
  assert.equal(paragraphs(renamed)[0].headingLevel, 2);
});
