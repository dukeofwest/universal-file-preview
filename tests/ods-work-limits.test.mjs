import test from "node:test";
import assert from "node:assert/strict";
import { parseOds, xmlWorkLimits } from "../dist/parsers/odf/ods.js";
import { parseFods } from "../dist/parsers/odf/flat.js";
import { sanitizeSvg } from "../dist/renderers/image-url.js";
import { NamespaceScope, XML_NAMESPACE } from "../dist/xml/xml-namespaces.js";
import { makeZip } from "./helpers/zip.mjs";

// XML work limits for documents parsed through the namespace-aware layer.
//
// 1. The element and token budgets of an OpenDocument content stream follow the cell budget and never
//    a byte size. They were derived from maxContentBytes, which put them at 33.5 million elements
//    whatever a file held, so a 114 KiB package of empty elements was tokenized for about 11 s.
//    A compression-ratio limit is not used, and the tests below pin why: a legitimate workbook
//    compresses past 100:1, and padding an attack with incompressible bytes puts it under any ratio
//    that keeps such workbooks working.
// 2. Entering an element that declares a namespace copies every binding in scope, and nothing
//    bounded that work. Nesting (256) and attributes per element (512) allow over 131,000 bindings in
//    scope, so a 560 KiB spreadsheet of 433 elements cost 17 s and 1.1 GiB, and a 3.1 MiB SVG of 431
//    elements cost 8.3 s and 1.3 GiB, well inside every element and token limit.

const KiB = 1024;
const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';
const ODS_MEDIA_TYPE = "application/vnd.oasis.opendocument.spreadsheet";
const content = (body, automatic = "") => `<office:document-content ${NS} office:version="1.3">${automatic ? `<office:automatic-styles>${automatic}</office:automatic-styles>` : ""}<office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document-content>`;

/** A minimal ODS package: a stored mimetype entry and content.xml, deflated unless `deflate` is false. */
function odsPackage(body, { deflate = true, automatic = "", styles } = {}) {
  const entries = { mimetype: ODS_MEDIA_TYPE, "content.xml": content(body, automatic), ...(styles ? { "styles.xml": styles } : {}) };
  return new Uint8Array(makeZip(entries, { deflate, stored: ["mimetype"] }));
}
const flatOds = body => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?><office:document ${NS} office:version="1.3" office:mimetype="${ODS_MEDIA_TYPE}"><office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document>`);
/** `count` float cells laid out 10 per row: the shape of a plain workbook, two elements per cell. */
function cells(count, perRow = 10) {
  const row = `<table:table-row>${Array.from({ length: perRow }, (_, index) => `<table:table-cell office:value-type="float" office:value="${index + 0.5}"><text:p>${index + 0.5}</text:p></table:table-cell>`).join("")}</table:table-row>`;
  return `<table:table table:name="S">${row.repeat(Math.ceil(count / perRow))}</table:table>`;
}
/** A stack of `levels` elements each declaring `perLevel` namespaces, then `leaves` that declare one more. */
function declarationStack(levels, perLevel, leaves, leaf) {
  let open = "", close = "";
  for (let level = 0; level < levels; level++) {
    const declarations = Array.from({ length: perLevel }, (_, index) => `xmlns:p${level}_${index}="urn:${level}:${index}"`).join(" ");
    open += `<d${level} ${declarations}>`; close = `</d${level}>` + close;
  }
  return open + leaf.repeat(leaves) + close;
}
/**
 * `length` characters that deflate cannot compress, from a fixed seed so the fixture is reproducible.
 * The alphabet avoids the characters XML escapes, so the text can sit in an attribute value.
 */
function incompressibleText(length, seed = 0x9e3779b9) {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  let state = seed >>> 0 || 1;
  const out = new Array(length);
  for (let index = 0; index < length; index++) {
    state ^= state << 13; state >>>= 0;
    state ^= state >>> 17;
    state ^= state << 5; state >>>= 0;
    out[index] = alphabet[state & 63];
  }
  return out.join("");
}
const thrown = work => { try { work(); return undefined; } catch (error) { return error; } };
const isRefusal = (error, code = "resource-limit") => error?.name === "PreviewError" && error.code === code;

// 5,000,000 elements is above the default budget (250,000 + 8 per cell, so 4,250,000 at 500,000 cells).
// The 114 MiB original of this attack and a padded variant are in the denial-of-service corpus.
const BOMB_ELEMENTS = 5_000_000;

test("the element budget is a function of the cell budget alone", () => {
  assert.deepEqual(xmlWorkLimits({ maxCells: 500_000 }), { maxElements: 4_250_000, maxTokens: 17_000_000 });
  assert.deepEqual(xmlWorkLimits({ maxCells: 1_000_000 }), { maxElements: 8_250_000, maxTokens: 33_000_000 });
  // Eight elements per cell is the contract: a 500,000-cell workbook with three nested spans in every
  // cell holds about 2.53 million elements and must be accepted, so the margin may not be narrowed.
  assert.ok(xmlWorkLimits({ maxCells: 500_000 }).maxElements > 2_525_006);
  // The token budget must not bind first: an indented producer costs four tokens per element, so the
  // element count it permits has to stay above that same workbook.
  assert.ok(xmlWorkLimits({ maxCells: 500_000 }).maxTokens / 4 > 2_525_006);
  // Lowering the limit stays monotonic and keeps the floor for styles and scaffolding.
  assert.equal(xmlWorkLimits({ maxCells: 1 }).maxElements, 250_008);
  // Clamped to a safe integer whatever a caller asks for.
  assert.equal(xmlWorkLimits({ maxCells: Number.MAX_SAFE_INTEGER }).maxElements, 2 ** 32);
});

test("a small package of empty elements is refused by the element budget", () => {
  const bytes = odsPackage("<x/>".repeat(BOMB_ELEMENTS));
  assert.ok(bytes.length < 64 * KiB, `package is ${(bytes.length / KiB).toFixed(0)} KiB`);
  const error = thrown(() => parseOds(bytes, "bomb.ods"));
  assert.ok(isRefusal(error), `expected resource-limit, got ${error?.code ?? "no error"}`);
  assert.match(error.message, /element count/i);
});

test("an element bomb padded below the OOXML compression ratio is still refused", () => {
  // Incompressible filler in an attribute value: a ratio limit alone would let this through.
  const padding = incompressibleText(600 * KiB);
  const body = `<table:table table:name="S" table:style-name="${padding}"/>${"<x/>".repeat(BOMB_ELEMENTS)}`;
  const bytes = odsPackage(body), ratio = content(body).length / bytes.length;
  assert.ok(ratio < 100, `ratio is ${ratio.toFixed(0)}:1, so a 100:1 limit would already refuse this file`);
  assert.ok(isRefusal(thrown(() => parseOds(bytes, "padded.ods"))));
});

test("an uncompressed element bomb is refused on the same budget", () => {
  assert.ok(isRefusal(thrown(() => parseOds(odsPackage("<x/>".repeat(BOMB_ELEMENTS), { deflate: false }), "stored.ods"))));
});

test("a flat document of empty elements is refused on the same budget", () => {
  assert.ok(isRefusal(thrown(() => parseFods(flatOds("<x/>".repeat(BOMB_ELEMENTS)), "bomb.fods"))));
});

test("elements that materialise no cells are charged too", () => {
  // Styled empty cells outside the drawn grid, and a skipped pivot-table subtree: both are discarded,
  // and both cost the tokenizing that the budget bounds.
  const styled = `<table:table table:name="S">${`<table:table-row><table:table-cell table:style-name="ce1" table:number-columns-repeated="1"/></table:table-row>`.repeat(BOMB_ELEMENTS / 2)}</table:table>`;
  assert.ok(isRefusal(thrown(() => parseOds(odsPackage(styled, { automatic: '<style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:background-color="#ff0000"/></style:style>' }), "styled.ods"))));
  const skipped = `<table:data-pilot-tables>${"<x/>".repeat(BOMB_ELEMENTS)}</table:data-pilot-tables>${cells(10)}`;
  assert.ok(isRefusal(thrown(() => parseOds(odsPackage(skipped), "pivot.ods"))));
});

test("the other OpenDocument parts keep the default element limit", () => {
  // styles.xml, the manifest and meta.xml are parsed as trees with the 250,000-element default.
  const styles = `<office:document-styles ${NS} office:version="1.3">${"<x/>".repeat(1_000_000)}</office:document-styles>`;
  assert.ok(isRefusal(thrown(() => parseOds(odsPackage(cells(10), { styles }), "styles.ods"))));
});

test("a workbook within the cell budget is unaffected", () => {
  assert.equal(parseOds(odsPackage(cells(50_000)), "real.ods").sheets[0].cells.length, 50_000);
});

test("a legitimate workbook compresses past the OOXML ratio limit", () => {
  // Why no ratio limit for OpenDocument parts: this file is ordinary content, and refusing it would be wrong.
  const body = cells(200_000);
  const bytes = odsPackage(body), ratio = content(body).length / bytes.length;
  assert.ok(ratio > 100, `ratio is ${ratio.toFixed(0)}:1`);
  assert.equal(parseOds(bytes, "real.ods").sheets[0].cells.length, 200_000);
});

test("entering an element costs only its own namespace declarations", () => {
  // 115,000 bindings in scope under 2,000 declaring elements: copying the bindings in scope for each
  // of them cost 17 s and 1.1 GiB. Chained scopes read it without copying, so this is now ordinary
  // content and is parsed, not refused. The denial-of-service corpus holds the time and heap bounds.
  const body = `${cells(10)}${declarationStack(230, 500, 2_000, '<leaf xmlns:q="urn:q"/>')}`;
  assert.equal(parseOds(odsPackage(body), "scope.ods").sheets[0].cells.length, 10);
  // The SVG sanitizer reads images through the same layer, so it is covered by the same change.
  // The wrappers are not on the allowlist, so the sanitized form is nearly empty; what matters is that
  // the document is read and reduced rather than costing seconds or being refused.
  const svg = new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg">${declarationStack(230, 500, 2_000, '<rect xmlns:q="urn:q" width="1" height="1" fill="#000"/>')}</svg>`);
  assert.match(new TextDecoder().decode(sanitizeSvg(svg)), /^<svg[^>]*\/?>/, "the SVG is sanitized rather than refused");
});

test("a prefix declared far above an element still resolves, and undeclaring still works", () => {
  // Chained lookups must give the same answers as a flattened map, including rebinding and xmlns="".
  const scope = new NamespaceScope();
  scope.enter([["xmlns:p", "urn:a"], ["xmlns", "urn:default"]]);
  scope.enter([["xmlns:p", "urn:b"]]);
  assert.equal(scope.lookup("p"), "urn:b", "the nearest declaration wins");
  assert.equal(scope.element("x").uri, "urn:default", "an inherited default namespace resolves");
  scope.enter([["xmlns", ""]]);
  assert.equal(scope.element("x").uri, "", "xmlns=\"\" puts unprefixed names in no namespace");
  assert.equal(scope.lookup("p"), "urn:b", "undeclaring the default leaves prefixes alone");
  scope.exit();
  assert.equal(scope.element("x").uri, "urn:default", "leaving the element restores the default");
  assert.deepEqual([...scope.snapshot()].sort(), [["p", "urn:b"], ["xml", XML_NAMESPACE], ["", "urn:default"]].sort());
});

test("documents that declare namespaces normally are unaffected", () => {
  // Every element carrying its own declaration, as some producers write: cheap while few are in scope.
  const row = `<table:table-row xmlns:q="urn:q"><table:table-cell xmlns:r="urn:r" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>`;
  assert.equal(parseOds(odsPackage(`<table:table table:name="S">${row.repeat(20_000)}</table:table>`), "many.ods").sheets[0].cells.length, 20_000);
  const svg = new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">${'<rect xmlns:q="urn:q" width="1" height="1" fill="#000"/>'.repeat(20_000)}</svg>`);
  assert.ok(sanitizeSvg(svg).length > 1000, "an ordinary SVG with per-element declarations is still sanitized");
});
