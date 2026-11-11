// Denial-of-service corpus (security audit, acceptance criterion 8): small crafted files, one per resource limit.
// Each entry is refused with a PreviewError ("refused") or handled within the library's work limits ("handled").
// The audit's budget is 2 seconds and a 256 MiB heap; an entry that needs more says so in `budget`, with the reason.
// Test-only; Node built-ins and the test helpers only. Used by tests/dos-budget.test.mjs.
import { makeZip } from "./zip.mjs";
import { makeFods, makeFodt, makeOds, makeOdt } from "./odf.mjs";

const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const imageRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
/** A one-pixel PNG, so an image fixture carries real bytes. */
const ONE_PIXEL_PNG = Buffer.from("89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea735a99d0000000049454e44ae426082", "hex");
const column = index => String.fromCharCode(65 + Math.floor(index / 26)) + String.fromCharCode(65 + index % 26);

/** Text deflate cannot compress, from a fixed seed so every fixture stays reproducible. */
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
/** A stack of `levels` elements each declaring `perLevel` namespaces, then `leaves` that declare one more. */
function declarationStack(levels, perLevel, leaves, leaf) {
  let open = "", close = "";
  for (let level = 0; level < levels; level++) {
    open += `<d${level} ${Array.from({ length: perLevel }, (_, index) => `xmlns:p${level}_${index}="urn:${level}:${index}"`).join(" ")}>`;
    close = `</d${level}>` + close;
  }
  return open + leaf.repeat(leaves) + close;
}

/** An XLSX workbook: `sheets` worksheet bodies (the part inside <worksheet>), optional styles and shared strings. */
function xlsx(sheets, { styles, shared } = {}) {
  const entries = {
    "xl/workbook.xml": `<workbook ${S}><sheets>${sheets.map((_, index) => `<sheet name="S${index}" sheetId="${index + 1}" r:id="s${index}"/>`).join("")}</sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships>${sheets.map((_, index) => `<Relationship Id="s${index}" Type="x/worksheet" Target="worksheets/sheet${index}.xml"/>`).join("")}${styles ? '<Relationship Id="st" Type="x/styles" Target="styles.xml"/>' : ""}${shared ? '<Relationship Id="ss" Type="x/sharedStrings" Target="sharedStrings.xml"/>' : ""}</Relationships>`,
    ...(styles ? { "xl/styles.xml": styles } : {}),
    ...(shared ? { "xl/sharedStrings.xml": `<sst ${S}>${shared.map(text => `<si><t>${text}</t></si>`).join("")}</sst>` } : {}),
  };
  sheets.forEach((body, index) => { entries[`xl/worksheets/sheet${index}.xml`] = `<worksheet ${S}>${body}</worksheet>`; });
  return new Uint8Array(makeZip(entries, { deflate: true }));
}
/** `count` cells laid out 100 per row, each produced by `cell(ref)`. */
function rows(count, cell, perRow = 100) {
  const out = [];
  for (let start = 0; start < count; start += perRow) out.push(`<row r="${start / perRow + 1}">${Array.from({ length: Math.min(perRow, count - start) }, (_, index) => cell(`${column(index)}${start / perRow + 1}`)).join("")}</row>`);
  return `<sheetData>${out.join("")}</sheetData>`;
}
const repeatedSheet = rows(120_000, ref => `<c r="${ref}"><v>1</v></c>`, 120);

export function dosCorpus() {
  return [
    // Packages and decompression.
    { id: "xml-part-expands-over-100x.docx", kind: "docx", expect: "refused", limit: "100:1 expansion for XML parts over 1 MiB",
      build: () => new Uint8Array(makeZip({ "word/document.xml": `<w:document ${W}><w:body><w:p><w:r><w:t>${"a".repeat(8 * 1024 * 1024)}</w:t></w:r></w:p></w:body></w:document>` }, { deflate: true })) },
    { id: "one-image-referenced-1000-times.docx", kind: "docx", expect: "handled", limit: "each part decompressed once",
      build: () => {
        const png = new Uint8Array(4 * 1024 * 1024); png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
        const runs = Array.from({ length: 1000 }, (_, index) => `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="i${index}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`).join("");
        return new Uint8Array(makeZip({
          "word/document.xml": `<w:document ${W}><w:body>${runs}</w:body></w:document>`,
          "word/_rels/document.xml.rels": `<Relationships>${Array.from({ length: 1000 }, (_, index) => `<Relationship Id="i${index}" Type="${imageRel}" Target="media/a.png"/>`).join("")}</Relationships>`,
          "word/media/a.png": png,
        }, { deflate: true }));
      } },
    { id: "one-slide-listed-10000-times.pptx", kind: "pptx", expect: "handled", limit: "each slide part parsed once",
      build: () => new Uint8Array(makeZip({
        "ppt/presentation.xml": `<p:presentation ${P}><p:sldIdLst>${'<p:sldId id="256" r:id="rId1"/>'.repeat(10_000)}</p:sldIdLst></p:presentation>`,
        "ppt/_rels/presentation.xml.rels": '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>',
        "ppt/slides/slide1.xml": `<p:sld ${P}><p:cSld><p:spTree/></p:cSld></p:sld>`,
      }, { deflate: true })) },
    // XLSX cells, text and work limits.
    { id: "600000-cells.xlsx", kind: "xlsx", expect: "refused", limit: "500,000 cells per workbook",
      // Refused at the 500,001st cell: up to then it costs what the largest accepted workbook costs (Addendum A4 budget).
      budget: { ms: 8000, heapMiB: 768, reason: "a workbook at the accepted maximum size (500,000 cells)" },
      build: () => xlsx(Array.from({ length: 5 }, () => repeatedSheet)) },
    { id: "shared-string-referenced-100000-times.xlsx", kind: "xlsx", expect: "refused", limit: "32,767 characters per cell (the 0.2.1 text amplification)",
      build: () => xlsx([rows(100_000, ref => `<c r="${ref}" t="s"><v>0</v></c>`)], { shared: ["y".repeat(1_000_000)] }) },
    { id: "number-format-text-amplification.xlsx", kind: "xlsx", expect: "refused", limit: "text of the whole model (64 Mi code units)",
      build: () => xlsx([rows(5000, ref => `<c r="${ref}" s="1"><v>1</v></c>`)], { styles: `<styleSheet ${S}><numFmts count="1"><numFmt numFmtId="200" formatCode="&quot;${"x".repeat(30_000)}&quot;0"/></numFmts><cellXfs count="2"><xf/><xf numFmtId="200"/></cellXfs></styleSheet>` }) },
    { id: "1500-sheets.xlsx", kind: "xlsx", expect: "refused", limit: "1,024 sheets",
      build: () => xlsx(Array.from({ length: 1500 }, () => "<sheetData/>")) },
    { id: "deep-formula-chain.xlsx", kind: "xlsx", expect: "handled", limit: "formula work and nesting limits",
      // About 2.3 s here: the formula work budget (kept at 20,000,000 units, 500,000 evaluations; reviewed for 0.3.0)
      // trades a few seconds in the worst case for calculating large generated workbooks.
      budget: { ms: 5000, heapMiB: 256, reason: "the formula work budget (20,000,000 units) allows about 2 s of evaluation" },
      build: () => xlsx([rows(50_000, ref => { const row = Number(ref.slice(2)), col = ref.slice(0, 2); return `<c r="${ref}"><f>${col}${row + 1}+1</f></c>`; })]) },
    { id: "whole-sheet-hyperlinks-and-conditional-formats.xlsx", kind: "xlsx", expect: "handled", limit: "hyperlink and conditional-format cell budgets",
      build: () => xlsx([`<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><conditionalFormatting sqref="A1:XFD1048576"><cfRule type="expression" dxfId="0" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting><hyperlinks>${'<hyperlink ref="A1:XFD1048576" r:id="h1"/>'.repeat(5)}</hyperlinks>`]) },
    { id: "20000-column-ranges-5000-merges.xlsx", kind: "xlsx", expect: "handled", limit: "XML element limits; work bounded by the renderer",
      build: () => xlsx([`<cols>${'<col min="1" max="16384" width="9"/>'.repeat(20_000)}</cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><mergeCells count="5000">${'<mergeCell ref="A1:SF5001"/>'.repeat(5000)}</mergeCells>`]) },
    // OpenDocument.
    { id: "repeat-bomb.ods", kind: "ods", expect: "refused", limit: "500,000 materialised cells",
      build: () => makeOds({ body: '<table:table table:name="S"><table:table-row table:number-rows-repeated="1048576"><table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="16384"/></table:table-row></table:table>' }) },
    { id: "repeated-long-text.ods", kind: "ods", expect: "refused", limit: "text of the whole model, counted per copy",
      build: () => makeOds({ body: `<table:table table:name="S"><table:table-row><table:table-cell office:value-type="string" table:number-columns-repeated="3000"><text:p>${"t".repeat(32_767)}</text:p></table:table-cell></table:table-row></table:table>` }) },
    { id: "styled-empty-repeat.ods", kind: "ods", expect: "handled", limit: "styled empty cells clipped to the drawn grid",
      build: () => makeOds({ automatic: '<style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:background-color="#ff0000"/></style:style>', body: '<table:table table:name="S"><table:table-row table:number-rows-repeated="1048576"><table:table-cell table:style-name="ce1" table:number-columns-repeated="16384"/></table:table-row></table:table>' }) },
    // The element budget follows the cell limit, not a byte size: before that change this 115 KiB file
    // was tokenized for about 11 s, because the budget was derived from the 128 MiB content.xml limit.
    { id: "empty-elements.ods", kind: "ods", expect: "refused", limit: "XML elements, budgeted from the cell limit",
      // Refused once 4,250,000 elements have been read, which is what a 500,000-cell workbook may hold;
      // the rest is inflating and decoding a 114 MiB part, which no element budget can reach.
      budget: { ms: 5000, heapMiB: 256, reason: "the element budget at the accepted maximum workbook size, plus decoding a 114 MiB part" },
      build: () => makeOds({ body: "<x/>".repeat(30_000_000) }) },
    // Incompressible padding puts this file under the 100:1 ratio OOXML parts are held to, so only the
    // element budget refuses it. A ratio limit is no use here anyway: ordinary workbooks pass 200:1.
    { id: "padded-empty-elements.ods", kind: "ods", expect: "refused", limit: "XML elements, independent of the compression ratio",
      budget: { ms: 4000, heapMiB: 256, reason: "the element budget at the accepted maximum workbook size" },
      build: () => makeOds({ body: `<table:table table:name="S" table:style-name="${incompressibleText(600 * 1024)}"/>${"<x/>".repeat(5_000_000)}` }) },
    // 115,000 bindings in scope under 2,000 declaring elements. Entering an element used to copy every
    // binding in scope, which cost 17 s and 1.1 GiB here and 8.3 s and 1.3 GiB for the SVG, inside every
    // element and token limit. Chained scopes read them without copying, so both are ordinary content
    // now: these entries are handled, and they fail if the copying ever comes back.
    { id: "deep-namespace-scope.ods", kind: "ods", expect: "handled", limit: "namespace lookups that do not copy the bindings in scope",
      build: () => makeOds({ body: `<table:table table:name="S"><table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row></table:table>${declarationStack(230, 500, 2_000, '<leaf xmlns:q="urn:q"/>')}` }) },
    { id: "deep-namespace-scope.svg", kind: "svg", expect: "handled", limit: "namespace lookups that do not copy the bindings in scope",
      build: () => new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg">${declarationStack(230, 500, 2_000, '<rect xmlns:q="urn:q" width="1" height="1" fill="#000"/>')}</svg>`) },
    { id: "repeat-bomb.fods", kind: "fods", expect: "refused", limit: "500,000 materialised cells",
      build: () => makeFods({ body: '<table:table table:name="S"><table:table-row table:number-rows-repeated="1048576"><table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="16384"/></table:table-row></table:table>' }) },
    // OpenDocument text (milestones 3.1 to 3.5).
    { id: "paragraph-bomb.odt", kind: "odt", expect: "refused", limit: "blocks in one document",
      build: () => makeOdt({ body: "<text:p>x</text:p>".repeat(200_000) }) },
    { id: "repeat-bomb.odt", kind: "odt", expect: "refused", limit: "table rows and columns, clamped per table",
      build: () => makeOdt({ body: '<table:table table:name="T"><table:table-row table:number-rows-repeated="1048576"><table:table-cell table:number-columns-repeated="16384"><text:p>x</text:p></table:table-cell></table:table-row></table:table>'.repeat(40) }) },
    { id: "space-bomb.odt", kind: "odt", expect: "refused", limit: "text of one paragraph",
      build: () => makeOdt({ body: `<text:p>${'<text:s text:c="10000"/>'.repeat(200)}</text:p>`.repeat(200) }) },
    { id: "deep-lists.odt", kind: "odt", expect: "handled", limit: "list nesting, which is bounded by XML depth",
      build: () => makeOdt({ body: `${"<text:list><text:list-item>".repeat(120)}<text:p>deep</text:p>${"</text:list-item></text:list>".repeat(120)}` }) },
    { id: "one-image-referenced-1000-times.odt", kind: "odt", expect: "handled", limit: "image bytes of the whole document",
      build: () => makeOdt({
        body: '<text:p><draw:frame svg:width="1cm" svg:height="1cm"><draw:image xlink:href="Pictures/a.png"/></draw:frame></text:p>'.repeat(1000),
        extra: { "Pictures/a.png": ONE_PIXEL_PNG },
      }) },
    { id: "style-chain.odt", kind: "odt", expect: "handled", limit: "style inheritance depth and cycles",
      build: () => makeOdt({
        automatic: Array.from({ length: 2000 }, (_, index) => `<style:style style:name="P${index}" style:family="paragraph" style:parent-style-name="P${index + 1}"><style:text-properties fo:color="#010203"/></style:style>`).join(""),
        body: Array.from({ length: 500 }, (_, index) => `<text:p text:style-name="P${index}">x</text:p>`).join(""),
      }) },
    { id: "inline-image-bomb.fodt", kind: "fodt", expect: "handled", limit: "decoded size of one inline image",
      build: () => makeFodt({ body: `<text:p><draw:frame svg:width="1cm" svg:height="1cm"><draw:image><office:binary-data>${"A".repeat(48 * 1024 * 1024)}</office:binary-data></draw:image></draw:frame></text:p>` }) },
    { id: "space-bomb.fodt", kind: "fodt", expect: "refused", limit: "text of one paragraph",
      build: () => makeFodt({ body: `<text:p>${'<text:s text:c="10000"/>'.repeat(200)}</text:p>`.repeat(200) }) },
    // XML structure.
    { id: "nesting-deeper-than-the-limit.docx", kind: "docx", expect: "refused", limit: "XML nesting depth",
      build: () => new Uint8Array(makeZip({ "word/document.xml": `<w:document ${W}><w:body>${"<w:sdt>".repeat(400)}${"</w:sdt>".repeat(400)}</w:body></w:document>` }, { deflate: true })) },
    { id: "too-many-attributes.pptx", kind: "pptx", expect: "refused", limit: "attributes per XML element",
      build: () => new Uint8Array(makeZip({ "ppt/presentation.xml": `<p:presentation ${P} ${Array.from({ length: 1000 }, (_, index) => `a${index}="1"`).join(" ")}/>` }, { deflate: true })) },
    // SVG.
    { id: "billion-laughs.svg", kind: "svg", expect: "refused", limit: "no DOCTYPE or entity declarations",
      build: () => new TextEncoder().encode(`<?xml version="1.0"?><!DOCTYPE svg [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;"><!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;"><!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;&c;&c;">]><svg xmlns="http://www.w3.org/2000/svg"><text>&d;&d;&d;&d;&d;&d;&d;&d;&d;&d;</text></svg>`) },
    { id: "deep-svg.svg", kind: "svg", expect: "refused", limit: "XML nesting depth",
      build: () => new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg">${"<g>".repeat(5000)}${"</g>".repeat(5000)}</svg>`) },
    { id: "many-svg-elements.svg", kind: "svg", expect: "handled", limit: "linear sanitizing work",
      build: () => new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg">${'<rect width="1" height="1" fill="#000"/>'.repeat(100_000)}</svg>`) },
  ];
}
