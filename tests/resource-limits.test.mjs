import test from "node:test";
import assert from "node:assert/strict";
import { deflateRawSync, crc32 as zlibCrc32 } from "node:zlib";
import { ZipArchive } from "../dist/archive/zip-reader.js";
import { OoxmlPackage, OOXML_LIMITS } from "../dist/package/ooxml-package.js";
import { parseDocx, parsePptx } from "../dist/parsers/ooxml.js";
import { parseXlsx } from "../dist/parsers/xlsx.js";
import { parseXml } from "../dist/xml/xml-parser.js";
import { tiffToPng } from "../dist/parsers/tiff.js";
import { createImageUrl, imageDimensions, revokeImageUrls, MAX_IMAGE_PIXELS, MAX_DOCUMENT_PIXELS } from "../dist/renderers/image-url.js";
import { makeZip } from "./helpers/zip.mjs";

// Resource limits for hostile documents: decompression budget and part caching, ZIP structure,
// reference limits, XLSX work limits, XML attribute limits and image dimensions.

const MiB = 1024 * 1024;
const big = Buffer.alloc(32 * MiB - 64); // just under the per-part limit; deflates to about 32 KiB
const failsWith = (code, pattern) => error => error?.name === "PreviewError" && error.code === code && pattern.test(error.message);
const rssGrowth = (work) => { globalThis.gc?.(); const before = process.memoryUsage().rss, result = work(); return { result, growth: process.memoryUsage().rss - before }; };
const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

// ------------------------------------------------------------------------------------------------
// Decompression budget and caching (DOCX, PPTX, XLSX)
// ------------------------------------------------------------------------------------------------

test("a DOCX with 1,000 relationships and 1,000 uses of one 32 MiB image decompresses it once", () => {
  const rels = Array.from({ length: 1000 }, (_, index) => `<Relationship Id="r${index}" Type="${IMAGE_REL}" Target="media/a.png"/>`).join("");
  const runs = Array.from({ length: 1000 }, (_, index) => `<w:r><w:drawing><a:blip r:embed="r${index}"/></w:drawing></w:r>`).join("");
  const file = makeZip({
    "word/document.xml": `<w:document xmlns:w="urn:w" xmlns:a="urn:a" xmlns:r="urn:r"><w:body><w:p>${runs}</w:p></w:body></w:document>`,
    "word/_rels/document.xml.rels": `<Relationships>${rels}</Relationships>`,
    "word/media/a.png": big,
  }, { deflate: true });
  assert.ok(file.length < 100_000, `fixture is small (${file.length} bytes)`);
  const started = performance.now(), { result: model, growth } = rssGrowth(() => parseDocx(new Uint8Array(file), "x.docx"));
  const images = model.blocks.flatMap(block => block.type === "paragraph" ? block.value.runs.filter(run => run.image) : []).map(run => run.image.bytes);
  assert.equal(images.length, 1000);
  assert.ok(images.every(bytes => bytes === images[0]), "every use shares one decompressed copy");
  assert.ok(growth < 200 * MiB, `memory grew ${Math.round(growth / MiB)} MiB`);
  assert.ok(performance.now() - started < 5000);
});

/**
 * A DOCX with nine unused 32 MiB images: 288 MiB, more than the 256 MiB budget, so reading them all
 * would fail. Built once (compressing it is the expensive part) and shared by the tests below.
 */
let nineImages;
const nineImageDocx = () => nineImages ??= makeZip({
  "word/document.xml": '<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>text</w:t></w:r></w:p></w:body></w:document>',
  "word/_rels/document.xml.rels": `<Relationships>${Array.from({ length: 9 }, (_, index) => `<Relationship Id="r${index}" Type="${IMAGE_REL}" Target="media/a${index}.png"/>`).join("")}</Relationships>`,
  ...Object.fromEntries(Array.from({ length: 9 }, (_, index) => [`word/media/a${index}.png`, big])),
}, { deflate: true });

test("unused DOCX media is never decompressed", () => {
  // Eager reading would exceed the budget and throw; lazy reading never touches the media.
  const { growth } = rssGrowth(() => parseDocx(new Uint8Array(nineImageDocx()), "x.docx"));
  assert.ok(growth < 100 * MiB, `memory grew ${Math.round(growth / MiB)} MiB for 288 MiB of unused media`);
});

test("a PPTX listing one slide 10,000 times parses it once", () => {
  const P = 'xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r"';
  const ids = Array.from({ length: 10_000 }, (_, index) => `<p:sldId id="${256 + index}" r:id="rId1"/>`).join("");
  const file = makeZip({
    "ppt/presentation.xml": `<p:presentation ${P}><p:sldIdLst>${ids}</p:sldIdLst></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>',
    "ppt/slides/slide1.xml": `<p:sld ${P}><p:cSld><p:spTree><p:pic><p:blipFill><a:blip r:embed="i1"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:pic></p:spTree></p:cSld></p:sld>`,
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships><Relationship Id="i1" Type="${IMAGE_REL}" Target="../media/a.png"/></Relationships>`,
    "ppt/media/a.png": big,
  }, { deflate: true });
  const started = performance.now(), { result: model, growth } = rssGrowth(() => parsePptx(new Uint8Array(file), "x.pptx"));
  assert.equal(model.slides.length, 1);
  assert.ok(growth < 200 * MiB, `memory grew ${Math.round(growth / MiB)} MiB`);
  assert.ok(performance.now() - started < 5000);
});

const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const workbook = (sheets, sheetXml = '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>', sheetRels = "") => makeZip({
  "xl/workbook.xml": `<workbook ${S}><sheets>${sheets}</sheets></workbook>`,
  "xl/_rels/workbook.xml.rels": '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  "xl/worksheets/sheet1.xml": `<worksheet ${S}>${sheetXml}</worksheet>`,
  ...(sheetRels ? { "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships>${sheetRels}</Relationships>` } : {}),
}, { deflate: true });

test("an XLSX listing one worksheet 1,000 times reads it once; more than 1,024 sheets is a resource limit", () => {
  const sheets = count => Array.from({ length: count }, (_, index) => `<sheet name="S${index}" sheetId="${index + 1}" r:id="rId1"/>`).join("");
  assert.equal(parseXlsx(new Uint8Array(workbook(sheets(1000))), "x.xlsx").sheets.length, 1);
  assert.throws(() => parseXlsx(new Uint8Array(workbook(sheets(1025))), "x.xlsx"), failsWith("resource-limit", /more than 1024 sheets/));
});

test("the package budget counts each part once and stops at 256 MiB of distinct parts", () => {
  const pkg = OoxmlPackage.open(new Uint8Array(nineImageDocx()));
  for (let repeat = 0; repeat < 3; repeat++) for (let index = 0; index < 8; index++) pkg.read(`word/media/a${index}.png`);
  assert.equal(pkg.totalBytes, 8 * big.length, "repeated reads are free");
  assert.ok(pkg.totalBytes <= OOXML_LIMITS.maxTotalBytes);
  assert.throws(() => pkg.read("word/media/a8.png"), failsWith("resource-limit", /total size limit/));
  assert.ok(pkg.totalBytes <= OOXML_LIMITS.maxTotalBytes, "nothing is decompressed past the budget");
});

test("XML parts that expand more than 100 times are refused before parsing; small parts are exempt", () => {
  // Without this, a 43 KiB PPTX of 16 slides x 249,000 empty shapes took about 7 s to parse within the byte budget.
  const P = 'xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r"';
  const slides = 16, entries = {
    "ppt/presentation.xml": `<p:presentation ${P}><p:sldIdLst>${Array.from({ length: slides }, (_, index) => `<p:sldId id="${256 + index}" r:id="s${index}"/>`).join("")}</p:sldIdLst></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": `<Relationships>${Array.from({ length: slides }, (_, index) => `<Relationship Id="s${index}" Target="slides/slide${index}.xml"/>`).join("")}</Relationships>`,
  };
  for (let index = 0; index < slides; index++) entries[`ppt/slides/slide${index}.xml`] = `<p:sld ${P}><p:cSld><p:spTree>${"<p:sp/>".repeat(249_000)}</p:spTree></p:cSld></p:sld>`;
  const file = makeZip(entries, { deflate: true }), started = performance.now();
  assert.ok(file.length < 64 * 1024);
  assert.throws(() => parsePptx(new Uint8Array(file), "x.pptx"), failsWith("resource-limit", /expands more than 100 times/));
  assert.ok(performance.now() - started < 1000);
  // Under 1 MiB, a highly compressible part is still parsed.
  const small = OoxmlPackage.open(new Uint8Array(makeZip({ "a.xml": `<a>${"<b/>".repeat(100_000)}</a>` }, { deflate: true })));
  assert.equal(small.xml("a.xml").children.length, 100_000);
});

test("parsed XML parts are cached and shared", () => {
  const pkg = OoxmlPackage.open(new Uint8Array(makeZip({ "a.xml": "<a><b/></a>" })));
  assert.equal(pkg.xml("a.xml"), pkg.xml("/a.xml"));
  assert.equal(pkg.read("a.xml"), pkg.read("a.xml"));
});

// ------------------------------------------------------------------------------------------------
// ZIP structure
// ------------------------------------------------------------------------------------------------

/**
 * Builds a stored ZIP from explicit records placed at `offset` in a buffer of `totalLength` bytes:
 * { name, data, offset, localName?, localSize?, localCrc?, skipLocal?, descriptor? ("signed" or
 * "unsigned": flag bit 3, zero sizes in the local header, a data descriptor after the data),
 * descriptorCrc?, localFlags?, centralFlags? }.
 */
function rawZip(records, totalLength) {
  const buffer = Buffer.alloc(totalLength), central = [];
  for (const record of records) {
    const name = Buffer.from(record.name), localName = Buffer.from(record.localName ?? record.name), data = Buffer.from(record.data), crc = zlibCrc32(data) >>> 0;
    const flags = record.descriptor ? 8 : 0, local = Buffer.alloc(30 + localName.length);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(record.localFlags ?? flags, 6);
    local.writeUInt32LE(record.localCrc ?? (record.descriptor ? 0 : crc), 14);
    local.writeUInt32LE(record.localSize ?? (record.descriptor ? 0 : data.length), 18); local.writeUInt32LE(record.localSize ?? (record.descriptor ? 0 : data.length), 22); local.writeUInt16LE(localName.length, 26); localName.copy(local, 30);
    if (!record.skipLocal) {
      local.copy(buffer, record.offset); data.copy(buffer, record.offset + local.length);
      if (record.descriptor) {
        const descriptor = Buffer.alloc(16);
        descriptor.writeUInt32LE(0x08074b50, 0); descriptor.writeUInt32LE(record.descriptorCrc ?? crc, 4); descriptor.writeUInt32LE(data.length, 8); descriptor.writeUInt32LE(data.length, 12);
        (record.descriptor === "signed" ? descriptor : descriptor.subarray(4)).copy(buffer, record.offset + local.length + data.length);
      }
    }
    const entry = Buffer.alloc(46 + name.length);
    entry.writeUInt32LE(0x02014b50, 0); entry.writeUInt16LE(20, 4); entry.writeUInt16LE(20, 6); entry.writeUInt16LE(record.centralFlags ?? flags, 8); entry.writeUInt32LE(crc, 16);
    entry.writeUInt32LE(data.length, 20); entry.writeUInt32LE(data.length, 24); entry.writeUInt16LE(name.length, 28); entry.writeUInt32LE(record.offset, 42); name.copy(entry, 46);
    central.push(entry);
  }
  const directory = Buffer.concat(central), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(records.length, 8); end.writeUInt16LE(records.length, 10); end.writeUInt32LE(directory.length, 12); end.writeUInt32LE(totalLength, 16);
  return new Uint8Array(Buffer.concat([buffer, directory, end]));
}

test("ZIP entries whose data overlaps another entry are rejected", () => {
  // Entry "b" has a complete, consistent local record hidden inside entry "a"'s stored data.
  const inner = { name: "b", data: "inner", offset: 40 };
  const innerRecord = rawZip([inner], 80).subarray(0, 80);
  const outerData = Buffer.alloc(100); Buffer.from(innerRecord.subarray(40, 40 + 30 + 1 + 5)).copy(outerData, 9);
  const archive = rawZip([{ name: "a", data: outerData, offset: 0 }, { ...inner, skipLocal: true }], 31 + 100);
  assert.throws(() => ZipArchive.open(archive), failsWith("invalid-file", /^ZIP entries overlap\.$/));
});

test("ZIP local headers must match the central directory, checked on open even for entries that are never read", () => {
  const read = archive => new TextDecoder().decode(ZipArchive.open(archive).read("a.txt"));
  assert.equal(read(rawZip([{ name: "a.txt", data: "xyz", offset: 0 }], 38)), "xyz", "a consistent archive still reads");
  assert.throws(() => ZipArchive.open(rawZip([{ name: "a.txt", localName: "b.txt", data: "x", offset: 0 }], 36)), failsWith("invalid-file", /names differ/));
  // The second entry is never read; its inconsistent metadata is still found when the archive is opened.
  const withUnused = unused => rawZip([{ name: "a.txt", data: "xyz", offset: 0 }, { name: "b.txt", data: "uvw", offset: 38, ...unused }], 76);
  assert.equal(read(withUnused({})), "xyz");
  assert.throws(() => ZipArchive.open(withUnused({ localSize: 2 })), failsWith("invalid-file", /sizes differ/));
  assert.throws(() => ZipArchive.open(withUnused({ localCrc: 1 })), failsWith("invalid-file", /sizes differ/));
  assert.throws(() => ZipArchive.open(withUnused({ localFlags: 8 })), failsWith("invalid-file", /flags differ/), "data-descriptor flag only in the local header");
  assert.throws(() => ZipArchive.open(withUnused({ centralFlags: 8 })), failsWith("invalid-file", /flags differ/), "data-descriptor flag only in the central directory");
});

test("ZIP data descriptors are validated and counted in the entry's range", () => {
  const read = archive => new TextDecoder().decode(ZipArchive.open(archive).read("a.txt"));
  // Valid archives with data descriptors, with and without the optional signature, still open and read.
  assert.equal(read(rawZip([{ name: "a.txt", data: "xyz", offset: 0, descriptor: "signed" }], 54)), "xyz");
  assert.equal(read(rawZip([{ name: "a.txt", data: "xyz", offset: 0, descriptor: "unsigned" }], 50)), "xyz");
  assert.equal(read(rawZip([{ name: "a.txt", data: "xyz", offset: 0, descriptor: "signed" }, { name: "b.txt", data: "uvw", offset: 54, descriptor: "unsigned" }], 104)), "xyz");
  // A descriptor that disagrees with the central directory, on an entry that is never read.
  assert.throws(() => ZipArchive.open(rawZip([{ name: "a.txt", data: "xyz", offset: 0 }, { name: "b.txt", data: "uvw", offset: 38, descriptor: "signed", descriptorCrc: 1 }], 92)), failsWith("invalid-file", /data descriptor/));
  // The descriptor of "a" (16 bytes after its data, at 38-53) overlaps the local header of "b" at 46.
  assert.throws(() => ZipArchive.open(rawZip([{ name: "a.txt", data: "xyz", offset: 0, descriptor: "signed" }, { name: "b.txt", data: "uvw", offset: 46 }], 84)), failsWith("invalid-file", /overlap|data descriptor/));
  // A descriptor that would run into the central directory.
  assert.throws(() => ZipArchive.open(rawZip([{ name: "a.txt", data: "xyz", offset: 0, descriptor: "signed" }], 44)), failsWith("invalid-file", /data descriptor/));
});

test("ZIP entry data may not run into the central directory", () => {
  const archive = Buffer.from(rawZip([{ name: "a.txt", data: "xyz", offset: 0 }], 38));
  archive.writeUInt32LE(50, 38 + 20); // central compressed size reaching past the local data
  assert.throws(() => ZipArchive.open(new Uint8Array(archive)), failsWith("invalid-file", /overlaps the central directory/));
});

// ------------------------------------------------------------------------------------------------
// XLSX work limits
// ------------------------------------------------------------------------------------------------

const link = '<Relationship Id="h1" Type="x/hyperlink" Target="https://example.invalid/" TargetMode="External"/>';

test("hyperlink ranges expand to at most 1,000,000 cells per workbook", () => {
  const links = Array.from({ length: 400 }, () => '<hyperlink ref="A1:CW1001" r:id="h1"/>').join("");
  const started = performance.now();
  const model = parseXlsx(new Uint8Array(workbook('<sheet name="S" sheetId="1" r:id="rId1"/>', `<sheetData/><hyperlinks>${links}</hyperlinks>`, link)), "x.xlsx");
  assert.ok(Object.keys(model.sheets[0].hyperlinks).length <= 1_000_000);
  assert.ok(performance.now() - started < 5000, "400 large ranges no longer take minutes");
  const one = parseXlsx(new Uint8Array(workbook('<sheet name="S" sheetId="1" r:id="rId1"/>', '<sheetData/><hyperlinks><hyperlink ref="B2:C3" r:id="h1"/></hyperlinks>', link)), "x.xlsx");
  assert.deepEqual(Object.keys(one.sheets[0].hyperlinks).sort(), ["B2", "B3", "C2", "C3"]);
});

test("conditional formats share one workbook-wide budget across sheets", () => {
  // Before: 40 sheets, each formatting 250,000 cells with a formula that reads no cells, took about 28 s from a 15 KiB file.
  const sheets = 40, entries = {
    "xl/workbook.xml": `<workbook ${S}><sheets>${Array.from({ length: sheets }, (_, index) => `<sheet name="S${index}" sheetId="${index + 1}" r:id="s${index}"/>`).join("")}</sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships>${Array.from({ length: sheets }, (_, index) => `<Relationship Id="s${index}" Type="x/worksheet" Target="worksheets/sheet${index}.xml"/>`).join("")}</Relationships>`,
    "xl/styles.xml": `<styleSheet ${S}><dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>`,
  };
  for (let index = 0; index < sheets; index++) entries[`xl/worksheets/sheet${index}.xml`] = `<worksheet ${S}><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="10000"><c r="Z10000"><v>1</v></c></row></sheetData><conditionalFormatting sqref="A1:Z10000"><cfRule type="expression" dxfId="0" priority="1"><formula>1+1+1+1+1+1+1+1</formula></cfRule></conditionalFormatting></worksheet>`;
  const file = makeZip(entries, { deflate: true }), started = performance.now(), model = parseXlsx(new Uint8Array(file), "x.xlsx"), elapsed = performance.now() - started;
  assert.ok(file.length < 32 * 1024);
  assert.ok(elapsed < 10_000, `took ${Math.round(elapsed)} ms`);
  assert.ok(Object.keys(model.sheets[0].conditional ?? {}).length > 0, "the first sheet is still formatted");
  const formatted = model.sheets.reduce((sum, sheet) => sum + Object.keys(sheet.conditional ?? {}).length, 0);
  assert.ok(formatted <= 1_000_000, `${formatted} formatted cells`);
});

test("table filter buttons and slicer items are bounded across the workbook", async () => {
  const { ooxmlFixtures } = await import("./helpers/ooxml-fixtures.mjs");
  const entries = ooxmlFixtures.find(fixture => fixture.id === "xlsx-slicer").entries();
  const normal = parseXlsx(new Uint8Array(makeZip(entries)), "x.xlsx");
  // Stretch the slicer's table to the last row of the sheet: items must be the same and the parse fast.
  const stretched = { ...entries, "xl/tables/table1.xml": entries["xl/tables/table1.xml"].replace('ref="A1:B4"', 'ref="A1:B1048576"').replace('<autoFilter ref="A1:B4">', '<autoFilter ref="A1:B1048576">') };
  const started = performance.now(), model = parseXlsx(new Uint8Array(makeZip(stretched)), "x.xlsx");
  assert.ok(performance.now() - started < 2000);
  const items = workbook => workbook.sheets.flatMap(sheet => sheet.drawings ?? []).filter(drawing => drawing.kind === "slicer").map(slicer => slicer.items.map(item => item.text));
  assert.deepEqual(items(model).map(list => list.filter(text => text !== "(blank)")), items(normal), "the same items, plus at most one (blank)");
  // An autoFilter across all 16,384 columns on 1,024 sheets is limited workbook-wide.
  const wide = Array.from({ length: 1024 }, (_, index) => `<sheet name="S${index}" sheetId="${index + 1}" r:id="s${index}"/>`).join("");
  const wideEntries = {
    "xl/workbook.xml": `<workbook ${S}><sheets>${wide}</sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships>${Array.from({ length: 1024 }, (_, index) => `<Relationship Id="s${index}" Type="x/worksheet" Target="worksheets/sheet${index}.xml"/>`).join("")}</Relationships>`,
  };
  for (let index = 0; index < 1024; index++) wideEntries[`xl/worksheets/sheet${index}.xml`] = `<worksheet ${S}><sheetData/><autoFilter ref="A1:XFD1"/></worksheet>`;
  const filterStarted = performance.now(), filtered = parseXlsx(new Uint8Array(makeZip(wideEntries, { deflate: true })), "x.xlsx");
  assert.ok(performance.now() - filterStarted < 10_000);
  assert.ok(filtered.sheets.reduce((sum, sheet) => sum + (sheet.filterButtons?.length ?? 0), 0) <= 100_000);
});

test("slicer work counts every filtered column checked per row", async () => {
  // A table with 1,000 filtered columns over 100,000 rows. The filters accept blank cells, so every row is checked
  // against all of them: 100 million filter checks without a shared budget.
  const { ooxmlFixtures } = await import("./helpers/ooxml-fixtures.mjs");
  const entries = ooxmlFixtures.find(fixture => fixture.id === "xlsx-slicer").entries();
  const filters = Array.from({ length: 1000 }, (_, index) => `<filterColumn colId="${index + 1}"><filters blank="1"/></filterColumn>`).join("");
  const crafted = {
    ...entries,
    "xl/worksheets/sheet1.xml": entries["xl/worksheets/sheet1.xml"].replace("</sheetData>", '<row r="100000"><c r="A100000" t="inlineStr"><is><t>far</t></is></c></row></sheetData>'),
    "xl/tables/table1.xml": entries["xl/tables/table1.xml"].replace('ref="A1:B4"', 'ref="A1:ALM100000"').replace('<autoFilter ref="A1:B4">', '<autoFilter ref="A1:ALM100000">').replace("</filterColumn></autoFilter>", `</filterColumn>${filters}</autoFilter>`),
  };
  const started = performance.now(), model = parseXlsx(new Uint8Array(makeZip(crafted)), "x.xlsx"), elapsed = performance.now() - started;
  assert.ok(elapsed < 5000, `took ${Math.round(elapsed)} ms`);
  const slicer = model.sheets.flatMap(sheet => sheet.drawings ?? []).find(drawing => drawing.kind === "slicer");
  assert.ok(slicer && slicer.items.length > 0, "the slicer still shows the items read within the budget");
});

test("an XLSX workbook may hold at most 500,000 parsed cells across all its sheets", () => {
  // Each sheet (120,000 cells) is far below the per-sheet XML element limit; five of them exceed the workbook limit.
  const row = index => `<row r="${index + 1}">${Array.from({ length: 120 }, (_, column) => `<c r="${String.fromCharCode(65 + Math.floor(column / 26))}${String.fromCharCode(65 + column % 26)}${index + 1}"><v>1</v></c>`).join("")}</row>`;
  const sheetXml = `<worksheet ${S}><sheetData>${Array.from({ length: 1000 }, (_, index) => row(index)).join("")}</sheetData></worksheet>`;
  const book = count => {
    const entries = {
      "xl/workbook.xml": `<workbook ${S}><sheets>${Array.from({ length: count }, (_, index) => `<sheet name="S${index}" sheetId="${index + 1}" r:id="s${index}"/>`).join("")}</sheets></workbook>`,
      "xl/_rels/workbook.xml.rels": `<Relationships>${Array.from({ length: count }, (_, index) => `<Relationship Id="s${index}" Type="x/worksheet" Target="worksheets/sheet${index}.xml"/>`).join("")}</Relationships>`,
    };
    for (let index = 0; index < count; index++) entries[`xl/worksheets/sheet${index}.xml`] = sheetXml;
    return new Uint8Array(makeZip(entries, { deflate: true }));
  };
  assert.equal(parseXlsx(book(1), "x.xlsx").sheets[0].cells.length, 120_000);
  assert.throws(() => parseXlsx(book(5), "x.xlsx"), failsWith("resource-limit", /more than 500,000 cells/));
});

const formulaSheet = (cells) => workbook('<sheet name="S" sheetId="1" r:id="rId1"/>', `<sheetData>${cells.map((cell, index) => `<row r="${index + 1}">${cell}</row>`).join("")}</sheetData>`);
const cellValue = (model, ref) => model.sheets[0].cells.find(cell => cell.ref === ref)?.value;

test("uncached formulas stop at the work limit instead of running for minutes", () => {
  const cells = Array.from({ length: 20 }, (_, index) => `<c r="A${index + 1}"><f>SUM(B1:CW100000)</f></c>`);
  const started = performance.now(), model = parseXlsx(new Uint8Array(formulaSheet(cells)), "x.xlsx"), elapsed = performance.now() - started;
  assert.ok(elapsed < 10_000, `took ${Math.round(elapsed)} ms`);
  assert.equal(cellValue(model, "A1"), "0", "formulas within the budget are still calculated");
  assert.equal(cellValue(model, "A20"), "", "formulas past the budget stay uncalculated");
});

test("deep chains of dependent formulas neither overflow the stack nor run unbounded", () => {
  const chain = depth => Array.from({ length: depth }, (_, index) => `<c r="A${index + 1}"><f>${index + 1 < depth ? `A${index + 2}+1` : "1"}</f></c>`);
  assert.equal(cellValue(parseXlsx(new Uint8Array(formulaSheet(chain(150))), "x.xlsx"), "A1"), "150", "a chain within the depth limit is calculated");
  // Without a depth limit, 60,000 nested evaluations overflow the call stack (RangeError) and the whole preview fails.
  const started = performance.now(), model = parseXlsx(new Uint8Array(formulaSheet(chain(60_000))), "x.xlsx");
  assert.ok(performance.now() - started < 15_000);
  assert.equal(cellValue(model, "A1"), "", "a chain deeper than the limit is left uncalculated");
});

// ------------------------------------------------------------------------------------------------
// XML
// ------------------------------------------------------------------------------------------------

test("parseXml limits attributes per element and name length, like the namespace-aware tokenizer", () => {
  const attributes = count => Array.from({ length: count }, (_, index) => `a${index}=""`).join(" ");
  assert.equal(Object.keys(parseXml(`<r ${attributes(512)}/>`).attributes).length, 512);
  assert.throws(() => parseXml(`<r ${attributes(513)}/>`), failsWith("resource-limit", /attributes/));
  assert.equal(parseXml(`<${"a".repeat(1024)}/>`).name.length, 1024);
  assert.throws(() => parseXml(`<${"a".repeat(1025)}/>`), failsWith("resource-limit", /name/));
  assert.throws(() => parseXml(`<r ${"b".repeat(1025)}=""/>`), failsWith("resource-limit", /name/));
});

// ------------------------------------------------------------------------------------------------
// Images
// ------------------------------------------------------------------------------------------------

test("TIFF tiles larger than the image are refused before any buffer is sized from them", () => {
  const tiff = (tileWidth, tileHeight) => {
    const entries = [[256, 3, 1], [257, 3, 1], [258, 3, 8], [259, 3, 1], [262, 3, 1], [277, 3, 1], [322, 4, tileWidth], [323, 4, tileHeight], [324, 4, 200], [325, 4, 1]];
    const bytes = Buffer.alloc(8 + 2 + entries.length * 12 + 4 + 300);
    bytes.write("II", 0); bytes.writeUInt16LE(42, 2); bytes.writeUInt32LE(8, 4); bytes.writeUInt16LE(entries.length, 8);
    entries.forEach(([tag, type, value], index) => { const at = 10 + index * 12; bytes.writeUInt16LE(tag, at); bytes.writeUInt16LE(type, at + 2); bytes.writeUInt32LE(1, at + 4); if (type === 3) bytes.writeUInt16LE(value, at + 8); else bytes.writeUInt32LE(value, at + 8); });
    return new Uint8Array(bytes);
  };
  const started = performance.now();
  assert.equal(tiffToPng(tiff(40_000, 40_000)), undefined);
  assert.equal(tiffToPng(tiff(0, 16)), undefined);
  assert.ok(performance.now() - started < 1000);
  assert.ok(tiffToPng(tiff(16, 16)) instanceof Uint8Array, "a 1x1 image in one 16x16 tile still converts");
});

/** Minimal headers declaring the given size (enough for imageDimensions, not decodable). */
const headers = {
  png: (w, h) => { const b = Buffer.alloc(33); Buffer.from("89504e470d0a1a0a0000000d49484452", "hex").copy(b); b.writeUInt32BE(w, 16); b.writeUInt32BE(h, 20); return b; },
  gif: (w, h) => { const b = Buffer.alloc(13); b.write("GIF89a"); b.writeUInt16LE(w, 6); b.writeUInt16LE(h, 8); return b; },
  bmp: (w, h) => { const b = Buffer.alloc(54); b.write("BM"); b.writeInt32LE(40, 14); b.writeInt32LE(w, 18); b.writeInt32LE(-h, 22); return b; },
  bmpOs2: (w, h) => { const b = Buffer.alloc(26); b.write("BM"); b.writeInt32LE(12, 14); b.writeUInt16LE(w, 18); b.writeUInt16LE(h, 20); return b; },
  webpLossy: (w, h) => { const b = Buffer.alloc(30); b.write("RIFF"); b.write("WEBPVP8 ", 8); b.writeUInt16LE(w, 26); b.writeUInt16LE(h, 28); return b; },
  webpLossless: (w, h) => { const b = Buffer.alloc(25); b.write("RIFF"); b.write("WEBPVP8L", 8); const bits = (w - 1) | ((h - 1) << 14); b.writeUInt32LE(bits >>> 0, 21); return b; },
  webpExtended: (w, h) => { const b = Buffer.alloc(30); b.write("RIFF"); b.write("WEBPVP8X", 8); b.writeUIntLE(w - 1, 24, 3); b.writeUIntLE(h - 1, 27, 3); return b; },
  jpeg: (w, h) => Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x04, 0x00, 0x00, 0xff, 0xc0, 0x00, 0x11, 0x08, h >> 8, h & 0xff, w >> 8, w & 0xff, 0x03, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
};
const types = { png: "image/png", gif: "image/gif", bmp: "image/bmp", bmpOs2: "image/bmp", webpLossy: "image/webp", webpLossless: "image/webp", webpExtended: "image/webp", jpeg: "image/jpeg" };

test("imageDimensions reads PNG, GIF, BMP, WebP and JPEG headers", () => {
  for (const [kind, build] of Object.entries(headers)) assert.deepEqual(imageDimensions(new Uint8Array(build(1234, 567)), types[kind]), { width: 1234, height: 567 }, kind);
  assert.equal(imageDimensions(new Uint8Array(4), "image/png"), undefined);
});

test("createImageUrl refuses images over 40 megapixels and over the per-document total", () => {
  const png = (w, h) => new Uint8Array(headers.png(w, h));
  assert.equal(MAX_IMAGE_PIXELS, 40_000_000);
  assert.equal(createImageUrl(png(10_000, 4_001), "image/png"), undefined, "40,010,000 pixels");
  assert.equal(createImageUrl(png(0, 10), "image/png"), undefined, "zero width");
  assert.equal(createImageUrl(new Uint8Array(headers.jpeg(65_535, 65_535)), "image/jpeg"), undefined);
  const live = [];
  for (let index = 0; index < 3; index++) { const url = createImageUrl(png(8000, 5000), "image/png", live); assert.ok(url, `image ${index + 1} fits`); live.push(url); }
  assert.equal(3 * 40_000_000, MAX_DOCUMENT_PIXELS);
  assert.equal(createImageUrl(png(10, 10), "image/png", live), undefined, "the document total is reached");
  revokeImageUrls(live);
  assert.equal(live.length, 0);
  const again = createImageUrl(png(10, 10), "image/png", live);
  assert.ok(again, "released images no longer count");
  revokeImageUrls([again]);
});
