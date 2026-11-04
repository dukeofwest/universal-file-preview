import test from "node:test";
import assert from "node:assert/strict";
import { detectFormat } from "../dist/detection/detect-format.js";
import { makeZip } from "./helpers/zip.mjs";
import { makeOdfPackage, odfXml, minimalContent, ODS_MEDIA_TYPE } from "./helpers/odf.mjs";

const OTS_MEDIA_TYPE = "application/vnd.oasis.opendocument.spreadsheet-template";
const ODT_MEDIA_TYPE = "application/vnd.oasis.opendocument.text";
const ods = makeOdfPackage({ content: minimalContent });
const ots = makeOdfPackage({ mimetype: OTS_MEDIA_TYPE, content: minimalContent });
const odt = makeOdfPackage({ mimetype: ODT_MEDIA_TYPE, content: minimalContent });
const flat = (prefixes = {}) => new TextEncoder().encode(odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document office:version="1.3" office:mimetype="${ODS_MEDIA_TYPE}"><office:body><office:spreadsheet/></office:body></office:document>`, prefixes));

/** A one-entry ZIP local header for "mimetype" with adjustable fields, followed by a central directory. */
function mimetypeHeader({ content = ODS_MEDIA_TYPE, method = 0, flags = 0, extra = 0, name = "mimetype", uncompressed } = {}) {
  const data = Buffer.from(content), nameBytes = Buffer.from(name), local = Buffer.alloc(30);
  local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(flags, 6); local.writeUInt16LE(method, 8);
  local.writeUInt32LE(data.length, 18); local.writeUInt32LE(uncompressed ?? data.length, 22);
  local.writeUInt16LE(nameBytes.length, 26); local.writeUInt16LE(extra, 28);
  return new Uint8Array(Buffer.concat([local, nameBytes, Buffer.alloc(extra), data, Buffer.alloc(64)]));
}

test("ODF spreadsheet packages are detected from their mimetype entry", () => {
  assert.equal(detectFormat("sheet.ods", "", ods), "ods");
  assert.equal(detectFormat("sheet.ots", "", ots), "ots");
  assert.equal(detectFormat("Sheet.ODS", "", ods), "ods");
  // The package content wins over a misleading or missing name and MIME type.
  assert.equal(detectFormat("data.zip", "", ods), "ods");
  assert.equal(detectFormat("upload.bin", "application/octet-stream", ods), "ods");
  assert.equal(detectFormat("report.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ods), "ods");
  assert.equal(detectFormat("template.ods", "", ots), "ots");
});

test("other OpenDocument media types keep the previous ZIP behaviour", () => {
  assert.equal(detectFormat("letter.odt", "", odt), "unknown");
  assert.equal(detectFormat("letter.docx", "", odt), "docx");
  assert.equal(detectFormat("letter.zip", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", odt), "xlsx");
  assert.equal(detectFormat("x.zip", "", mimetypeHeader({ content: "application/zip" })), "unknown");
});

test("OOXML packages are unaffected", () => {
  const xlsx = new Uint8Array(makeZip({ "[Content_Types].xml": "<Types/>", "xl/workbook.xml": "<workbook/>" }));
  assert.equal(detectFormat("book.xlsx", "", xlsx), "xlsx");
  assert.equal(detectFormat("book.ods", "", xlsx), "ods", "an OOXML ZIP named .ods falls back to the extension, as before for any mapped extension");
  assert.equal(detectFormat("book.zip", "", xlsx), "unknown");
});

test("the mimetype entry is only trusted when stored first, uncompressed, without a data descriptor", () => {
  assert.equal(detectFormat("x.zip", "", mimetypeHeader()), "ods");
  assert.equal(detectFormat("x.zip", "", mimetypeHeader({ content: `${ODS_MEDIA_TYPE}\n` })), "ods", "surrounding whitespace is trimmed");
  assert.equal(detectFormat("x.zip", "", mimetypeHeader({ extra: 16 })), "ods", "an extra field is skipped");
  const rejected = {
    "compressed (method 8)": mimetypeHeader({ method: 8 }),
    "data descriptor flag": mimetypeHeader({ flags: 0x8 }),
    "different first entry name": mimetypeHeader({ name: "mimetypx" }),
    "extra field over 1 KiB": mimetypeHeader({ extra: 1025 }),
    "sizes disagree": mimetypeHeader({ uncompressed: 10 }),
    "empty content": mimetypeHeader({ content: "" }),
    "content over 256 bytes": mimetypeHeader({ content: ODS_MEDIA_TYPE.padEnd(257, " ") }),
    "non-ASCII content": mimetypeHeader({ content: `${ODS_MEDIA_TYPE}\u0001` }),
  };
  for (const [label, bytes] of Object.entries(rejected)) {
    assert.equal(detectFormat("x.zip", "", bytes), "unknown", label);
    assert.equal(detectFormat("x.ods", "", bytes), "ods", `${label}: falls back to the extension`);
  }
  // A package whose mimetype entry is compressed is still found by extension (makeOdfPackage storeMimetype: false).
  const compressed = makeOdfPackage({ content: minimalContent, storeMimetype: false });
  assert.equal(detectFormat("data.zip", "", compressed), "unknown");
  assert.equal(detectFormat("sheet.ods", "", compressed), "ods");
});

test("truncated package headers never throw and are only detected once complete", () => {
  for (let length = 0; length <= ods.length; length++) {
    const result = detectFormat("x.zip", "", ods.subarray(0, length));
    const complete = length >= 30 + 8 + ODS_MEDIA_TYPE.length;
    assert.equal(result, complete ? "ods" : "unknown", `length ${length}`);
  }
});

test("flat OpenDocument spreadsheets are detected by namespace and media type", () => {
  assert.equal(detectFormat("sheet.fods", "", flat()), "fods");
  assert.equal(detectFormat("sheet.fods", "", flat({ office: "o" })), "fods", "any prefix");
  assert.equal(detectFormat("export.xml", ODS_MEDIA_TYPE, flat()), "fods", "the ODS MIME type with flat XML is fods, not ods");
  assert.equal(detectFormat("export.xml", "", flat()), "unknown", "only checked for a .fods name or the ODS MIME type");
  // Without the ODF namespace, flat detection does not apply; the name or MIME type decides as usual.
  const plain = new TextEncoder().encode(`<?xml version="1.0"?><document mimetype="${ODS_MEDIA_TYPE}"/>`);
  assert.equal(detectFormat("export.xml", ODS_MEDIA_TYPE, plain), "ods");
  assert.equal(detectFormat("sheet.fods", "", new TextEncoder().encode("<svg></svg>")), "svg");
  assert.equal(detectFormat("sheet.fods", "", new TextEncoder().encode("hello")), "fods", "unrecognised content falls back to the extension, as for every format");
  // The markers must appear in the first 4 KiB.
  const late = new TextEncoder().encode(`<?xml version="1.0"?><!--${" ".repeat(5000)}-->${new TextDecoder().decode(flat()).replace(/^<\?xml[^>]*\?>/, "")}`);
  assert.equal(detectFormat("export.xml", ODS_MEDIA_TYPE, late), "ods");
});

test("existing signatures keep precedence over ODF detection", () => {
  assert.equal(detectFormat("sheet.ods", ODS_MEDIA_TYPE, Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a)), "png");
  assert.equal(detectFormat("sheet.fods", ODS_MEDIA_TYPE, new TextEncoder().encode("%PDF-1.7")), "pdf");
  assert.equal(detectFormat("sheet.ods", "", Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1)), "ods", "a CFB container named .ods resolves by extension, as for other formats");
});

test("MIME types and names cannot resolve inherited object properties", () => {
  for (const name of ["report.constructor", "x.__proto__", "x.hasownproperty", "x.valueof"]) assert.equal(detectFormat(name), "unknown", name);
  for (const mime of ["constructor", "__proto__", "toString"]) assert.equal(detectFormat("x.bin", mime), "unknown", mime);
  for (const bytes of [ods, new Uint8Array(makeZip({ a: "b" })), Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1)]) assert.equal(detectFormat("x.constructor", "constructor", bytes), bytes === ods ? "ods" : "unknown");
});
