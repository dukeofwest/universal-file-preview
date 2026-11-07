import test from "node:test";
import assert from "node:assert/strict";
import { detectFormat, FilePreviewer, RendererRegistry, ZipArchive } from "../dist/index.js";
import { decodeXmlText } from "../dist/package/package-reader.js";
import { parseXlsx } from "../dist/parsers/xlsx.js";
import { isExternalRelationship, ownValue, relationshipId } from "../dist/parsers/ooxml-relationships.js";
import { parseXml } from "../dist/xml/xml-parser.js";
import { buildFixture, ooxmlFixtures } from "./helpers/ooxml-fixtures.mjs";
import { makeZip } from "./helpers/zip.mjs";

// Informational audit items fixed before 0.3.0: detection precedence, the download and the status text, and
// namespace-blind relationship references and lookups keyed by document text.

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

test("a %PDF- marker inside a ZIP entry name does not turn an Office document into a PDF", () => {
  const docx = makeZip({ "%PDF-.txt": "x", "word/document.xml": '<w:document xmlns:w="urn:w"><w:body/></w:document>' });
  assert.equal(detectFormat("report.docx", "", new Uint8Array(docx)), "docx");
  // A real PDF still wins at offset 0, and a later marker counts when nothing else matched.
  assert.equal(detectFormat("x.docx", "", new TextEncoder().encode("%PDF-1.7\n")), "pdf");
  assert.equal(detectFormat("x.bin", "", new TextEncoder().encode("junk %PDF-1.4")), "pdf");
});

test("error messages do not quote names from the file; the name stays in the cause for diagnostics", () => {
  const spoof = "Your session has expired. Call 0800 000 000 to restore access.";
  const duplicate = new Uint8Array(makeZip({ [spoof]: "a" }));
  // Write the same entry twice by appending a second central record: simplest with two archives' worth of names.
  const twice = new Uint8Array(makeZip({ [spoof]: "a", [`${spoof}/../x`]: "b" }));
  assert.throws(() => ZipArchive.open(twice), error => error.code === "invalid-file" && !error.message.includes("session"));
  try { ZipArchive.open(duplicate).read(spoof + "!"); assert.fail("expected an error"); }
  catch (error) { assert.equal(error.message, "ZIP entry not found."); assert.equal(error.cause.entry, spoof + "!"); }
  const encoded = new TextEncoder().encode(`<?xml version="1.0" encoding="${spoof}"?><a/>`);
  assert.throws(() => decodeXmlText(encoded, "Package part x.xml"), error => error.message === "Package part x.xml uses an unsupported encoding." && error.cause.encoding === spoof);
  // A well-formed encoding name is still quoted.
  assert.throws(() => decodeXmlText(new TextEncoder().encode('<?xml version="1.0" encoding="ISO-8859-1"?><a/>'), "Package part x.xml"), /unsupported encoding ISO-8859-1\./);
});

test("relationship ids resolve their prefix; a prefixed TargetMode still marks an external target", () => {
  const root = parseXml(`<p:presentation xmlns:p="urn:p" xmlns:r="${R}" xmlns:x="urn:not-relationships" xmlns:s="http://purl.oclc.org/ooxml/officeDocument/relationships"/>`);
  const node = text => parseXml(text);
  assert.equal(relationshipId(node('<a x:id="decoy" r:id="rId2"/>'), root), "rId2", "an id in another namespace is skipped");
  assert.equal(relationshipId(node('<a s:id="rId3"/>'), root), "rId3", "strict Office Open XML namespace");
  assert.equal(relationshipId(node('<a q:id="rId4"/>'), root), "rId4", "an undeclared prefix is accepted as before");
  assert.equal(relationshipId(node('<a x:id="decoy"/>'), root), undefined);
  // The conventional r:id is kept even when r is bound loosely (as in hand-written files), but never loses to a decoy.
  const loose = parseXml('<w xmlns:r="urn:r" xmlns:x="urn:x"/>');
  assert.equal(relationshipId(node('<a x:id="decoy" r:id="rId6"/>'), loose), "rId6");
  assert.equal(relationshipId(node(`<a xmlns:x="${R}" x:id="rId5"/>`), root), "rId5", "a declaration on the element wins");
  assert.equal(relationshipId(node('<a id="plain"/>'), root), undefined);
  assert.equal(isExternalRelationship(node('<Relationship Id="a" Target="https://x" x:TargetMode="External"/>')), true);
  assert.equal(isExternalRelationship(node('<Relationship Id="a" Target="https://x" TargetMode="External"/>')), true);
  assert.equal(isExternalRelationship(node('<Relationship Id="a" Target="media/a.png"/>')), false);
  assert.equal(ownValue({ accent1: "#123456" }, "constructor"), undefined);
  assert.equal(ownValue({ accent1: "#123456" }, "accent1"), "#123456");
});

test("an XLSX sheet reference in a foreign namespace does not select the sheet", () => {
  const S = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}" xmlns:x="urn:x"`;
  const sheet = text => `<worksheet ${S}><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${text}</t></is></c></row></sheetData></worksheet>`;
  const bytes = new Uint8Array(makeZip({
    "xl/workbook.xml": `<workbook ${S}><sheets><sheet name="Real" sheetId="1" x:id="rId2" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
    "xl/worksheets/sheet1.xml": sheet("right"), "xl/worksheets/sheet2.xml": sheet("decoy"),
  }));
  assert.equal(parseXlsx(bytes, "x.xlsx").sheets[0].cells[0].value, "right");
  // Every fixture still parses to the same model (characterization snapshots cover the details).
  for (const fixture of ooxmlFixtures.filter(entry => entry.kind === "xlsx" && !entry.options)) assert.doesNotThrow(() => parseXlsx(new Uint8Array(buildFixture(fixture)), fixture.file), fixture.id);
});

test("the download is always application/octet-stream and its URL is revoked after the click", async t => {
  const created = [], revoked = [], timers = new Map();
  let nextTimer = 1;
  class Element {
    constructor(tag) { this.tagName = tag; this.children = []; this.attributes = {}; this.className = ""; this.textContent = ""; this.hidden = false; this.style = { setProperty() {} }; }
    setAttribute(name, value) { this.attributes[name] = value; }
    append(...nodes) { this.children.push(...nodes); }
    replaceChildren(...nodes) { this.children = nodes; }
    remove() {}
    addEventListener() {}
    dispatchEvent() {}
    click() { this.clicked = true; }
    getRootNode() { return document; }
    get ownerDocument() { return document; }
  }
  const document = { nodeType: 9, adoptedStyleSheets: [], defaultView: { CSSStyleSheet: class { replaceSync() {} } }, createElement: tag => new Element(tag) };
  const saved = { document: globalThis.document, create: URL.createObjectURL, revoke: URL.revokeObjectURL, setTimeout: globalThis.setTimeout, clearTimeout: globalThis.clearTimeout };
  globalThis.document = document;
  URL.createObjectURL = blob => { created.push(blob); return `blob:test/${created.length}`; };
  URL.revokeObjectURL = url => revoked.push(url);
  globalThis.setTimeout = (callback, delay) => { const id = nextTimer++; timers.set(id, { callback, delay }); return id; };
  globalThis.clearTimeout = id => timers.delete(id);
  t.after(() => { globalThis.document = saved.document; URL.createObjectURL = saved.create; URL.revokeObjectURL = saved.revoke; globalThis.setTimeout = saved.setTimeout; globalThis.clearTimeout = saved.clearTimeout; });

  const container = new Element("div");
  const docx = new Uint8Array(buildFixture(ooxmlFixtures.find(fixture => fixture.id === "docx-basic")));
  // A stub renderer: the fake DOM cannot run the real one, and the download only needs the loaded bytes.
  const registry = new RendererRegistry();
  registry.register({ formats: ["docx"], render: () => ({ element: new Element("rendered") }) });
  const previewer = new FilePreviewer({ container, file: { name: "a.docx", source: new Blob([docx], { type: "text/html" }) }, zoom: false, registry });
  await previewer.open();
  const toolbar = container.children.find(child => child.className === "ufp-root").children[0];
  const download = toolbar.children.find(child => child.attributes["aria-label"] === "Download file");
  download.onclick();
  {
    assert.equal(created.at(-1).type, "application/octet-stream", "never the declared text/html");
    const timer = [...timers.values()].find(entry => entry.delay === 30_000);
    assert.ok(timer, "the URL is revoked after 30 s");
    timer.callback();
    assert.ok(revoked.includes(`blob:test/${created.length}`));
    download.onclick();
    previewer.destroy();
    assert.ok(revoked.includes(`blob:test/${created.length}`), "destroy() revokes a pending download URL");
  }
});
