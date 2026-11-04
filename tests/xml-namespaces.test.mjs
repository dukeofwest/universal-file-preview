import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { NamespaceScope, XML_NAMESPACE } from "../dist/xml/xml-namespaces.js";
import { parseXmlNs, nsAttr, nsChildren, nsFirst, nsText, nsValuePrefix } from "../dist/xml/ns-tree.js";
import { parseXml } from "../dist/xml/xml-parser.js";
import { ZipArchive } from "../dist/archive/zip-reader.js";
import { PreviewError } from "../dist/errors.js";
import { writeDemoSamples } from "../scripts/demo-samples.mjs";

const NS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  calcext: "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
  of: "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
};
const rejects = (source, message, code = "parse-error") => assert.throws(() => parseXmlNs(source), error => error instanceof PreviewError && error.code === code && message.test(error.message), source.slice(0, 80));

/** A LibreOffice-shaped spreadsheet fragment rendered with any prefix map; "" means the default namespace. */
function spreadsheet(prefix) {
  const q = (ns, local) => prefix[ns] ? `${prefix[ns]}:${local}` : local;
  // Attributes in the default namespace must still be prefixed, so default-namespace parts declare a second prefix too.
  const a = (ns, local) => `${prefix[ns] || `${ns}2`}:${local}`;
  const declarations = Object.entries(prefix).map(([ns, p]) => p ? `xmlns:${p}="${NS[ns]}"` : `xmlns="${NS[ns]}" xmlns:${ns}2="${NS[ns]}"`).join(" ");
  return `<?xml version="1.0" encoding="UTF-8"?>`
    + `<${q("office", "document-content")} ${declarations} ${a("office", "version")}="1.3">`
    + `<${q("office", "body")}><${q("office", "spreadsheet")}><${q("table", "table")} ${a("table", "name")}="Q1">`
    + `<${q("table", "table-row")} ${a("table", "number-rows-repeated")}="2">`
    + `<${q("table", "table-cell")} ${a("office", "value-type")}="currency" ${a("calcext", "value-type")}="currency" ${a("office", "currency")}="EUR" ${a("office", "value")}="1234.5" ${a("table", "formula")}="${prefix.of || "of2"}:=SUM([.A1:.A3])">`
    + `<${q("text", "p")}>1.234,50 <![CDATA[EUR]]></${q("text", "p")}></${q("table", "table-cell")}>`
    + `</${q("table", "table-row")}></${q("table", "table")}></${q("office", "spreadsheet")}></${q("office", "body")}></${q("office", "document-content")}>`;
}

/** Prefix-independent view of a tree: names and attributes by {uri}local; formula values by their resolved prefix. */
function project(node) {
  const attributes = node.attributes.map(attribute => {
    const resolved = attribute.local === "formula" ? nsValuePrefix(node, attribute.value) : undefined;
    return [`{${attribute.uri}}${attribute.local}`, resolved ? `{${resolved.uri}}${resolved.local}` : attribute.value];
  }).sort(([a], [b]) => a < b ? -1 : 1);
  return { name: `{${node.uri}}${node.local}`, attributes, children: node.children.map(child => typeof child === "string" ? child : project(child)) };
}

test("the same document with canonical and renamed prefixes gives identical {uri, local} trees", () => {
  const canonical = parseXmlNs(spreadsheet({ office: "office", table: "table", text: "text", calcext: "calcext", of: "of" }));
  const renamed = parseXmlNs(spreadsheet({ office: "", table: "t", text: "x", calcext: "c", of: "f" }));
  const renamedAgain = parseXmlNs(spreadsheet({ office: "o", table: "office", text: "table", calcext: "text", of: "calcext" }));
  assert.deepEqual(project(renamed), project(canonical));
  assert.deepEqual(project(renamedAgain), project(canonical));
  const cell = nsFirst(nsFirst(nsFirst(nsFirst(nsFirst(renamed, NS.office, "body"), NS.office, "spreadsheet"), NS.table, "table"), NS.table, "table-row"), NS.table, "table-cell");
  assert.equal(nsAttr(cell, NS.office, "value-type"), "currency");
  assert.equal(nsAttr(cell, NS.calcext, "value-type"), "currency");
  assert.equal(nsAttr(cell, NS.office, "value"), "1234.5");
  assert.deepEqual(nsValuePrefix(cell, nsAttr(cell, NS.table, "formula")), { uri: NS.of, local: "=SUM([.A1:.A3])", prefix: "f" });
  assert.equal(nsText(cell), "1.234,50 EUR");
});

test("unbound element and attribute prefixes are parse errors", () => {
  rejects("<p:a/>", /^Unbound namespace prefix p \(at character 0\)\.$/);
  rejects('<a p:x="1"/>', /^Unbound namespace prefix p/);
  rejects('<a xmlns:p="urn:p"><q:b/></a>', /^Unbound namespace prefix q/);
  // A binding ends with its element.
  rejects('<a><b xmlns:p="urn:p"/><p:c/></a>', /^Unbound namespace prefix p/);
});

test("attributes with the same expanded name are duplicates even with different prefixes", () => {
  rejects('<a xmlns:p="urn:x" xmlns:q="urn:x" p:b="1" q:b="2"/>', /^Duplicate attribute \{urn:x\}b/);
  const ok = parseXmlNs('<a xmlns:p="urn:x" p:b="1" b="2"/>');
  assert.deepEqual(ok.attributes.map(attribute => [attribute.uri, attribute.local, attribute.qname, attribute.value]), [["urn:x", "b", "p:b", "1"], ["", "b", "b", "2"]]);
});

test("the xml prefix is pre-bound and namespace declarations are not attributes", () => {
  const root = parseXmlNs('<a xml:lang="de" xmlns="urn:d" xmlns:p="urn:p" p:x="1" y="2"/>');
  assert.deepEqual(root.attributes, [
    { uri: XML_NAMESPACE, local: "lang", qname: "xml:lang", value: "de" },
    { uri: "urn:p", local: "x", qname: "p:x", value: "1" },
    { uri: "", local: "y", qname: "y", value: "2" },
  ]);
  assert.equal(root.uri, "urn:d");
  assert.equal(parseXmlNs(`<a xmlns:xml="${XML_NAMESPACE}" xml:space="preserve"/>`).attributes[0].uri, XML_NAMESPACE);
});

test("xmlns=\"\" undeclares the default namespace for its element and descendants", () => {
  const root = parseXmlNs('<a xmlns="urn:d"><b xmlns=""><c/></b><d/></a>');
  const [b, d] = root.children, [c] = b.children;
  assert.deepEqual([root.uri, b.uri, c.uri, d.uri], ["urn:d", "", "", "urn:d"]);
  assert.equal(b.namespaces.has(""), false);
});

test("reserved prefixes, namespaces and malformed qualified names are rejected", () => {
  rejects('<a xmlns:xmlns="urn:x"/>', /^The xmlns prefix must not be declared/);
  rejects('<a xmlns:xml="urn:x"/>', /^The xml prefix must not be bound to another namespace/);
  rejects(`<a xmlns:p="${XML_NAMESPACE}"/>`, /^The XML namespace must not be bound to another prefix/);
  rejects(`<a xmlns="${XML_NAMESPACE}"/>`, /^The XML namespace must not be bound to another prefix/);
  rejects('<a xmlns:p="http://www.w3.org/2000/xmlns/"/>', /^The xmlns namespace must not be declared/);
  rejects('<a xmlns:p=""/>', /^Namespace prefix p must not be undeclared/);
  rejects('<xmlns:a xmlns:p="urn:p"/>', /^The xmlns prefix must not be used for elements/);
  rejects('<a:b:c xmlns:a="urn:a"/>', /^Invalid qualified name a:b:c/);
  rejects("<:a/>", /^Invalid qualified name :a/);
  rejects('<a xmlns:p="urn:p" p:="1"/>', /^Invalid qualified name p:/);
  rejects('<a xmlns:1p="urn:p"/>', /^Invalid namespace prefix 1p/);
});

test("nsValuePrefix resolves bound prefixes in attribute content and never throws", () => {
  const node = parseXmlNs(`<a xmlns:of="${NS.of}" xmlns:msoxl="http://schemas.microsoft.com/office/excel/formula"/>`);
  assert.deepEqual(nsValuePrefix(node, "of:=SUM([.A1])"), { uri: NS.of, local: "=SUM([.A1])", prefix: "of" });
  assert.deepEqual(nsValuePrefix(node, "msoxl:=SUM(A1:B2)"), { uri: "http://schemas.microsoft.com/office/excel/formula", local: "=SUM(A1:B2)", prefix: "msoxl" });
  const unbound = parseXmlNs("<a/>");
  assert.equal(nsValuePrefix(unbound, "of:=SUM([.A1])"), undefined);
  for (const value of ["=SUM([.A1])", "=SUM(A1:B2)", ":=x", "", "1of:=x"]) assert.equal(nsValuePrefix(node, value), undefined, value);
});

test("namespace maps are shared between elements unless declarations change them", () => {
  const root = parseXmlNs('<a xmlns:p="urn:p"><b><c/></b><d xmlns:q="urn:q"><e/></d></a>');
  const [b, d] = root.children, [c] = b.children, [e] = d.children;
  assert.equal(b.namespaces, root.namespaces);
  assert.equal(c.namespaces, root.namespaces);
  assert.notEqual(d.namespaces, root.namespaces);
  assert.equal(e.namespaces, d.namespaces);
  assert.deepEqual([...d.namespaces].sort(), [["p", "urn:p"], ["q", "urn:q"], ["xml", XML_NAMESPACE]]);
});

test("lookup helpers match by namespace URI and local name only", () => {
  const root = parseXmlNs('<r xmlns:a="urn:one" xmlns:b="urn:one" xmlns:c="urn:two"><a:x>1</a:x><b:x>2<![CDATA[3]]>4</b:x><c:x>5</c:x><x>6</x></r>');
  assert.deepEqual(nsChildren(root, "urn:one", "x").map(nsText), ["1", "234"]);
  assert.equal(nsText(nsFirst(root, "urn:two", "x")), "5");
  assert.equal(nsText(nsFirst(root, "", "x")), "6");
  assert.equal(nsFirst(root, "urn:three", "x"), undefined);
  assert.equal(nsFirst(undefined, "urn:one", "x"), undefined);
  assert.equal(nsAttr(undefined, "", "x"), undefined);
  assert.deepEqual(nsChildren(root, "urn:one", "x")[1].children, ["234"], "adjacent text and CDATA are merged");
  assert.equal(nsText(root), "123456");
});

test("tokenizer limits, security rules and errors apply to parseXmlNs", () => {
  assert.throws(() => parseXmlNs("<a><b/><c/></a>", { maxElements: 2 }), error => error.code === "resource-limit" && /element count/.test(error.message));
  assert.throws(() => parseXmlNs("<a><b/></a>", { maxDepth: 1 }), error => error.code === "resource-limit" && /nesting/.test(error.message));
  assert.throws(() => parseXmlNs("<a/>", { maxSourceLength: 3 }), error => error.code === "resource-limit");
  rejects("<!DOCTYPE a><a/>", /^DOCTYPE and entity declarations are not supported/);
  rejects("<a></b>", /^Mismatched closing tag b/);
  rejects("<a>&ent;</a>", /^Unescaped or unsupported ampersand/);
  const controller = new AbortController(); controller.abort();
  assert.throws(() => parseXmlNs("<a/>", { signal: controller.signal }), error => error instanceof PreviewError && error.code === "aborted");
});

test("NamespaceScope can be driven directly", () => {
  const scope = new NamespaceScope();
  assert.equal(scope.lookup("xml"), XML_NAMESPACE);
  scope.enter([["xmlns", "urn:d"], ["xmlns:p", "urn:p"], ["a", "1"]]);
  assert.deepEqual(scope.element("x"), { uri: "urn:d", local: "x", prefix: "" });
  assert.deepEqual(scope.element("p:x"), { uri: "urn:p", local: "x", prefix: "p" });
  assert.deepEqual(scope.attribute("x"), { uri: "", local: "x", prefix: "" });
  assert.deepEqual(scope.valuePrefix("p:=1"), { uri: "urn:p", local: "=1", prefix: "p" });
  assert.throws(() => scope.attribute("xmlns:p"), RangeError);
  assert.throws(() => scope.element("q:x", 42), error => error instanceof PreviewError && error.message === "Unbound namespace prefix q (at character 42).");
  assert.throws(() => scope.element("q:x"), error => error.message === "Unbound namespace prefix q.");
  scope.exit();
  assert.equal(scope.lookup("p"), undefined);
  assert.throws(() => scope.exit(), RangeError);
});

test("namespace-declaring OOXML demo parts parse with the same local-name structure as parseXml", async () => {
  const directory = mkdtempSync(join(tmpdir(), "ufp-m3-")), parts = [];
  try {
    await writeDemoSamples(directory);
    for (const file of ["sample-report.docx", "sample-workbook.xlsx", "sample-deck.pptx"]) {
      const zip = ZipArchive.open(new Uint8Array(readFileSync(join(directory, file))));
      for (const entry of zip.entries) parts.push([`${file}:${entry.name}`, new TextDecoder("utf-8", { fatal: true }).decode(zip.read(entry.name))]);
    }
  } finally { rmSync(directory, { recursive: true, force: true }); }
  const shape = node => [node.localName ?? node.local, node.children.filter(child => typeof child !== "string").map(shape)];
  for (const [label, source] of parts) assert.deepEqual(shape(parseXmlNs(source)), shape(parseXml(source)), label);
  assert.ok(parts.length >= 10);
});

test("a 1 MB spreadsheet fragment builds in under one second", () => {
  const row = `<table:table-row>${'<table:table-cell office:value-type="float" office:value="12.5" calcext:value-type="float"><text:p>12.5</text:p></table:table-cell>'.repeat(8)}</table:table-row>`;
  const rows = Math.ceil(1_000_000 / row.length);
  const source = `<office:document-content xmlns:office="${NS.office}" xmlns:table="${NS.table}" xmlns:text="${NS.text}" xmlns:calcext="${NS.calcext}"><table:table>${row.repeat(rows)}</table:table></office:document-content>`;
  const started = performance.now(), root = parseXmlNs(source), elapsed = performance.now() - started;
  assert.equal(nsFirst(root, NS.table, "table").children.length, rows);
  console.log(`# parseXmlNs: ${(source.length / 1e6).toFixed(2)} M chars, ${(rows * 17 + 2).toLocaleString("en-US")} elements in ${elapsed.toFixed(0)} ms`);
  assert.ok(elapsed < 1000, `took ${elapsed.toFixed(0)} ms`);
});
