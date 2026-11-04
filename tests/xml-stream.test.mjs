import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { streamXml } from "../dist/xml/xml-stream.js";
import { parseXmlNs, nsChildren, nsFirst, nsAttr } from "../dist/xml/ns-tree.js";
import { ZipArchive } from "../dist/archive/zip-reader.js";
import { PreviewError } from "../dist/errors.js";
import { writeDemoSamples } from "../scripts/demo-samples.mjs";

const NS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
};
const label = name => `{${name.uri}}${name.local}`;

/** Records every event as a readable string. `decide(element)` chooses the start action. */
function record(source, decide = () => undefined, limits) {
  const events = [];
  streamXml(source, {
    start: (element, depth) => { events.push(`start ${label(element.name)} ${depth}${element.selfClosing ? " /" : ""}`); return decide(element); },
    end: (name, depth) => events.push(`end ${label(name)} ${depth}`),
    text: (value, depth) => events.push(`text ${JSON.stringify(value)} ${depth}`),
    captured: (node, depth) => events.push(`captured ${label(node)} ${depth} children=${node.children.length}`),
  }, limits);
  return events;
}

/** Spreadsheet fragment with a configurable prefix per namespace ("" = default namespace). */
function sheet(prefix, rows = 2) {
  const q = (ns, local) => prefix[ns] ? `${prefix[ns]}:${local}` : local;
  const a = (ns, local) => `${prefix[ns] || `${ns}2`}:${local}`;
  const declarations = Object.entries(prefix).map(([ns, p]) => p ? `xmlns:${p}="${NS[ns]}"` : `xmlns="${NS[ns]}" xmlns:${ns}2="${NS[ns]}"`).join(" ");
  const cell = value => `<${q("table", "table-cell")} ${a("office", "value-type")}="float" ${a("office", "value")}="${value}"><${q("text", "p")}>${value}<![CDATA[!]]></${q("text", "p")}></${q("table", "table-cell")}>`;
  const row = index => `<${q("table", "table-row")}>${cell(index)}${cell(index + 0.5)}<${q("table", "covered-table-cell")}/></${q("table", "table-row")}>`;
  return `<${q("office", "document-content")} ${declarations}><${q("office", "body")}><${q("table", "table")} ${a("table", "name")}="S">`
    + Array.from({ length: rows }, (_, index) => row(index + 1)).join("")
    + `</${q("table", "table")}></${q("office", "body")}></${q("office", "document-content")}>`;
}
const canonical = { office: "office", table: "table", text: "text" };

test("events arrive in document order with element depths", () => {
  assert.deepEqual(record('<r xmlns="urn:r"><a x="1">t<b/><![CDATA[c]]></a><c/></r>'), [
    "start {urn:r}r 1",
    "start {urn:r}a 2",
    'text "t" 2',
    "start {urn:r}b 3 /",
    "end {urn:r}b 3",
    'text "c" 2',
    "end {urn:r}a 2",
    "start {urn:r}c 2 /",
    "end {urn:r}c 2",
    "end {urn:r}r 1",
  ]);
});

test("start events carry resolved attributes and in-scope namespaces", () => {
  let seen;
  streamXml('<r xmlns:p="urn:p" p:a="1" b="2"><p:c/></r>', { start: element => { seen ??= element; } });
  assert.deepEqual(seen.name, { uri: "", local: "r", prefix: "" });
  assert.deepEqual(seen.attributes, [{ uri: "urn:p", local: "a", qname: "p:a", value: "1" }, { uri: "", local: "b", qname: "b", value: "2" }]);
  assert.equal(seen.namespaces.get("p"), "urn:p");
});

test("captured subtrees equal the corresponding parseXmlNs nodes", () => {
  const source = sheet(canonical), captured = [];
  streamXml(source, {
    start: element => element.name.uri === NS.table && /^(table-cell|covered-table-cell)$/.test(element.name.local) ? "capture" : undefined,
    captured: (node, depth) => captured.push([node, depth]),
  });
  const table = nsFirst(nsFirst(parseXmlNs(source), NS.office, "body"), NS.table, "table");
  const expected = nsChildren(table, NS.table, "table-row").flatMap(row => row.children);
  assert.equal(captured.length, 6);
  assert.deepEqual(captured.map(([node]) => node), expected);
  // document-content (1) > body (2) > table (3) > table-row (4) > cell (5)
  assert.deepEqual(captured.map(([, depth]) => depth), [5, 5, 5, 5, 5, 5]);
  assert.deepEqual(captured[0][0].children[0].children, ["1!"], "text and CDATA are merged inside captures");
  // Capturing the root delivers the whole document tree.
  let whole;
  streamXml(source, { start: () => "capture", captured: node => { whole = node; } });
  assert.deepEqual(whole, parseXmlNs(source));
});

test("capture replaces all events of the element; skip suppresses descendants but still ends", () => {
  const source = '<r xmlns="urn:r"><keep>k</keep><skip>s<x><y/></x></skip><cap>c<x/></cap><skip/><cap/></r>';
  assert.deepEqual(record(source, element => element.name.local === "skip" ? "skip" : element.name.local === "cap" ? "capture" : undefined), [
    "start {urn:r}r 1",
    "start {urn:r}keep 2", 'text "k" 2', "end {urn:r}keep 2",
    "start {urn:r}skip 2", "end {urn:r}skip 2",
    "start {urn:r}cap 2", "captured {urn:r}cap 2 children=2",
    "start {urn:r}skip 2 /", "end {urn:r}skip 2",
    "start {urn:r}cap 2 /", "captured {urn:r}cap 2 children=0",
    "end {urn:r}r 1",
  ]);
});

test("the capture budget raises resource-limit", () => {
  const source = '<r><cell><p>1</p><p>2</p></cell></r>', capture = element => element.name.local === "cell" ? "capture" : undefined;
  assert.equal(record(source, capture, { maxCaptureElements: 3 }).filter(event => event.startsWith("captured")).length, 1);
  assert.throws(() => record(source, capture, { maxCaptureElements: 2 }), error => error instanceof PreviewError && error.code === "resource-limit" && /^Captured XML subtree exceeds the configured element limit\.$/.test(error.message));
  assert.throws(() => record(`<r><cell>${"<p/>".repeat(10_000)}</cell></r>`, capture), error => error.code === "resource-limit", "default limit is 10,000");
  for (const value of [0, -1, 1.5, Number.NaN]) assert.throws(() => streamXml("<r/>", {}, { maxCaptureElements: value }), RangeError);
});

test("canonical and renamed prefixes produce the same events", () => {
  const project = events => events.map(event => event.replace(/ children=\d+/, ""));
  const capture = element => element.name.local === "table-cell" ? "capture" : undefined;
  const expected = project(record(sheet(canonical), capture));
  assert.deepEqual(project(record(sheet({ office: "", table: "t", text: "x" }), capture)), expected);
  assert.deepEqual(project(record(sheet({ office: "table", table: "text", text: "office" }), capture)), expected);
  const values = [];
  streamXml(sheet({ office: "o", table: "t", text: "x" }), { start: element => element.name.local === "table-cell" ? "capture" : undefined, captured: node => values.push(nsAttr(node, NS.office, "value")) });
  assert.deepEqual(values, ["1", "1.5", "2", "2.5"]);
});

test("well-formedness and namespace errors are reported inside skipped and captured subtrees", () => {
  const skipAll = element => element.name.local === "s" ? "skip" : undefined, captureAll = element => element.name.local === "s" ? "capture" : undefined;
  for (const decide of [skipAll, captureAll]) {
    assert.throws(() => record("<r><s><p:x/></s></r>", decide), error => error.code === "parse-error" && /^Unbound namespace prefix p/.test(error.message));
    assert.throws(() => record("<r><s><x></y></s></r>", decide), error => error.code === "parse-error" && /^Mismatched closing tag y/.test(error.message));
    assert.throws(() => record('<r xmlns:a="u" xmlns:b="u"><s><x a:k="1" b:k="2"/></s></r>', decide), error => error.code === "parse-error" && /^Duplicate attribute \{u\}k/.test(error.message));
  }
  assert.throws(() => record("<!DOCTYPE r><r/>"), error => error.code === "parse-error" && /DOCTYPE/.test(error.message));
  assert.throws(() => record("<r>"), error => /^Unclosed element r/.test(error.message));
});

test("abort signals stop the stream", () => {
  const aborted = new AbortController(); aborted.abort();
  assert.throws(() => streamXml("<r/>", {}, { signal: aborted.signal }), error => error instanceof PreviewError && error.code === "aborted");
  const controller = new AbortController();
  let starts = 0, afterAbort = 0;
  assert.throws(() => streamXml(`<r>${"<c/>".repeat(100_000)}</r>`, {
    start: () => { starts++; if (controller.signal.aborted) afterAbort++; if (starts === 1000) controller.abort(); },
  }, { signal: controller.signal, maxElements: 200_000 }), error => error.code === "aborted");
  assert.ok(afterAbort <= 16_384, `${afterAbort} elements after abort`);
});

test("handler exceptions stop the stream and propagate", () => {
  const failure = new Error("handler failed");
  let starts = 0;
  assert.throws(() => streamXml("<r><a/><b/><c/></r>", { start: element => { starts++; if (element.name.local === "b") throw failure; } }), error => error === failure);
  assert.equal(starts, 3);
});

test("trees rebuilt from events equal parseXmlNs for demo parts and spreadsheet content", async () => {
  const sources = [sheet(canonical, 5), sheet({ office: "", table: "t", text: "x" }, 3)];
  const directory = mkdtempSync(join(tmpdir(), "ufp-m4-"));
  try {
    await writeDemoSamples(directory);
    for (const file of ["sample-report.docx", "sample-workbook.xlsx", "sample-deck.pptx"]) {
      const zip = ZipArchive.open(new Uint8Array(readFileSync(join(directory, file))));
      for (const entry of zip.entries) sources.push(new TextDecoder("utf-8", { fatal: true }).decode(zip.read(entry.name)));
    }
  } finally { rmSync(directory, { recursive: true, force: true }); }
  for (const source of sources) {
    const stack = [{ children: [] }];
    streamXml(source, {
      start: element => {
        const qname = element.name.prefix ? `${element.name.prefix}:${element.name.local}` : element.name.local;
        const node = { uri: element.name.uri, local: element.name.local, qname, attributes: element.attributes, children: [], namespaces: element.namespaces };
        stack[stack.length - 1].children.push(node); stack.push(node);
      },
      end: () => { stack.pop(); },
      text: value => { const children = stack[stack.length - 1].children; if (typeof children[children.length - 1] === "string") children[children.length - 1] += value; else children.push(value); },
    });
    assert.deepEqual(stack[0].children[0], parseXmlNs(source));
  }
  assert.ok(sources.length >= 12);
});

test("streaming 200,000 captured cells stays fast with bounded memory", () => {
  const cell = '<table:table-cell office:value-type="float" office:value="12.5"><text:p>12.5</text:p></table:table-cell>';
  const rows = 25_000, row = `<table:table-row>${cell.repeat(8)}</table:table-row>`;
  const source = `<office:document-content xmlns:office="${NS.office}" xmlns:table="${NS.table}" xmlns:text="${NS.text}"><office:body><table:table>${row.repeat(rows)}</table:table></office:body></office:document-content>`;
  const baseline = process.memoryUsage().heapUsed;
  let cells = 0, peak = baseline, total = 0;
  const started = performance.now();
  streamXml(source, {
    start: element => element.name.local === "table-cell" ? "capture" : undefined,
    captured: node => {
      total += Number(nsAttr(node, NS.office, "value"));
      if (++cells % 5000 === 0) peak = Math.max(peak, process.memoryUsage().heapUsed);
    },
  }, { maxSourceLength: 64 * 1024 * 1024, maxElements: 1_000_000, maxTokens: 5_000_000 });
  const elapsed = performance.now() - started, growth = Math.max(0, peak - baseline);
  assert.equal(cells, 200_000);
  assert.equal(total, 200_000 * 12.5);
  console.log(`# streamXml: ${(source.length / 1e6).toFixed(1)} M chars, ${cells.toLocaleString("en-US")} captured cells in ${elapsed.toFixed(0)} ms; peak heap growth ${(growth / 1e6).toFixed(1)} MB (${(growth / source.length).toFixed(2)}x source)`);
  assert.ok(elapsed < 10_000, `took ${elapsed.toFixed(0)} ms`);
  assert.ok(growth < 2 * source.length, `peak heap growth ${growth} bytes exceeds 2x the ${source.length}-character source`);
});
