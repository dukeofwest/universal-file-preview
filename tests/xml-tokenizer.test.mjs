import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { XmlTokenizer } from "../dist/xml/xml-tokenizer.js";
import { parseXml } from "../dist/xml/xml-parser.js";
import { ZipArchive } from "../dist/archive/zip-reader.js";
import { PreviewError } from "../dist/errors.js";
import { ooxmlFixtures } from "./helpers/ooxml-fixtures.mjs";
import { writeDemoSamples } from "../scripts/demo-samples.mjs";

const drain = (source, limits) => { const tokenizer = new XmlTokenizer(source, limits), tokens = []; for (let token; (token = tokenizer.next());) tokens.push(token); return tokens; };
const stripPosition = message => message.replace(/ \(at character \d+\)\.$/, "");
const outcome = run => {
  try { run(); return { ok: true }; }
  catch (error) { assert.ok(error instanceof PreviewError, `expected PreviewError, got ${error}`); return { ok: false, code: error.code, message: stripPosition(error.message) }; }
};
const rejects = (source, code, message, limits) => assert.throws(() => drain(source, limits), error => error instanceof PreviewError && error.code === code && message.test(error.message), JSON.stringify(source).slice(0, 80));

/** Builds the same shape as parseXml (name, attributes, children with merged adjacent text). */
function tokensToTree(tokens) {
  const root = { children: [] }, stack = [root];
  for (const token of tokens) {
    const parent = stack[stack.length - 1];
    if (token.type === "start") {
      const node = { name: token.qname, attributes: Object.fromEntries(token.attributes), children: [] };
      parent.children.push(node);
      if (!token.selfClosing) stack.push(node);
    } else if (token.type === "end") stack.pop();
    else parent.children.push(token.value);
  }
  return mergeText(root.children[0]);
}
function parserTree(node) { return mergeText({ name: node.name, attributes: { ...node.attributes }, children: node.children.map(child => typeof child === "string" ? child : parserTree(child)) }); }
function mergeText(node) {
  const children = [];
  for (const child of node.children) {
    if (typeof child === "string" && typeof children[children.length - 1] === "string") children[children.length - 1] += child;
    else children.push(typeof child === "string" ? child : mergeText(child));
  }
  return { ...node, children: children.filter(child => child !== "") };
}

// -----------------------------------------------------------------------------------------------
// Parity with parseXml (inputs from tests/xml.test.mjs and tests/characterization.test.mjs)
// -----------------------------------------------------------------------------------------------

const parityInputs = [
  '<?xml version="1.0"?><w:doc xmlns:w="urn:test"><w:p>A&amp;B<![CDATA[<raw>]]><w:br/></w:p></w:doc>',
  "<a><b></a>", "<!DOCTYPE a><a/>", "<a>&external;</a>",
  "<!doctype a><a/>", '<!ENTITY x "y"><a/>',
  '<?xml version="1.0"?><!DOCTYPE office:document-content PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "office.dtd"><a/>',
  '<a b="&lt;&#x41;&#X42;">&amp;&#67;&quot;&apos;&gt;</a>',
  "<a>&#0;</a>", "<a>&#xD800;</a>", "<a>&#x110000;</a>", "<a>&nbsp;</a>", '<a b="x & y"/>',
  '<p:a xmlns:p="urn:p" xmlns:q="urn:q" p:x="1" y="2"><p:b/><q:b>t</q:b><c/></p:a>',
  "<a> <b/> </a>", "<?xml version='1.0'?><!-- c --><a><?pi x?><!--d--></a>", "\uFEFF<a/>", "  <a/>  \n",
  "<a><![CDATA[<raw>]]></a>", '<a x="1\n2\t3"/>', "<a><![CDATA[a &amp; b]]></a>",
  "<a/><b/>", "x<a/>", "<a>", "<a></b>", "<a b='1' b='2'/>", "<a b=1/>", "<a b/>", "<a",
  "<a><!-- x</a>", "<a><![CDATA[x</a>", "<?xml version='1.0'", "",
];

/** [source, parseXml limits, tokenizer limits] */
const parityLimitCases = [
  [`<r>${"<a/>".repeat(249_999)}</r>`, undefined, undefined],
  [`<r>${"<a/>".repeat(250_000)}</r>`, undefined, undefined],
  ["<r><a/><a/></r>", { maxNodes: 2 }, { maxElements: 2 }],
  [`${"<a>".repeat(256)}${"</a>".repeat(256)}`, undefined, undefined],
  [`${"<a>".repeat(257)}${"</a>".repeat(257)}`, undefined, undefined],
  ["<a><b/></a>", { maxDepth: 1 }, { maxDepth: 1 }],
  ["x".repeat(16 * 1024 * 1024 + 1), undefined, undefined],
  ["<a>12345</a>", { maxTextLength: 10 }, { maxSourceLength: 10 }],
];

/** The one deliberate accept/reject difference: parseXml entity-decodes CDATA (recorded in M0). */
const deliberateDifferences = new Set(["<a><![CDATA[a & b]]></a>"]);

test("tokenizer accepts and rejects exactly like parseXml on the existing inputs", () => {
  const cases = [...parityInputs.map(source => [source, undefined, undefined]), ...parityLimitCases];
  for (const [source, parserLimits, tokenizerLimits] of cases) {
    const expected = outcome(() => parseXml(source, parserLimits)), actual = outcome(() => drain(source, tokenizerLimits));
    assert.deepEqual(actual, expected, `parity for ${JSON.stringify(source).slice(0, 60)}`);
  }
});

test("tokenizer keeps CDATA literal, the one deliberate difference from parseXml", () => {
  for (const source of deliberateDifferences) {
    assert.equal(outcome(() => parseXml(source)).ok, false);
    assert.deepEqual(drain(source).map(token => token.type === "text" ? token.value : token.type), ["start", "a & b", "end"]);
  }
  assert.deepEqual(drain("<a><![CDATA[a &amp; b]]></a>")[1].value, "a &amp; b");
  assert.deepEqual(drain("<a><![CDATA[]]></a>").map(token => token.type), ["start", "end"]);
});

test("tokenizer builds the same trees as parseXml for every OOXML fixture and demo part", async () => {
  const parts = [];
  for (const fixture of ooxmlFixtures) for (const [path, value] of Object.entries(fixture.entries())) if (typeof value === "string") parts.push([`${fixture.id}:${path}`, value]);
  const directory = mkdtempSync(join(tmpdir(), "ufp-m2-"));
  try {
    await writeDemoSamples(directory);
    for (const file of ["sample-report.docx", "sample-workbook.xlsx", "sample-deck.pptx"]) {
      const zip = ZipArchive.open(new Uint8Array(readFileSync(join(directory, file))));
      for (const entry of zip.entries) parts.push([`${file}:${entry.name}`, new TextDecoder("utf-8", { fatal: true }).decode(zip.read(entry.name))]);
    }
    parts.push(["sample-illustration.svg", readFileSync(join(directory, "sample-illustration.svg"), "utf8")]);
  } finally { rmSync(directory, { recursive: true, force: true }); }
  assert.ok(parts.length >= 60, `expected many parts, got ${parts.length}`);
  for (const [label, source] of parts) assert.deepEqual(tokensToTree(drain(source)), parserTree(parseXml(source)), label);
});

// -----------------------------------------------------------------------------------------------
// Token stream
// -----------------------------------------------------------------------------------------------

test("tokenizer emits start, end and text tokens with offsets, attributes and depth", () => {
  const source = `<?xml version="1.0"?>\n<r a="1" b='x'><s/>t&amp;u<![CDATA[<c>]]></r>`;
  const tokenizer = new XmlTokenizer(source), tokens = [], depths = [];
  for (let token; (token = tokenizer.next());) { tokens.push(token); depths.push(tokenizer.depth); }
  assert.deepEqual(tokens, [
    { type: "start", qname: "r", attributes: [["a", "1"], ["b", "x"]], selfClosing: false, offset: 22 },
    { type: "start", qname: "s", attributes: [], selfClosing: true, offset: 37 },
    { type: "text", value: "t&u", offset: 41 },
    { type: "text", value: "<c>", offset: 48 },
    { type: "end", qname: "r", offset: 63 },
  ]);
  assert.deepEqual(depths, [1, 1, 1, 1, 0]);
  assert.equal(tokenizer.next(), undefined);
  assert.equal(tokenizer.next(), undefined);
});

test("tokenizer handles a self-closing root, BOM, whitespace text and prefixed names", () => {
  assert.deepEqual(drain("\uFEFF<a/>"), [{ type: "start", qname: "a", attributes: [], selfClosing: true, offset: 1 }]);
  assert.deepEqual(drain("<a> <b/> </a>").map(token => token.type === "text" ? JSON.stringify(token.value) : token.type), ["start", '" "', "start", '" "', "end"]);
  const [root] = drain('<t:table xmlns:t="urn:t" t:name="Q1"/>');
  assert.deepEqual([root.qname, root.attributes], ["t:table", [["xmlns:t", "urn:t"], ["t:name", "Q1"]]]);
  assert.equal(drain('<日本 属性="値"/>')[0].qname, "日本");
  assert.equal(drain("<a>\u{1F600}&#x1F600;</a>")[1].value, "\u{1F600}\u{1F600}");
});

// -----------------------------------------------------------------------------------------------
// XML 1.0 rules that are stricter than, or differ from, parseXml
// -----------------------------------------------------------------------------------------------

test("attribute values are normalised before entity decoding (XML 1.0 §3.3.3)", () => {
  const [element] = drain('<a x="a&#10;b" y="c\nd" z="e\r\nf" w="g\th" v="i&#9;j" u="k\rl"/>');
  assert.deepEqual(Object.fromEntries(element.attributes), { x: "a\nb", y: "c d", z: "e f", w: "g h", v: "i\tj", u: "k l" });
});

test("line ends in text and CDATA are normalised to LF (XML 1.0 §2.11)", () => {
  assert.equal(drain("<a>1\r\n2\r3\n4</a>")[1].value, "1\n2\n3\n4");
  assert.equal(drain("<a><![CDATA[x\r\ny]]></a>")[1].value, "x\ny");
  assert.equal(drain("<a>&#13;&#10;</a>")[1].value, "\r\n");
});

test("tokenizer rejects malformed XML that parseXml tolerates", () => {
  rejects("<a>\u0001</a>", "parse-error", /^Invalid XML character U\+0001 \(at character 3\)\.$/);
  rejects("<a>\uD800</a>", "parse-error", /^Invalid XML character U\+D800/);
  rejects("<a>\uFFFF</a>", "parse-error", /^Invalid XML character U\+FFFF/);
  rejects("<a>&#1;</a>", "parse-error", /^Invalid XML character reference/);
  rejects("<a>&#xFFFE;</a>", "parse-error", /^Invalid XML character reference/);
  rejects('<a b="x<y"/>', "parse-error", /^Character '<' is not allowed in attribute values/);
  rejects('<a b="1"c="2"/>', "parse-error", /^Expected whitespace before attribute/);
  rejects("<a>x]]>y</a>", "parse-error", /^The sequence '\]\]>' is not allowed in text/);
  rejects("<![CDATA[x]]><a/>", "parse-error", /^CDATA section outside the root element/);
  rejects("<a/ >", "parse-error", /^Expected '>' after '\/'/);
  rejects("<!ELEMENT a ANY><a/>", "parse-error", /^Expected XML name/);
  rejects("<a>&#x;</a>", "parse-error", /^Unescaped or unsupported ampersand/);
  // Characters that are legal XML stay accepted.
  assert.equal(drain("<a>\u00A0\t\u2028\uFFFD</a>")[1].value, "\u00A0\t\u2028\uFFFD");
});

// -----------------------------------------------------------------------------------------------
// Limits, abort, and terminal errors
// -----------------------------------------------------------------------------------------------

test("tokenizer enforces its limits", () => {
  rejects("<a><b/><c/></a>", "resource-limit", /^XML token count exceeds the configured limit\.$/, { maxTokens: 3 });
  assert.equal(drain("<a><b/><c/></a>", { maxTokens: 4 }).length, 4);
  rejects('<a x="1" y="2" z="3"/>', "resource-limit", /^XML element has too many attributes\.$/, { maxAttributesPerElement: 2 });
  assert.equal(drain('<a x="1" y="2"/>', { maxAttributesPerElement: 2 }).length, 1);
  rejects("<abcd/>", "resource-limit", /^XML name exceeds the configured length limit\.$/, { maxNameLength: 3 });
  rejects('<a bcde="1"/>', "resource-limit", /^XML name exceeds/, { maxNameLength: 3 });
  rejects("<a><b/></a>", "resource-limit", /^XML element count exceeds/, { maxElements: 1 });
  rejects("<a><b/></a>", "resource-limit", /^XML nesting exceeds/, { maxDepth: 1 });
  rejects("<a/>", "resource-limit", /^XML document exceeds/, { maxSourceLength: 3 });
  assert.equal(drain(`<a>${"x".repeat(10)}</a>`, { maxSourceLength: 17 }).length, 3);
  for (const name of ["maxSourceLength", "maxTokens", "maxElements", "maxDepth", "maxAttributesPerElement", "maxNameLength"]) {
    for (const value of [0, -1, 1.5, Number.NaN, Infinity]) assert.throws(() => new XmlTokenizer("<a/>", { [name]: value }), RangeError, `${name}=${value}`);
  }
});

test("an aborted signal stops tokenizing before the first token and within 16,384 tokens", () => {
  const aborted = new AbortController(); aborted.abort(new Error("stop"));
  assert.throws(() => new XmlTokenizer("<a/>", { signal: aborted.signal }).next(), error => error instanceof PreviewError && error.code === "aborted" && error.cause?.message === "stop");
  const controller = new AbortController();
  const tokenizer = new XmlTokenizer(`<r>${"<b/>".repeat(100_000)}</r>`, { signal: controller.signal, maxElements: 200_000 });
  for (let index = 0; index < 20_000; index++) tokenizer.next();
  controller.abort();
  let extra = 0;
  assert.throws(() => { while (tokenizer.next()) extra++; }, error => error instanceof PreviewError && error.code === "aborted");
  assert.ok(extra <= 16_384, `stopped after ${extra} more tokens`);
});

test("tokenizer errors are terminal", () => {
  const tokenizer = new XmlTokenizer("<a></b>");
  assert.equal(tokenizer.next().type, "start");
  let first;
  try { tokenizer.next(); } catch (error) { first = error; }
  assert.ok(first instanceof PreviewError && /^Mismatched closing tag b \(at character 3\)\.$/.test(first.message));
  assert.throws(() => tokenizer.next(), error => error === first);
});

// -----------------------------------------------------------------------------------------------
// Performance (logged; generous threshold per the specification)
// -----------------------------------------------------------------------------------------------

test("tokenizer processes a 50 MB document within 10 seconds", () => {
  const unit = '<c a="1"><p>x</p></c>', repetitions = Math.ceil(50_000_000 / unit.length);
  const source = `<r>${unit.repeat(repetitions)}</r>`;
  const started = performance.now();
  const tokenizer = new XmlTokenizer(source, { maxSourceLength: 64 * 1024 * 1024, maxElements: 10_000_000, maxTokens: 20_000_000 });
  let count = 0;
  while (tokenizer.next()) count++;
  const elapsed = performance.now() - started;
  assert.equal(count, repetitions * 5 + 2);
  console.log(`# tokenizer: ${(source.length / 1e6).toFixed(1)} M chars, ${count.toLocaleString("en-US")} tokens in ${elapsed.toFixed(0)} ms (${(source.length / 1e6 / (elapsed / 1000)).toFixed(1)} M chars/s)`);
  assert.ok(elapsed < 10_000, `took ${elapsed.toFixed(0)} ms`);
});
