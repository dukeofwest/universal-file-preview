import test from "node:test";
import assert from "node:assert/strict";
import { sanitizeSvg } from "../dist/renderers/image-url.js";
import { parseXmlNs } from "../dist/xml/ns-tree.js";
import { evilSvg, svgSanitizerCases } from "./helpers/security-fixtures.mjs";

// The SVG sanitizer without DOMParser (0.3.0): parsed with the library's XML tokenizer, reduced to the same
// element and attribute allowlist as before, and serialized with escaping. It runs without a DOM, so it is tested
// here directly; scripts/browser-security.mjs compares it with the previous DOM sanitizer in a browser.

const ATTACKER = "http://attacker.invalid:9", SVG = "http://www.w3.org/2000/svg";
const sanitize = text => new TextDecoder().decode(sanitizeSvg(new TextEncoder().encode(text)));
const rejected = error => error?.name === "PreviewError" && error.code === "invalid-file" && error.message === "Invalid SVG image.";
const escapeRegExp = text => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
// The same check as scripts/browser-security.mjs: active content, external references and anything outside the allowlist.
const unsafe = new RegExp(`<script|\\son[a-z]+\\s*=|javascript:|${escapeRegExp(ATTACKER.replace(/^https?:\/\//, ""))}|<foreignObject|<image|<use|<style|<animate|<set[\\s>]|xlink|xml:base|xmlns(?::[\\w.-]+)?="(?!http://www\\.w3\\.org/2000/svg")|\\\\|url\\(\\s*[^#\\s]|<!--|<\\?|CDATA|xhtml|style=|class=`, "i");
const allowedElements = new Set(["svg", "g", "path", "rect", "circle", "ellipse", "line", "polyline", "polygon", "text", "tspan", "defs", "linearGradient", "radialGradient", "stop", "clipPath", "mask", "title", "desc"]);
const allowedAttributes = new Set(["viewBox", "width", "height", "x", "y", "x1", "x2", "y1", "y2", "cx", "cy", "r", "rx", "ry", "d", "points", "fill-rule", "fill-opacity", "stroke-width", "stroke-linecap", "stroke-linejoin", "stroke-dasharray", "stroke-opacity", "opacity", "transform", "font-size", "font-weight", "text-anchor", "stop-opacity", "offset", "preserveAspectRatio", "fill", "stroke", "stop-color", "clip-path", "mask", "font-family", "id"]);

/** Asserts the structural guarantees of a sanitized document. */
function assertSafe(output, label) {
  assert.doesNotMatch(output, unsafe, `${label}: active or external content in ${output}`);
  assert.ok(output.startsWith(`<svg xmlns="${SVG}"`), `${label}: the root declares the SVG namespace first`);
  assert.equal(output.match(/xmlns/g).length, 1, `${label}: exactly one namespace declaration`);
  const walk = node => {
    assert.equal(node.uri, SVG, `${label}: element outside the SVG namespace`);
    assert.ok(allowedElements.has(node.local), `${label}: element ${node.local} is not in the allowlist`);
    assert.equal(node.qname, node.local, `${label}: prefixed element ${node.qname}`);
    for (const attribute of node.attributes) assert.ok(attribute.uri === "" && allowedAttributes.has(attribute.local), `${label}: attribute ${attribute.qname} is not in the allowlist`);
    for (const child of node.children) if (typeof child !== "string") walk(child);
  };
  walk(parseXmlNs(output));
  assert.equal(sanitize(output), output, `${label}: sanitizing the output again changes it`);
}

test("every sanitizer case is rejected or reduced to the allowlist, keeping its legitimate content", () => {
  for (const entry of [...svgSanitizerCases(ATTACKER), { name: "script, onload and an external fetch", svg: evilSvg(ATTACKER), keep: ['<rect width="100" height="100" fill="red"/>'] }]) {
    if (entry.invalid) { assert.throws(() => sanitize(entry.svg), rejected, entry.name); continue; }
    const output = sanitize(entry.svg);
    assertSafe(output, entry.name);
    for (const text of entry.keep ?? []) assert.ok(output.includes(text), `${entry.name}: ${JSON.stringify(text)} missing from ${output}`);
  }
});

test("any DOCTYPE is refused, whatever its case or content", () => {
  for (const doctype of ["<!DOCTYPE svg>", "<!doctype svg>", '<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" "http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd">', '<!DOCTYPE svg [<!ENTITY a "b">]>']) {
    assert.throws(() => sanitize(`<?xml version="1.0"?>${doctype}<svg xmlns="${SVG}"/>`), rejected, doctype);
  }
  assert.throws(() => sanitize(`<svg xmlns="${SVG}"><text>&nbsp;</text></svg>`), rejected, "undeclared entity");
});

test("text and attribute values are escaped, and line ends and tabs in attribute values survive", () => {
  const output = sanitize(`<svg xmlns="${SVG}"><text font-family="&quot;A&quot; 'B'" transform="rotate(1&#10;2&#9;3&#13;4)">x &amp;&lt;&gt;" ' y&#13;z</text></svg>`);
  assert.equal(output, `<svg xmlns="${SVG}"><text font-family="&quot;A&quot; 'B'" transform="rotate(1&#10;2&#9;3&#13;4)">x &amp;&lt;&gt;" ' y&#13;z</text></svg>`);
  // Literal line ends in attribute values are normalised to spaces by XML parsing, as before.
  assert.equal(sanitize(`<svg xmlns="${SVG}"><path d="M0 0\nL1 1"/></svg>`), `<svg xmlns="${SVG}"><path d="M0 0 L1 1"/></svg>`);
});

test("characters XML does not allow are refused, as references or raw; supplementary-plane characters survive", () => {
  for (const reference of ["&#1;", "&#x1;", "&#xB;", "&#xFFFE;", "&#xFFFF;", "&#xD800;", "&#xDFFF;", "&#0;"]) {
    assert.throws(() => sanitize(`<svg xmlns="${SVG}"><text>a${reference}b</text></svg>`), rejected, `text ${reference}`);
    assert.throws(() => sanitize(`<svg xmlns="${SVG}"><text font-family="a${reference}b">t</text></svg>`), rejected, `attribute ${reference}`);
  }
  for (const raw of ["\u0001", "\u0008", "\u000B", "\u000C", "\u001F", "\uFFFE", "\uFFFF"]) {
    assert.throws(() => sanitize(`<svg xmlns="${SVG}"><text>a${raw}b</text></svg>`), rejected, `raw U+${raw.codePointAt(0).toString(16)} in text`);
    assert.throws(() => sanitize(`<svg xmlns="${SVG}"><text font-family="a${raw}b">t</text></svg>`), rejected, `raw U+${raw.codePointAt(0).toString(16)} in an attribute`);
  }
  // Bytes that are not UTF-8 (a lone surrogate cannot be encoded in UTF-8) are refused before parsing.
  assert.throws(() => sanitizeSvg(Uint8Array.from([...new TextEncoder().encode(`<svg xmlns="${SVG}"><text>`), 0xed, 0xa0, 0x80, ...new TextEncoder().encode("</text></svg>")])), rejected);
  const emoji = "\u{1F600}", output = sanitize(`<svg xmlns="${SVG}"><text>${emoji} &#x1F600; &#128512; \u{10FFFD}</text></svg>`);
  assert.equal(output, `<svg xmlns="${SVG}"><text>${emoji} ${emoji} ${emoji} \u{10FFFD}</text></svg>`);
});

test("prefixed SVG elements are written unprefixed with one default namespace declaration", () => {
  assert.equal(sanitize(`<s:svg xmlns:s="${SVG}" xmlns:x="urn:x" width="1"><s:g><s:rect width="1"/><x:rect/></s:g></s:svg>`), `<svg xmlns="${SVG}" width="1"><g><rect width="1"/></g></svg>`);
});

test("CDATA, comments and processing instructions are removed; whitespace text is kept", () => {
  assert.equal(sanitize(`<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns="${SVG}">\n  <!-- c --><?p x?><text>a<![CDATA[<b>]]>c</text>\n</svg>\n`), `<svg xmlns="${SVG}">\n  <text>ac</text>\n</svg>`);
  // A byte order mark is accepted.
  assert.equal(sanitize(`\uFEFF<svg xmlns="${SVG}"/>`), `<svg xmlns="${SVG}"/>`);
});

test("malformed or oversized documents are refused", () => {
  for (const source of ["", "<svg", `<svg xmlns="${SVG}"><rect></svg>`, `<svg xmlns="${SVG}" a="1" a="2"/>`, `<svg xmlns="${SVG}"><p:rect/></svg>`, `<svg xmlns="${SVG}"/><svg xmlns="${SVG}"/>`, `<svg xmlns="${SVG}">&unknown;</svg>`]) {
    assert.throws(() => sanitize(source), rejected, JSON.stringify(source));
  }
  const deep = `<svg xmlns="${SVG}">${"<g>".repeat(300)}${"</g>".repeat(300)}</svg>`;
  assert.throws(() => sanitize(deep), rejected, "nesting deeper than the XML depth limit");
});
