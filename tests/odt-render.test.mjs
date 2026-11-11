import test from "node:test";
import assert from "node:assert/strict";
import { OdtRenderer } from "../dist/renderers/odt-renderer.js";
import { DocxRenderer } from "../dist/renderers/ooxml-renderers.js";
import { parseModel } from "../dist/parsers/parse-model.js";
import { buildFixture, ooxmlFixtures } from "./helpers/ooxml-fixtures.mjs";
import { makeOdt } from "./helpers/odf.mjs";

// Milestone 3.5: the text renderer really builds a DOM, and the Word renderer it composes keeps
// producing exactly what it produced before. Real browsers are covered by scripts/browser-security.mjs
// and the regression script; this is the deterministic end-to-end check.

class FakeElement {
  constructor(tag) { this.tagName = tag; this.children = []; this.attributes = {}; this.className = ""; this.textContent = ""; this.title = ""; this.hidden = false; this.parent = null; this.dataset = {}; this.style = { setProperty() {} }; }
  setAttribute(name, value) { this.attributes[name] = String(value); }
  append(...nodes) { for (const node of nodes) { node.remove?.(); node.parent = this; this.children.push(node); } }
  replaceChildren(...nodes) { for (const child of this.children) child.parent = null; this.children = []; this.append(...nodes); }
  remove() { if (this.parent) this.parent.children = this.parent.children.filter(child => child !== this); this.parent = null; }
  addEventListener() {}
  get classList() {
    const names = () => this.className.split(/\s+/).filter(Boolean);
    return {
      add: (...added) => { this.className = [...new Set([...names(), ...added])].join(" "); },
      remove: (...removed) => { this.className = names().filter(name => !removed.includes(name)).join(" "); },
      contains: name => names().includes(name),
      toggle: () => {},
    };
  }
  /** The renderer reads this to decide whether a page already holds content before breaking. */
  get childElementCount() { return this.children.length; }
  insertRow() { const row = new FakeElement("tr"); this.append(row); return row; }
  insertCell() { const cell = new FakeElement("td"); this.append(cell); return cell; }
  /** Enough of a selector engine for the renderer: it looks up one class name at a time. */
  querySelector(selector) {
    const wanted = selector.replace(/^\./, "");
    const match = node => node.className.split(/\s+/).includes(wanted) || node.tagName === selector;
    const walk = node => {
      for (const child of node.children) {
        if (match(child)) return child;
        const found = walk(child);
        if (found) return found;
      }
      return null;
    };
    return walk(this);
  }
  querySelectorAll() { return []; }
  getRootNode() { return fakeDocument; }
  get ownerDocument() { return fakeDocument; }
}
const fakeDocument = { nodeType: 9, adoptedStyleSheets: [], defaultView: { CSSStyleSheet: class { replaceSync() {} } }, createElement: tag => new FakeElement(tag) };

function withDom(t) {
  const previous = globalThis.document, previousUrl = globalThis.URL;
  globalThis.document = fakeDocument;
  globalThis.URL = { ...globalThis.URL, createObjectURL: () => "blob:test", revokeObjectURL() {} };
  t.after(() => { globalThis.document = previous; globalThis.URL = previousUrl; });
}

/** The text an element shows, which the renderer puts on the innermost spans. */
const textOf = element => element.textContent || element.children.map(textOf).join("");
/** Every tag and text in a rendered tree, so the shape can be asserted without a real DOM. */
function flatten(element, out = []) {
  out.push({ tag: element.tagName, text: textOf(element), className: element.className });
  for (const child of element.children) flatten(child, out);
  return out;
}
const context = { file: { name: "letter.odt" }, format: "odt", bytes: new Uint8Array(), signal: new AbortController().signal };

test("a text document renders to a DOM with its headings, paragraphs and lists", t => {
  withDom(t);
  const styles = '<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style>';
  const body = '<text:h text:outline-level="1">Title</text:h><text:h text:outline-level="3">Sub</text:h><text:p>Body text</text:p><text:list text:style-name="L1"><text:list-item><text:p>Item one</text:p></text:list-item></text:list>';
  const model = parseModel("odt", makeOdt({ styles, body }), "letter.odt", { locale: "en-US" });
  const instance = new OdtRenderer().renderParsed(model, context);
  const nodes = flatten(instance.element);
  assert.ok(nodes.some(node => node.tag === "h1" && node.text === "Title"), "an outline level becomes its heading level");
  assert.ok(nodes.some(node => node.tag === "h3" && node.text === "Sub"));
  assert.ok(nodes.some(node => node.text === "Body text"));
  assert.ok(nodes.some(node => node.text === "Item one"));
  assert.equal(typeof instance.destroy, "function", "the instance can be destroyed like any other");
});

test("a line break wraps the line instead of starting a new page", t => {
  withDom(t);
  // text:line-break is a line break, not a page break: the page stylesheet sets white-space: pre-wrap,
  // so it belongs in the run text. Setting the model's pageBreak field instead split the paragraph
  // across pages, because splitAtPageBreaks treats such a run as the start of a new page.
  const body = "<text:p>one<text:line-break/>two<text:line-break/>three</text:p>";
  const model = parseModel("odt", makeOdt({ body }), "letter.odt", { locale: "en-US" });
  const paragraph = model.blocks.find(block => block.type === "paragraph");
  assert.equal(paragraph.value.runs.some(run => run.pageBreak), false, "no run claims a page break");
  assert.equal(paragraph.value.runs.map(run => run.text).join(""), "one\ntwo\nthree", "the breaks are newlines");

  const root = new OdtRenderer().renderParsed(model, context).element;
  const pages = flatten(root).filter(node => node.className.split(/\s+/).includes("ufp-word-page"));
  assert.equal(pages.length, 1, "two line breaks stay on one page");
  assert.equal(textOf(root).includes("one\ntwo\nthree"), true, "the text keeps its newlines");
});

test("a nested list does not take over the numbering of the list that encloses it", t => {
  withDom(t);
  // The inner list's levels were assigned over the outer list's and never restored, so the outer
  // list's later items were labelled from the inner style.
  const styles = [
    '<text:list-style style:name="OUTER"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style>',
    '<text:list-style style:name="INNER"><text:list-level-style-bullet text:level="1" text:bullet-char="-"/><text:list-level-style-bullet text:level="2" text:bullet-char="-"/></text:list-style>',
  ].join("");
  const body = '<text:list text:style-name="OUTER">'
    + "<text:list-item><text:p>first</text:p></text:list-item>"
    + '<text:list-item><text:p>second</text:p><text:list text:style-name="INNER"><text:list-item><text:p>inner</text:p></text:list-item></text:list></text:list-item>'
    + "<text:list-item><text:p>third</text:p></text:list-item>"
    + "</text:list>";
  const model = parseModel("odt", makeOdt({ styles, body }), "letter.odt", { locale: "en-US" });
  const labelled = model.blocks
    .filter(block => block.type === "paragraph" && block.value.list?.label)
    .map(block => [block.value.runs.map(run => run.text).join(""), block.value.list.label]);
  assert.deepEqual(labelled, [["first", "1."], ["second", "2."], ["inner", "-"], ["third", "3."]]);
});

test("a paragraph style's own text properties apply although a default text style exists", t => {
  withDom(t);
  // A style:default-style for the text family resolves every name in that family, which used to hide
  // the paragraph style's own text properties.
  const styles = '<style:default-style style:family="text"><style:text-properties fo:font-size="11pt"/></style:default-style>'
    + '<style:style style:name="Strong" style:family="paragraph"><style:text-properties fo:font-weight="bold" fo:color="#112233"/></style:style>';
  const model = parseModel("odt", makeOdt({ styles, body: '<text:p text:style-name="Strong">bold text</text:p>' }), "letter.odt", { locale: "en-US" });
  const run = model.blocks.find(block => block.type === "paragraph").value.runs[0];
  assert.equal(run.bold, true, "the paragraph style's weight is applied to its runs");
  assert.equal(run.color, "#112233");
});

test("styles.xml is read although content.xml declares no automatic styles", t => {
  withDom(t);
  const styles = '<style:style style:name="Strong" style:family="paragraph"><style:text-properties fo:font-weight="bold"/></style:style>';
  const bytes = makeOdt({ styles, body: '<text:p text:style-name="Strong">bold</text:p>', automatic: null });
  const model = parseModel("odt", bytes, "letter.odt", { locale: "en-US" });
  const run = model.blocks.find(block => block.type === "paragraph").value.runs[0];
  assert.equal(run.bold, true, "common styles survive a content stream without office:automatic-styles");
});

test("a table renders as a table element with its cells", t => {
  withDom(t);
  const body = '<table:table table:name="T"><table:table-row><table:table-cell><text:p>A1</text:p></table:table-cell><table:table-cell><text:p>B1</text:p></table:table-cell></table:table-row></table:table>';
  const model = parseModel("odt", makeOdt({ body }), "letter.odt", { locale: "en-US" });
  const nodes = flatten(new OdtRenderer().renderParsed(model, context).element);
  assert.ok(nodes.some(node => node.tag === "table"), "a table element is produced");
  assert.ok(nodes.some(node => node.text === "A1") && nodes.some(node => node.text === "B1"));
});

test("the notice lists what the preview cannot show, above the document", t => {
  withDom(t);
  const body = '<text:p>Before <text:note text:note-class="footnote"><text:note-body><text:p>Note</text:p></text:note-body></text:note></text:p>';
  const model = parseModel("odt", makeOdt({ body }), "letter.odt", { locale: "en-US" });
  assert.ok(model.diagnostics?.length, "the model carries the diagnostics");
  const root = new OdtRenderer().renderParsed(model, context).element;
  const nodes = flatten(root);
  assert.ok(nodes.some(node => node.className === "ufp-sheet-notice"), "the notice is rendered");
  assert.ok(nodes.some(node => /not shown/i.test(node.text)), "it says what is missing");
  assert.equal(root.children[0].className, "ufp-sheet-notice", "the notice comes first");
});

test("a document with nothing to report renders without a notice", t => {
  withDom(t);
  const model = parseModel("odt", makeOdt({ body: "<text:p>Plain</text:p>" }), "letter.odt", { locale: "en-US" });
  const root = new OdtRenderer().renderParsed(model, context).element;
  assert.equal(flatten(root).some(node => node.className === "ufp-sheet-notice"), false);
  assert.equal(root.className, "ufp-word-document", "the Word renderer's own root is returned unchanged");
});

test("the Word renderer still produces exactly what it produced for DOCX", t => {
  withDom(t);
  // The text renderer composes DocxRenderer; rendering a DOCX must be unaffected by that.
  for (const fixture of ooxmlFixtures.filter(entry => entry.id.startsWith("docx-"))) {
    const bytes = new Uint8Array(buildFixture(fixture));
    const model = parseModel("docx", bytes, `${fixture.id}.docx`, { locale: "en-US" });
    const direct = flatten(new DocxRenderer().renderParsed(model, { ...context, format: "docx", file: { name: `${fixture.id}.docx` } }).element);
    const composed = flatten(new OdtRenderer().renderParsed(model, { ...context, format: "docx", file: { name: `${fixture.id}.docx` } }).element);
    assert.deepEqual(composed, direct, `${fixture.id}: composing the renderer changes nothing`);
    assert.ok(direct.length > 1, `${fixture.id}: the fixture renders something`);
  }
});
