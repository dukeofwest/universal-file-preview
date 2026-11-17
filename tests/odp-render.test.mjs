import test from "node:test";
import assert from "node:assert/strict";
import { OdpRenderer } from "../dist/renderers/odp-renderer.js";
import { PptxRenderer } from "../dist/renderers/ooxml-renderers.js";
import { parseModel } from "../dist/parsers/parse-model.js";
import { buildFixture, ooxmlFixtures } from "./helpers/ooxml-fixtures.mjs";
import { makeOdp } from "./helpers/odf.mjs";

// Milestone 4.2: the presentation renderer really builds a DOM, and the PowerPoint renderer it
// composes keeps producing exactly what it produced before.

class FakeElement {
  constructor(tag) { this.tagName = tag; this.children = []; this.attributes = {}; this.className = ""; this.textContent = ""; this.title = ""; this.hidden = false; this.disabled = false; this.parent = null; this.dataset = {}; this.style = { setProperty() {}, removeProperty() {} }; }
  setAttribute(name, value) { this.attributes[name] = String(value); }
  append(...nodes) { for (const node of nodes) { node.remove?.(); node.parent = this; this.children.push(node); } }
  replaceChildren(...nodes) { for (const child of this.children) child.parent = null; this.children = []; this.append(...nodes); }
  remove() { if (this.parent) this.parent.children = this.parent.children.filter(child => child !== this); this.parent = null; }
  addEventListener() {}
  removeEventListener() {}
  get parentElement() { return this.parent; }
  get classList() {
    const names = () => this.className.split(/\s+/).filter(Boolean);
    return { add: (...added) => { this.className = [...new Set([...names(), ...added])].join(" "); }, remove: () => {}, contains: name => names().includes(name), toggle: () => {} };
  }
  get childElementCount() { return this.children.length; }
  querySelector() { return null; }
  querySelectorAll() { return []; }
  getRootNode() { return fakeDocument; }
  get ownerDocument() { return fakeDocument; }
}
/** The text renderer measures runs on a canvas to fit them; without a real one every width is 0. */
class FakeCanvas extends FakeElement {
  getContext() { return { font: "", measureText: text => ({ width: text.length * 7 }) }; }
}
const fakeDocument = {
  nodeType: 9, adoptedStyleSheets: [], defaultView: { CSSStyleSheet: class { replaceSync() {} } },
  createElement: tag => (tag === "canvas" ? new FakeCanvas(tag) : new FakeElement(tag)),
  createElementNS: (_ns, tag) => new FakeElement(tag),
};

function withDom(t) {
  const previous = globalThis.document, previousUrl = globalThis.URL;
  globalThis.document = fakeDocument;
  globalThis.URL = { ...globalThis.URL, createObjectURL: () => "blob:test", revokeObjectURL() {} };
  t.after(() => { globalThis.document = previous; globalThis.URL = previousUrl; });
}

const textOf = element => element.textContent || element.children.map(textOf).join("");
function flatten(element, out = []) {
  out.push({ tag: element.tagName, text: textOf(element), className: element.className });
  for (const child of element.children) flatten(child, out);
  return out;
}
const context = { file: { name: "deck.odp" }, format: "odp", bytes: new Uint8Array(), signal: new AbortController().signal };
const frame = lines => `<draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="3cm"><draw:text-box>${lines.map(line => `<text:p>${line}</text:p>`).join("")}</draw:text-box></draw:frame>`;
const deck = body => parseModel("odp", makeOdp({ body }), "deck.odp", { locale: "en-US" });

test("a presentation renders to a DOM with its slide text", t => {
  withDom(t);
  const model = deck(`<draw:page draw:name="p1">${frame(["Title here", "Body line"])}</draw:page>`);
  const instance = new OdpRenderer().renderParsed(model, context);
  const nodes = flatten(instance.element);
  assert.ok(nodes.some(node => node.text.includes("Title here")), "the slide's text is rendered");
  assert.ok(nodes.some(node => node.text.includes("Body line")));
  assert.equal(typeof instance.destroy, "function", "the instance can be destroyed like any other");
});

test("a presentation with nothing to report renders without a notice", t => {
  withDom(t);
  const model = deck(`<draw:page draw:name="p1">${frame(["Plain"])}</draw:page>`);
  const root = new OdpRenderer().renderParsed(model, context).element;
  assert.equal(flatten(root).some(node => node.className === "ufp-sheet-notice"), false);
  assert.equal(root.className, "ufp-presentation", "the slide renderer's own root is returned unchanged");
});

test("the notice lists what the preview cannot show, above the slides", t => {
  withDom(t);
  const model = deck('<draw:page draw:name="p1"><draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="4cm"><draw:image xlink:href="Pictures/a.png"/></draw:frame></draw:page>');
  assert.ok(model.diagnostics?.length, "the model carries the diagnostics");
  const root = new OdpRenderer().renderParsed(model, context).element;
  assert.equal(root.className, "ufp-slides-wrapper");
  assert.equal(root.children[0].className, "ufp-sheet-notice", "the notice comes first");
  assert.ok(flatten(root).some(node => /not shown/i.test(node.text)), "it says what is missing");
});

test("the PowerPoint renderer still produces exactly what it produced for PPTX", t => {
  withDom(t);
  // The presentation renderer composes PptxRenderer; rendering a PPTX must be unaffected by that.
  const fixtures = ooxmlFixtures.filter(entry => entry.id.startsWith("pptx-"));
  assert.ok(fixtures.length, "there are PPTX fixtures to compare");
  for (const fixture of fixtures) {
    const bytes = new Uint8Array(buildFixture(fixture));
    const model = parseModel("pptx", bytes, `${fixture.id}.pptx`, { locale: "en-US" });
    const own = { ...context, format: "pptx", file: { name: `${fixture.id}.pptx` } };
    const direct = flatten(new PptxRenderer().renderParsed(model, own).element);
    const composed = flatten(new OdpRenderer().renderParsed(model, own).element);
    assert.deepEqual(composed, direct, `${fixture.id}: composing the renderer changes nothing`);
    assert.ok(direct.length > 1, `${fixture.id}: the fixture renders something`);
  }
});

test("the caller's signal reaches a main-thread parse", t => {
  withDom(t);
  // The text renderer once dropped context.signal, so an ODT parse could not be cancelled; the
  // presentation renderer passes it through from the start.
  const controller = new AbortController();
  controller.abort();
  const bytes = makeOdp({ body: `<draw:page draw:name="p1">${frame(["Plain"])}</draw:page>` });
  assert.throws(() => new OdpRenderer().render({ ...context, bytes, signal: controller.signal }),
    error => error?.name === "PreviewError" && error.code === "aborted");
});
