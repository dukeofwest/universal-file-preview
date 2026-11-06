import test from "node:test";
import assert from "node:assert/strict";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { filterButtonSvg, iconSvg, slicerIconsSvg } from "../dist/renderers/spreadsheet-icons.js";
import { svgElement } from "../dist/renderers/svg-elements.js";
import { ICON_SNAPSHOT_INDICES, ICON_SNAPSHOT_SETS } from "./helpers/icon-cases.mjs";
import { matchSnapshot } from "./helpers/snapshot.mjs";

// Spreadsheet icons and buttons are built with createElementNS and setAttribute instead of innerHTML, so the
// library has no markup sink (Trusted Types). The snapshot was written from the markup the renderer assigned to
// innerHTML before this change; the elements built now must serialize to exactly the same markup.

const SVG = "http://www.w3.org/2000/svg";

/** A minimal DOM: records namespaces, attributes in order and children, and serializes them as markup. */
function fakeDocument() {
  const created = [];
  class FakeElement {
    constructor(namespace, tag) { this.namespaceURI = namespace; this.tagName = tag; this.attributes = []; this.children = []; created.push(this); }
    setAttribute(name, value) { assert.equal(typeof value, "string"); this.attributes.push([name, value]); }
    append(...nodes) { this.children.push(...nodes); }
    get outerHTML() {
      const escape = text => text.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
      const attributes = this.attributes.map(([name, value]) => ` ${name}="${escape(value)}"`).join("");
      return this.children.length || this.tagName === "svg" ? `<${this.tagName}${attributes}>${this.children.map(child => child.outerHTML).join("")}</${this.tagName}>` : `<${this.tagName}${attributes}/>`;
    }
  }
  return { created, createElementNS: (namespace, tag) => new FakeElement(namespace, tag), createElement() { throw new Error("HTML elements are not expected"); } };
}

test("every icon, the filter button and the slicer buttons build exactly the markup the renderer used before", () => {
  const document = fakeDocument(), build = node => svgElement(node, document).outerHTML;
  const markup = { filterButton: build(filterButtonSvg), slicerFiltered: build(slicerIconsSvg(true)), slicerUnfiltered: build(slicerIconsSvg(false)) };
  for (const set of ICON_SNAPSHOT_SETS) for (const index of ICON_SNAPSHOT_INDICES) markup[`icon ${JSON.stringify(set)} ${index}`] = build(iconSvg(set, index));
  matchSnapshot("spreadsheet-icons", markup);
  assert.ok(document.created.length > 263);
  assert.ok(document.created.every(element => element.namespaceURI === SVG), "every element is created in the SVG namespace");
});

test("attribute values are set as text and never parsed as markup", () => {
  const document = fakeDocument();
  const element = svgElement(["svg", { viewBox: "0 0 1 1" }, [["path", { d: '"><script>alert(1)</script>', fill: 1.5 }]]], document);
  assert.equal(element.children.length, 1);
  assert.equal(element.children[0].children.length, 0);
  assert.deepEqual(element.children[0].attributes, [["d", '"><script>alert(1)</script>'], ["fill", "1.5"]]);
});

test("the library source has no HTML or markup sinks", () => {
  const root = fileURLToPath(new URL("../src/", import.meta.url)), sinks = /\b(innerHTML|outerHTML|insertAdjacentHTML|createContextualFragment|document\.write|new DOMParser|parseFromString|setHTMLUnsafe)\b/;
  const files = readdirSync(root, { recursive: true }).filter(file => /\.ts$/.test(file));
  assert.ok(files.length > 20);
  const found = files.filter(file => sinks.test(readFileSync(join(root, file), "utf8")));
  assert.deepEqual(found, [], "innerHTML, DOMParser and similar markup sinks are not used (Trusted Types)");
});
