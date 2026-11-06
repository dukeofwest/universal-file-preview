import test from "node:test";
import assert from "node:assert/strict";
import { attachViewerStyles, VIEWER_CSS } from "../dist/api/viewer-styles.js";
import { FilePreviewer } from "../dist/index.js";

// The viewer stylesheet is adopted as one constructable stylesheet per document or shadow root (no 'unsafe-inline'
// needed); a <style> element with the application's nonce is the fallback. Browser behaviour under a Content
// Security Policy is checked by scripts/browser-security.mjs; this checks the bookkeeping with a minimal fake DOM.

class FakeSheet { replaceSync(css) { this.css = css; } }
function fakeDocument({ constructable = true, Sheet = FakeSheet } = {}) {
  const document = {
    nodeType: 9, adoptedStyleSheets: [],
    defaultView: constructable ? { CSSStyleSheet: Sheet } : {},
    createElement(tag) {
      const element = { tagName: tag, attributes: {}, textContent: "", removed: false, setAttribute(name, value) { this.attributes[name] = value; }, remove() { this.removed = true; } };
      document.created.push(element);
      return element;
    },
    created: [],
  };
  if (!constructable) delete document.adoptedStyleSheets;
  return document;
}
const containerIn = (document, root = document) => ({ ownerDocument: document, getRootNode: () => root });

test("viewers in one document share one adopted stylesheet, removed with the last of them", () => {
  const document = fakeDocument(), application = new FakeSheet();
  document.adoptedStyleSheets = [application];
  const first = attachViewerStyles(containerIn(document), undefined), second = attachViewerStyles(containerIn(document), "ignored");
  assert.equal(first.element, undefined, "no <style> element where constructable stylesheets exist");
  assert.equal(document.created.length, 0);
  assert.equal(document.adoptedStyleSheets.length, 2);
  assert.equal(document.adoptedStyleSheets[0], application, "the application's sheets keep their order");
  assert.equal(document.adoptedStyleSheets[1].css, VIEWER_CSS);
  first.release(); first.release(); // releasing twice counts once
  assert.equal(document.adoptedStyleSheets.length, 2, "still used by the second viewer");
  second.release();
  assert.deepEqual(document.adoptedStyleSheets, [application]);
  // A later viewer gets a new sheet.
  const third = attachViewerStyles(containerIn(document), undefined);
  assert.equal(document.adoptedStyleSheets.length, 2);
  third.release();
});

test("the sheet is adopted again when the application replaced adoptedStyleSheets", () => {
  const document = fakeDocument();
  const first = attachViewerStyles(containerIn(document), undefined), sheet = document.adoptedStyleSheets[0];
  document.adoptedStyleSheets = [];
  const second = attachViewerStyles(containerIn(document), undefined);
  assert.deepEqual(document.adoptedStyleSheets, [sheet]);
  first.release(); second.release();
  assert.deepEqual(document.adoptedStyleSheets, []);
});

test("a container in a shadow root adopts the sheet there; a detached container uses its document", () => {
  const document = fakeDocument(), shadow = { nodeType: 11, host: {}, ownerDocument: document, adoptedStyleSheets: [] };
  const inShadow = attachViewerStyles(containerIn(document, shadow), undefined);
  assert.equal(shadow.adoptedStyleSheets.length, 1);
  assert.equal(document.adoptedStyleSheets.length, 0);
  const detached = attachViewerStyles(containerIn(document, { nodeType: 1 }), undefined);
  assert.equal(document.adoptedStyleSheets.length, 1);
  inShadow.release(); detached.release();
  assert.equal(shadow.adoptedStyleSheets.length, 0);
  assert.equal(document.adoptedStyleSheets.length, 0);
});

test("without constructable stylesheets the viewer uses a <style> element carrying styleNonce", () => {
  const document = fakeDocument({ constructable: false });
  const withNonce = attachViewerStyles(containerIn(document), "abc123"), without = attachViewerStyles(containerIn(document), undefined);
  assert.equal(withNonce.element.tagName, "style");
  assert.equal(withNonce.element.textContent, VIEWER_CSS);
  assert.deepEqual(withNonce.element.attributes, { nonce: "abc123" });
  assert.deepEqual(without.element.attributes, {});
  withNonce.release();
  assert.equal(withNonce.element.removed, true);
  assert.equal(without.element.removed, false);
  without.release();
});

test("a browser that exposes constructable stylesheets but refuses them falls back to a <style> element", () => {
  class Refusing { replaceSync() { throw new Error("not allowed"); } }
  const document = fakeDocument({ Sheet: Refusing });
  const styles = attachViewerStyles(containerIn(document), "n");
  assert.equal(styles.element?.tagName, "style");
  assert.equal(document.adoptedStyleSheets.length, 0);
  styles.release();
});

test("styleNonce must be a non-empty string", () => {
  const file = { name: "a.txt", source: new Uint8Array() };
  for (const styleNonce of ["", 5, null, {}]) assert.throws(() => new FilePreviewer({ container: {}, file, styleNonce }), TypeError, String(styleNonce));
});
