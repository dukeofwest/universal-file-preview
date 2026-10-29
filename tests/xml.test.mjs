import test from "node:test";
import assert from "node:assert/strict";
import { parseXml, elements, textContent } from "../dist/xml/xml-parser.js";

test("XML parser handles namespaces, entities, CDATA, and self-closing elements", () => {
  const root = parseXml('<?xml version="1.0"?><w:doc xmlns:w="urn:test"><w:p>A&amp;B<![CDATA[<raw>]]><w:br/></w:p></w:doc>');
  const paragraph = elements(root, "p")[0];
  assert.equal(root.localName, "doc");
  assert.equal(textContent(paragraph), "A&B<raw>");
});

test("XML parser rejects mismatched tags, DTDs, and unsupported entities", () => {
  assert.throws(() => parseXml("<a><b></a>"), /Mismatched/);
  assert.throws(() => parseXml("<!DOCTYPE a><a/>"), /DOCTYPE/);
  assert.throws(() => parseXml("<a>&external;</a>"), /unsupported ampersand/);
});
