import test from "node:test";
import assert from "node:assert/strict";
import { DiagnosticCollector, DIAGNOSTIC_MESSAGES } from "../dist/diagnostics.js";

// Codes and severities from docs/research/odf-implementation-specification.md §4.3.
const specifiedCodes = {
  "odf.formula.cached": "accuracy",
  "odf.formula.missing-result": "accuracy",
  "odf.formula.unknown-dialect": "info",
  "odf.display.raw": "accuracy",
  "odf.display.mismatch": "accuracy",
  "odf.value.invalid": "accuracy",
  "odf.hidden.sheets": "accuracy",
  "odf.hidden.rows": "accuracy",
  "odf.hidden.columns": "accuracy",
  "odf.hidden.filtered-rows": "accuracy",
  "odf.precision-as-shown": "accuracy",
  "odf.covered-cell-content": "warning",
  "odf.bounds.clipped": "warning",
  "odf.unsupported.drawings": "warning",
  "odf.unsupported.charts": "warning",
  "odf.unsupported.conditional-formatting": "warning",
  "odf.unsupported.pivot": "info",
  "odf.external-links": "warning",
  "odf.package.no-manifest": "info",
  "odf.package.mimetype-position": "info",
  "odf.package.signed": "info",
  "odf.package.macros": "info",
  "odf.version.newer": "info",
  "odf.settings.ignored": "info",
};

test("diagnostics catalogue matches the specified codes and severities", () => {
  assert.deepEqual(Object.keys(DIAGNOSTIC_MESSAGES).sort(), Object.keys(specifiedCodes).sort());
  for (const [code, severity] of Object.entries(specifiedCodes)) {
    assert.equal(DIAGNOSTIC_MESSAGES[code].severity, severity, code);
    for (const count of [1, 2, 1234]) {
      const message = DIAGNOSTIC_MESSAGES[code].message(count, {});
      assert.ok(message.length > 0 && !/[{}]/.test(message) && !/undefined/.test(message), `${code} renders cleanly for ${count}`);
    }
  }
});

test("collector de-duplicates by code and sheet, counts, and keeps first location", () => {
  const collector = new DiagnosticCollector();
  collector.report("odf.hidden.rows", { sheet: "Q1", cell: "A5", part: "content.xml" });
  collector.report("odf.hidden.rows", { sheet: "Q1", cell: "A9", part: "other.xml" });
  collector.report("odf.hidden.rows", { sheet: "Q1", count: 3 });
  collector.report("odf.hidden.rows", { sheet: "Q2", cell: "B2" });
  collector.report("odf.hidden.rows", { cell: "C3" });
  assert.deepEqual(collector.list(), [
    { code: "odf.hidden.rows", severity: "accuracy", message: "5 hidden rows in the used range.", count: 5, sheet: "Q1", cell: "A5", part: "content.xml" },
    { code: "odf.hidden.rows", severity: "accuracy", message: "1 hidden row in the used range.", count: 1, sheet: "Q2", cell: "B2" },
    { code: "odf.hidden.rows", severity: "accuracy", message: "1 hidden row in the used range.", count: 1, cell: "C3" },
  ]);
});

test("collector fills a missing first location from later reports", () => {
  const collector = new DiagnosticCollector();
  collector.report("odf.value.invalid", { sheet: "S" });
  collector.report("odf.value.invalid", { sheet: "S", cell: "D4", part: "content.xml" });
  const [entry] = collector.list();
  assert.deepEqual([entry.cell, entry.part, entry.count], ["D4", "content.xml", 2]);
});

test("messages are rendered with the final count and correct plurals", () => {
  const collector = new DiagnosticCollector();
  collector.report("odf.formula.missing-result");
  assert.equal(collector.list()[0].message, '1 formula cell has no saved result and is shown as "Cannot calculate".');
  collector.report("odf.formula.missing-result", { count: 1233 });
  assert.equal(collector.list()[0].message, '1,234 formula cells have no saved result and are shown as "Cannot calculate".');
  collector.add("custom.code", "warning", "{n} things; {n} again.");
  collector.add("custom.code", "warning", "ignored: the first message is kept");
  assert.deepEqual(collector.list()[1], { code: "custom.code", severity: "warning", message: "2 things; 2 again.", count: 2 });
});

test("catalogue parameters are substituted, sanitised, and length-limited", () => {
  const collector = new DiagnosticCollector();
  collector.report("odf.formula.cached", {}, { producer: "LibreOffice/24.8\n\t(Windows)" });
  collector.report("odf.version.newer", {}, { version: "1.5" });
  collector.report("odf.formula.cached", { sheet: "Other" });
  collector.report("odf.version.newer", { sheet: "S" }, { version: "x".repeat(500) });
  const [cached, version, cachedNoProducer, longVersion] = collector.list();
  assert.equal(cached.message, "Formula results are shown as last saved by LibreOffice/24.8 (Windows); they were not recalculated.");
  assert.equal(version.message, "The document uses OpenDocument version 1.5; newer features may be missing.");
  assert.equal(cachedNoProducer.message, "Formula results are shown as last saved; they were not recalculated.");
  assert.ok(longVersion.message.includes(`${"x".repeat(119)}…;`) && !longVersion.message.includes("x".repeat(120)));
});

test("collector caps distinct entries and summarises the rest", () => {
  const collector = new DiagnosticCollector(2);
  collector.add("a.one", "info", "One");
  collector.add("a.two", "info", "Two");
  collector.add("a.three", "info", "Three");
  collector.add("a.four", "info", "Four");
  collector.add("a.one", "info", "One"); // existing entries keep counting after the cap
  assert.deepEqual(collector.list(), [
    { code: "a.one", severity: "info", message: "One", count: 2 },
    { code: "a.two", severity: "info", message: "Two", count: 1 },
    { code: "diagnostics.truncated", severity: "info", message: "2 further diagnostics were omitted.", count: 2 },
  ]);
  const single = new DiagnosticCollector(1);
  single.add("a.one", "info", "One"); single.add("a.two", "info", "Two");
  assert.equal(single.list()[1].message, "1 further diagnostic was omitted.");
});

test("collector defaults to 200 distinct entries", () => {
  const collector = new DiagnosticCollector();
  for (let index = 0; index < 205; index++) collector.add("bulk.entry", "info", "Entry", { sheet: `S${index}` });
  const list = collector.list();
  assert.equal(list.length, 201);
  assert.deepEqual(list.at(-1), { code: "diagnostics.truncated", severity: "info", message: "5 further diagnostics were omitted.", count: 5 });
});

test("collector keeps insertion order and returns independent copies", () => {
  const collector = new DiagnosticCollector();
  collector.report("odf.package.signed"); collector.report("odf.hidden.sheets"); collector.report("odf.package.signed");
  const first = collector.list();
  assert.deepEqual(first.map(entry => entry.code), ["odf.package.signed", "odf.hidden.sheets"]);
  first[0].count = 99; first.length = 0;
  assert.deepEqual(collector.list().map(entry => [entry.code, entry.count]), [["odf.package.signed", 2], ["odf.hidden.sheets", 1]]);
  assert.deepEqual(new DiagnosticCollector().list(), []);
});

test("optional location fields are omitted rather than set to undefined", () => {
  const collector = new DiagnosticCollector();
  collector.report("odf.bounds.clipped");
  const [entry] = collector.list();
  assert.deepEqual(Object.keys(entry).sort(), ["code", "count", "message", "severity"]);
});

test("collector rejects invalid arguments", () => {
  for (const value of [0, -1, 1.5, Number.NaN, Infinity]) assert.throws(() => new DiagnosticCollector(value), RangeError);
  const collector = new DiagnosticCollector();
  for (const code of ["", "Upper.case", "has space", "trailing.", ".leading", "diagnostics.truncated"]) assert.throws(() => collector.add(code, "info", "x"), RangeError, code);
  assert.throws(() => collector.add("a.b", "fatal", "x"), RangeError);
  assert.throws(() => collector.add("a.b", "info", ""), RangeError);
  for (const count of [0, -2, 1.5, Number.NaN]) assert.throws(() => collector.report("odf.hidden.rows", { count }), RangeError, String(count));
  assert.throws(() => collector.report("odf.not-in-catalogue"), RangeError);
  assert.throws(() => collector.report("toString"), RangeError);
  assert.deepEqual(collector.list(), [], "rejected calls record nothing");
});
