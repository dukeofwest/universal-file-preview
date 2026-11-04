import test from "node:test";
import assert from "node:assert/strict";
import { allowsGeneralShrink, cellAccessibleLabel, columnPixels, columnWidthPixels, defaultColumnPixels, noticeDiagnostics, rowInfoAt } from "../dist/renderers/workbook-layout.js";

// The grid renderer's code before milestone M9b, copied verbatim, as the parity reference.
const originalColumnPixels = (width, mdw) => Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
const originalColumnWidth = (info, defaultWidth, mdw) => info?.hidden ? 0 : info?.width !== undefined ? originalColumnPixels(info.width, mdw) : defaultWidth;
const originalDefaultColumnPixels = (sheet, mdw) => {
  const stored = sheet.defaultColumnWidth;
  if (stored !== undefined && stored > 0) return originalColumnPixels(stored, mdw);
  const base = stored !== undefined && stored < 0 ? -stored : 8;
  return Math.ceil((base * mdw + 5) / 8) * 8;
};
const sheet = extra => ({ name: "S", cells: [], maxRow: 1, maxColumn: 1, ...extra });

test("column widths without widthPoints match the previous renderer exactly", () => {
  let checked = 0;
  for (let mdw = 6; mdw <= 12; mdw++) {
    for (let hundredths = 0; hundredths <= 25_500; hundredths++) {
      const width = hundredths / 100, column = { min: 1, max: 1, width };
      assert.equal(columnPixels(width, mdw), originalColumnPixels(width, mdw));
      assert.equal(columnWidthPixels(column, 64, mdw), originalColumnWidth(column, 64, mdw));
      checked++;
    }
    for (const column of [undefined, { min: 1, max: 1 }, { min: 1, max: 1, hidden: true }, { min: 1, max: 1, width: 10, hidden: true }]) {
      assert.equal(columnWidthPixels(column, 64, mdw), originalColumnWidth(column, 64, mdw));
    }
  }
  assert.equal(checked, 7 * 25_501);
});

test("absolute column widths are converted from points and hidden columns stay hidden", () => {
  assert.equal(columnWidthPixels({ min: 1, max: 1, widthPoints: 64.01 }, 64, 7), 64.01 * 4 / 3);
  assert.equal(columnWidthPixels({ min: 1, max: 1, widthPoints: 72, width: 30 }, 64, 7), 96, "widthPoints wins over width");
  assert.equal(columnWidthPixels({ min: 1, max: 1, widthPoints: 72, hidden: true }, 64, 7), 0);
  assert.equal(columnWidthPixels({ min: 1, max: 1, widthPoints: 0 }, 64, 7), 0);
});

test("default column widths match the previous renderer and honour defaultColumnWidthPoints", () => {
  for (let mdw = 6; mdw <= 12; mdw++) {
    for (const stored of [undefined, 8.43, 10, 24.5703125, 0.01, 255, -8, -10, -1]) {
      const value = sheet(stored === undefined ? {} : { defaultColumnWidth: stored });
      assert.equal(defaultColumnPixels(value, mdw), originalDefaultColumnPixels(value, mdw), `stored ${stored}, mdw ${mdw}`);
    }
  }
  assert.equal(defaultColumnPixels(sheet({ defaultColumnWidthPoints: 48, defaultColumnWidth: 20 }), 7), 64);
});

test("row properties come from rows first, then from row ranges by binary search", () => {
  const plain = sheet({ rows: { 2: { height: 30 }, 5: { hidden: true } } });
  assert.deepEqual([rowInfoAt(plain, 2), rowInfoAt(plain, 5), rowInfoAt(plain, 3)], [{ height: 30 }, { hidden: true }, undefined]);
  assert.equal(rowInfoAt(sheet({}), 1), undefined);

  // 10,000 sorted, non-overlapping ranges with gaps, checked against a linear scan.
  let seed = 42;
  const random = () => (seed = (seed * 1103515245 + 12345) >>> 0) / 4294967296;
  const ranges = [];
  for (let start = 1; ranges.length < 10_000; ) {
    start += Math.floor(random() * 5);
    const end = start + Math.floor(random() * 50);
    ranges.push({ min: start, max: end, height: ranges.length + 1, ...(random() < 0.3 ? { hidden: true } : {}) });
    start = end + 1;
  }
  const ranged = sheet({ rowRanges: ranges, rows: { [ranges[500].min]: { height: 999 } } });
  const last = ranges[ranges.length - 1].max;
  for (let check = 0; check < 50_000; check++) {
    const row = 1 + Math.floor(random() * (last + 100));
    const expected = ranged.rows[row] ?? ranges.find(range => range.min <= row && row <= range.max);
    assert.equal(rowInfoAt(ranged, row), expected, `row ${row}`);
  }
  assert.deepEqual(rowInfoAt(ranged, ranges[500].min), { height: 999 }, "rows wins over a containing range");
  assert.equal(rowInfoAt(ranged, last + 1), undefined);
  assert.equal(rowInfoAt(ranged, 0), undefined);
});

test("General shrinking is only allowed for unformatted, computed or plain values", () => {
  const cell = { ref: "A1", value: "1234.5678" };
  const file = { ...cell, data: { value: { kind: "number", lexical: "1234.5678", number: 1234.5678 }, display: { text: "1.234,57", source: "file" } } };
  const raw = { ...file, data: { ...file.data, display: { text: "1234.5678", source: "raw" } } };
  const computed = { ...file, data: { ...file.data, display: { text: "1234.5678", source: "computed" } } };
  const cases = [
    [cell, undefined, true],
    [cell, {}, true],
    [cell, { numberFormat: "#,##0.00" }, false],
    [cell, { numberFormatInfo: { category: "general", supported: true } }, true],
    [cell, { numberFormatInfo: { category: "currency", supported: true } }, false],
    [file, undefined, false],
    [raw, undefined, false],
    [computed, undefined, true],
    [computed, { numberFormat: "0.00" }, false],
    [computed, { numberFormatInfo: { category: "percentage", supported: true } }, false],
  ];
  for (const [value, style, expected] of cases) assert.equal(allowsGeneralShrink(value, style), expected, JSON.stringify([value.data?.display.source, style]));
});

test("the notice lists accuracy diagnostics first, then warnings, and never info", () => {
  const entry = (code, severity) => ({ code, severity, message: code, count: 1 });
  const workbook = { kind: "workbook", title: "t", sheets: [], diagnostics: [entry("w1", "warning"), entry("i1", "info"), entry("a1", "accuracy"), entry("w2", "warning"), entry("a2", "accuracy")] };
  assert.deepEqual(noticeDiagnostics(workbook).map(item => item.code), ["a1", "a2", "w1", "w2"]);
  assert.deepEqual(noticeDiagnostics({ kind: "workbook", title: "t", sheets: [] }), []);
  assert.deepEqual(noticeDiagnostics({ ...workbook, diagnostics: [entry("i", "info")] }), []);
});

test("accessible cell labels add the currency code and saved formula results", () => {
  assert.equal(cellAccessibleLabel("A1", "12", { ref: "A1", value: "12" }), "A1: 12", "unchanged for plain cells");
  const currency = { ref: "B2", value: "1234.5", data: { value: { kind: "currency", lexical: "1234.5", number: 1234.5, currency: "EUR" }, display: { text: "1.234,50 EUR", source: "file" }, formula: { text: "of:=1", expression: "=1", dialect: "openformula", result: "cached" } } };
  assert.equal(cellAccessibleLabel("B2", "1.234,50 EUR", currency), "B2: 1.234,50 EUR, currency EUR, saved formula result");
  const noCode = { ...currency, data: { ...currency.data, value: { kind: "currency", lexical: "1", number: 1 }, formula: { ...currency.data.formula, result: "missing" } } };
  assert.equal(cellAccessibleLabel("B2", "1", noCode), "B2: 1");
});
