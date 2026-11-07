import test from "node:test";
import assert from "node:assert/strict";
import { XmlTokenizer } from "../dist/xml/xml-tokenizer.js";
import { parseOds } from "../dist/parsers/odf/ods.js";
import { makeOds } from "./helpers/odf.mjs";

// Throughput benchmarks with absolute time limits (the specification's performance budgets). They measure
// wall-clock time, so they run alone (npm run test:performance, also in CI), not in the parallel npm test run:
// there, other test processes competing for cores, caches and memory bandwidth made the 50 MB tokenizer run take
// 8 to 12 s instead of about 2.6 s (measured), and the limits measured the machine's load instead of the parser.
// Moved here unchanged from tests/xml-tokenizer.test.mjs and tests/ods.test.mjs.

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

test("a 100,000-row by 10-column spreadsheet parses within the specification's budget", () => {
  const cells = Array.from({ length: 10 }, (_, column) => `<table:table-cell office:value-type="float" office:value="${column + 0.5}" calcext:value-type="float"><text:p>${column + 0.5}</text:p></table:table-cell>`).join("");
  const rowXml = `<table:table-row>${cells}</table:table-row>`;
  const bytes = makeOds({ body: `<table:table table:name="Big"><table:table-column table:number-columns-repeated="10"/>${rowXml.repeat(100_000)}</table:table>` });
  const before = process.memoryUsage().heapUsed, started = performance.now();
  // One million cells is above the default limit (500,000 since 0.3.0); callers can raise it through limits.maxCells.
  const workbook = parseOds(bytes, "big.ods", { limits: { maxCells: 1_000_000 } });
  const elapsed = performance.now() - started, growth = process.memoryUsage().heapUsed - before;
  assert.equal(workbook.sheets[0].cells.length, 1_000_000);
  assert.equal(workbook.sheets[0].maxRow, 100_000);
  console.log(`# parseOds: ${(bytes.length / 1048576).toFixed(1)} MiB package, 1,000,000 cells in ${elapsed.toFixed(0)} ms; heap growth ${(growth / 1048576).toFixed(0)} MiB`);
  // This test deliberately parses one million cells. Test files run in parallel, so the limit allows
  // for other suites (such as the resource-limit tests) competing for the CPU at the same time.
  assert.ok(elapsed < 30_000, `took ${elapsed.toFixed(0)} ms`);
  assert.ok(growth < 600 * 1048576, `heap grew ${(growth / 1048576).toFixed(0)} MiB`);
});
