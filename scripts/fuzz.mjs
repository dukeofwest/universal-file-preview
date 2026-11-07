// Zero-dependency mutation fuzzer for the parsers (security audit, Phase 3.3).
//
//   npm run build
//   node scripts/fuzz.mjs [--seconds 60] [--seed <number>] [--out <folder>]
//
// Mutates the test corpus (OOXML fixtures, demo samples, OpenDocument and SVG files, the denial-of-service
// corpus) with bit flips, byte overwrites, insertions, deletions, ZIP header field changes and splices, and runs
// every mutant through format detection, the matching parser and the SVG sanitizer in a worker thread. A mutant
// fails the run when the library throws anything other than a PreviewError (a bug), takes longer than the
// watchdog, or exhausts the worker's capped heap. Failing inputs are written to the output folder with their seed
// and iteration, so a run can be reproduced exactly. Exit code 1 when anything failed.
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { Worker } from "node:worker_threads";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2), option = (name, fallback) => { const index = args.indexOf(name); return index >= 0 ? args[index + 1] : fallback; };
const seconds = Number(option("--seconds", "60")), seed = Number(option("--seed", String(Date.now() % 2 ** 31))) >>> 0;
const out = resolve(option("--out", join(tmpdir(), "ufp-fuzz")));
const WATCHDOG_MS = 10_000, HEAP_MIB = 512;

/** Deterministic pseudo-random numbers (mulberry32). */
function random(state) {
  return () => { state = (state + 0x6d2b79f5) >>> 0; let t = state; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}

async function corpus() {
  const files = [];
  const { ooxmlFixtures, buildFixture } = await import(pathToFileURL(join(root, "tests", "helpers", "ooxml-fixtures.mjs")));
  for (const fixture of ooxmlFixtures) files.push({ name: fixture.file, bytes: new Uint8Array(buildFixture(fixture)) });
  const { writeDemoSamples } = await import(pathToFileURL(join(root, "scripts", "demo-samples.mjs")));
  const demo = mkdtempSync(join(tmpdir(), "ufp-fuzz-demo-"));
  try { await writeDemoSamples(demo); for (const name of readdirSync(demo)) files.push({ name, bytes: new Uint8Array(readFileSync(join(demo, name))) }); }
  finally { rmSync(demo, { recursive: true, force: true }); }
  const { makeOds, makeFods } = await import(pathToFileURL(join(root, "tests", "helpers", "odf.mjs")));
  const body = '<table:table table:name="S"><table:table-row><table:table-cell office:value-type="float" office:value="1.5"><text:p>1.5</text:p></table:table-cell><table:table-cell office:value-type="string" table:number-columns-repeated="3"><text:p>a<text:span>b</text:span></text:p></table:table-cell></table:table-row></table:table>';
  files.push({ name: "s.ods", bytes: makeOds({ body }) }, { name: "s.fods", bytes: makeFods({ body }) });
  const { svgSanitizerCases } = await import(pathToFileURL(join(root, "tests", "helpers", "security-fixtures.mjs")));
  for (const entry of svgSanitizerCases("http://attacker.invalid")) files.push({ name: "case.svg", bytes: new TextEncoder().encode(entry.svg) });
  // The small denial-of-service files (large ones would make each iteration slow).
  const { dosCorpus } = await import(pathToFileURL(join(root, "tests", "helpers", "dos-corpus.mjs")));
  for (const entry of dosCorpus()) { const bytes = entry.build(); if (bytes.byteLength < 64 * 1024) files.push({ name: entry.id, bytes }); }
  // The parts of every ZIP corpus file, for structure-aware mutations that keep the package valid (checksums
  // would otherwise reject almost every mutant before a parser sees it).
  const { ZipArchive } = await import(pathToFileURL(join(root, "dist", "archive", "zip-reader.js")));
  for (const file of files) {
    try {
      const zip = ZipArchive.open(file.bytes);
      file.entries = Object.fromEntries(zip.entries.map(entry => [entry.name, zip.read(entry.name)]));
    } catch { /* not a ZIP package */ }
  }
  return files;
}

const snippets = [
  '<c r="XFD1048576"><v>1</v></c>', '<f>SUM(A1:XFD1048576)</f>', '<mergeCell ref="A1:XFD1048576"/>', '<hyperlink ref="A1:XFD1048576" r:id="rId1"/>',
  '<col min="1" max="16384" width="1e308"/>', '<row r="1048577"/>', '<c r="A0"><v>1</v></c>', '<w:p><w:r><w:t>x</w:t></w:r></w:p>', '<a:off x="-9223372036854775808" y="1e309"/>',
  '<table:table-cell table:number-columns-repeated="16384" table:number-rows-spanned="1048576"/>', 'xmlns:r="urn:other"', "&#0;", "<![CDATA[", "<!--", "]]>",
  "<x>".repeat(300), ' r:id="rIdMissing"', '<sheet name="A" sheetId="1" r:id="rId1"/>', '<p:sldId id="256" r:id="rId1"/>',
];
const extremes = ["", "-1", "0", "NaN", "Infinity", "1e309", "-1e309", "4294967296", "99999999999999999999", "0x10", "1.5", "A1:XFD1048576", "x".repeat(70_000)];

/** A structure-aware mutation: changes the XML text of one part and writes the package again with valid checksums. */
function mutateInside(file, next, makeZip) {
  const names = Object.keys(file.entries).filter(name => /.(xml|rels)$/i.test(name));
  if (!names.length) return undefined;
  const pick = max => Math.floor(next() * max), name = names[pick(names.length)];
  let text = new TextDecoder().decode(file.entries[name]), how;
  const kind = pick(5), at = pick(text.length + 1);
  if (kind === 0) {
    // Replace an attribute value with an extreme one.
    const values = [...text.matchAll(/="[^"]*"/g)];
    if (!values.length) return undefined;
    const match = values[pick(values.length)], value = extremes[pick(extremes.length)];
    text = text.slice(0, match.index) + `="${value}"` + text.slice(match.index + match[0].length); how = `attribute = ${value.slice(0, 20)}`;
  } else if (kind === 1) { const snippet = snippets[pick(snippets.length)]; text = text.slice(0, at) + snippet + text.slice(at); how = `insert ${snippet.slice(0, 30)}`; }
  else if (kind === 2) { const length = 1 + pick(200); text = text.slice(0, at) + text.slice(at + length); how = `delete ${length} characters`; }
  else if (kind === 3) { const length = 1 + pick(400), slice = text.slice(at, at + length), times = 2 + pick(50); text = text.slice(0, at) + slice.repeat(times) + text.slice(at + length); how = `repeat ${length} characters ${times} times`; }
  else { const other = text.slice(pick(text.length), pick(text.length) + 200); text = text.slice(0, at) + other + text.slice(at); how = "copy a slice"; }
  const entries = { ...file.entries, [name]: new TextEncoder().encode(text) };
  return { bytes: new Uint8Array(makeZip(entries, { deflate: pick(2) === 1 })), how: `${name}: ${how}` };
}

/** One mutation of `input`, chosen and parameterised by `next`. */
function mutate(input, next, files) {
  const bytes = new Uint8Array(input), pick = max => Math.floor(next() * max), kind = pick(7);
  if (!bytes.length) return { bytes: new Uint8Array([pick(256)]), how: "empty" };
  if (kind === 0) { const count = 1 + pick(8); for (let i = 0; i < count; i++) { const at = pick(bytes.length); bytes[at] ^= 1 << pick(8); } return { bytes, how: `flip ${count} bits` }; }
  if (kind === 1) { const count = 1 + pick(16); for (let i = 0; i < count; i++) bytes[pick(bytes.length)] = [0, 0xff, 0x7f, 0x80, pick(256)][pick(5)]; return { bytes, how: `overwrite ${count} bytes` }; }
  if (kind === 2) { const at = pick(bytes.length), length = 1 + pick(64), inserted = Uint8Array.from({ length }, () => pick(256)); return { bytes: concat(bytes.subarray(0, at), inserted, bytes.subarray(at)), how: `insert ${length} bytes at ${at}` }; }
  if (kind === 3) { const at = pick(bytes.length), length = 1 + pick(Math.min(256, bytes.length - at)); return { bytes: concat(bytes.subarray(0, at), bytes.subarray(at + length)), how: `delete ${length} bytes at ${at}` }; }
  if (kind === 4) {
    // ZIP structure: change a 16- or 32-bit field near a local, central or end record signature.
    const signatures = [];
    for (let i = 0; i + 4 <= bytes.length; i++) if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && [0x0102, 0x0304, 0x0506, 0x0708].includes(bytes[i + 2] << 8 | bytes[i + 3])) signatures.push(i);
    if (!signatures.length) return mutate(input, next, files);
    const at = signatures[pick(signatures.length)] + 4 + 2 * pick(20), view = new DataView(bytes.buffer);
    if (at + 4 > bytes.length) return mutate(input, next, files);
    const value = [0, 1, 0xffff, 0xffffffff, 0x7fffffff, pick(2 ** 32)][pick(6)];
    if (pick(2)) view.setUint16(at, value & 0xffff, true); else view.setUint32(at, value >>> 0, true);
    return { bytes, how: `ZIP field at ${at} = ${value}` };
  }
  if (kind === 5) { const other = files[pick(files.length)].bytes, from = pick(other.length), length = 1 + pick(Math.min(4096, other.length - from)), at = pick(bytes.length); return { bytes: concat(bytes.subarray(0, at), other.subarray(from, from + length), bytes.subarray(at)), how: `splice ${length} bytes` }; }
  // Repeat a slice (length fields stay as they were).
  const at = pick(bytes.length), length = 1 + pick(Math.min(512, bytes.length - at)), times = 2 + pick(20);
  return { bytes: concat(bytes.subarray(0, at), ...Array.from({ length: times }, () => bytes.subarray(at, at + length)), bytes.subarray(at + length)), how: `repeat ${length} bytes ${times} times` };
}
function concat(...parts) { const total = parts.reduce((sum, part) => sum + part.length, 0), out = new Uint8Array(total); let at = 0; for (const part of parts) { out.set(part, at); at += part.length; } return out; }

const workerSource = `
import { parentPort } from "node:worker_threads";
const { detectFormat } = await import(${JSON.stringify(pathToFileURL(join(root, "dist", "detection", "detect-format.js")).href)});
const { parseModel } = await import(${JSON.stringify(pathToFileURL(join(root, "dist", "parsers", "parse-model.js")).href)});
const { sanitizeSvg } = await import(${JSON.stringify(pathToFileURL(join(root, "dist", "renderers", "image-url.js")).href)});
const documents = new Set(["docx", "xlsx", "pptx", "ods", "ots", "fods"]);
parentPort.on("message", ({ name, bytes }) => {
  const started = performance.now();
  let result = "handled", detail;
  try {
    const format = detectFormat(name, "", bytes);
    if (documents.has(format)) parseModel(format, bytes, name, { locale: "en-US" });
    else if (format === "svg") sanitizeSvg(bytes);
    // Also try the format the name suggests, so mutants that no longer look like their format are still parsed.
    const named = name.split(".").pop().toLowerCase();
    if (named !== format && documents.has(named)) parseModel(named, bytes, name, { locale: "en-US" });
  } catch (error) {
    if (error?.name === "PreviewError") result = "refused";
    else { result = "bug"; detail = String(error?.stack ?? error).slice(0, 2000); }
  }
  parentPort.postMessage({ result, detail, ms: Math.round(performance.now() - started) });
});
parentPort.postMessage({ ready: true });`;

function startWorker() {
  const worker = new Worker(workerSource, { eval: true, resourceLimits: { maxOldGenerationSizeMb: HEAP_MIB } });
  return new Promise((resolveReady, reject) => { worker.once("message", () => resolveReady(worker)); worker.once("error", reject); });
}
function runOne(worker, input) {
  return new Promise(resolveRun => {
    const timer = setTimeout(() => finish({ result: "hang", detail: `no answer within ${WATCHDOG_MS} ms` }), WATCHDOG_MS);
    const onMessage = message => finish(message), onError = error => finish({ result: "crash", detail: String(error?.message ?? error) }), onExit = code => finish({ result: "crash", detail: `worker exited with code ${code}` });
    function finish(outcome) { clearTimeout(timer); worker.off("message", onMessage); worker.off("error", onError); worker.off("exit", onExit); resolveRun(outcome); }
    worker.on("message", onMessage); worker.once("error", onError); worker.once("exit", onExit);
    worker.postMessage(input);
  });
}

const files = await corpus(), next = random(seed), deadline = Date.now() + seconds * 1000;
const { makeZip } = await import(pathToFileURL(join(root, "tests", "helpers", "zip.mjs")));
let worker = await startWorker(), iterations = 0;
const counts = { handled: 0, refused: 0, bug: 0, hang: 0, crash: 0 }, failures = [];
console.log(`fuzzing ${files.length} corpus files for ${seconds} s, seed ${seed}`);
while (Date.now() < deadline) {
  const source = files[Math.floor(next() * files.length)];
  const mutation = (source.entries && next() < 0.6 ? mutateInside(source, next, makeZip) : undefined) ?? mutate(source.bytes, next, files);
  const outcome = await runOne(worker, { name: source.name, bytes: mutation.bytes });
  counts[outcome.result] = (counts[outcome.result] ?? 0) + 1;
  if (outcome.result === "bug" || outcome.result === "hang" || outcome.result === "crash") {
    mkdirSync(out, { recursive: true });
    const file = join(out, `seed-${seed}-iteration-${iterations}-${outcome.result}-${source.name.replace(/[^\w.-]/g, "_")}`);
    writeFileSync(file, mutation.bytes);
    writeFileSync(`${file}.txt`, `source ${source.name}\nmutation ${mutation.how}\nresult ${outcome.result}\n${outcome.detail ?? ""}\n`);
    failures.push(`${outcome.result}: ${source.name} (${mutation.how}) -> ${file}\n  ${String(outcome.detail ?? "").split("\n").slice(0, 3).join("\n  ")}`);
    // A hang or crash leaves the worker unusable.
    if (outcome.result !== "bug") { await worker.terminate(); worker = await startWorker(); }
  }
  iterations++;
}
await worker.terminate();
console.log(`${iterations} mutants: ${counts.handled} handled, ${counts.refused} refused, ${counts.bug} bugs, ${counts.hang} hangs, ${counts.crash} crashes`);
if (failures.length) { console.log(failures.join("\n")); process.exit(1); }
