// Browser regression check for the built viewer (development tool; decision D8).
//
// Renders the demo samples and every synthetic OOXML test fixture in headless Edge or Chrome,
// driven through the Chrome DevTools Protocol with Node's built-in WebSocket. There are no
// npm dependencies. For each case it records the rendered markup, the viewer stylesheet and a
// pixel-exact screenshot, so a refactor can be compared before and after on the same machine.
//
//   npm run build
//   node scripts/browser-regression.mjs record     # baseline from the current dist/
//   (change code, then npm run build)
//   node scripts/browser-regression.mjs compare    # exit code 1 if anything differs
//
// Options: --dir <folder>      results folder (default: <os temp>/ufp-browser-regression)
//          --browser <path>    browser executable (default: UFP_BROWSER, then Edge/Chrome)
//          --locales <a,b>     browser locales (default: en-US,de-DE)
//          --worker            render in worker mode (dist/worker.js as a module worker); comparing a
//                              worker-mode run with a main-thread baseline checks that both modes render alike
//
// Screenshots depend on installed fonts and the browser build, so baselines are only
// comparable on the machine and browser that recorded them. Never commit them.
import { spawn, execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync } from "node:fs";
import { mkdir, mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { createServer } from "node:http";
import { tmpdir } from "node:os";
import { dirname, extname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { inflateSync } from "node:zlib";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const VIEWER = { width: 1200, height: 900 };
const CASE_TIMEOUT_MS = 30_000;

function usage(message) {
  if (message) console.error(`error: ${message}`);
  console.error("usage: node scripts/browser-regression.mjs <record|compare> [--dir <folder>] [--browser <path>] [--locales en-US,de-DE] [--worker]");
  process.exit(2);
}

const [mode, ...rest] = process.argv.slice(2);
if (mode !== "record" && mode !== "compare") usage(mode ? `unknown mode ${mode}` : "missing mode");
const options = { dir: join(tmpdir(), "ufp-browser-regression"), browser: process.env.UFP_BROWSER, locales: ["en-US", "de-DE"], worker: false };
for (let index = 0; index < rest.length; index += 2) {
  if (rest[index] === "--worker") { options.worker = true; index -= 1; continue; }
  const [flag, value] = [rest[index], rest[index + 1]];
  if (value === undefined) usage(`${flag} needs a value`);
  if (flag === "--dir") options.dir = resolve(value);
  else if (flag === "--browser") options.browser = value;
  else if (flag === "--locales") options.locales = value.split(",").map(locale => locale.trim()).filter(Boolean);
  else usage(`unknown option ${flag}`);
}
if (!existsSync(join(root, "dist", "index.js"))) usage("dist/ is missing; run npm run build first");
const browserPath = options.browser ?? findBrowser();
if (!browserPath) usage("no Edge or Chrome found; set UFP_BROWSER or pass --browser");

function findBrowser() {
  const candidates = process.platform === "win32"
    ? [`${process.env["ProgramFiles(x86)"]}\\Microsoft\\Edge\\Application\\msedge.exe`, `${process.env.ProgramFiles}\\Microsoft\\Edge\\Application\\msedge.exe`, `${process.env.ProgramFiles}\\Google\\Chrome\\Application\\chrome.exe`, `${process.env.LOCALAPPDATA}\\Google\\Chrome\\Application\\chrome.exe`]
    : process.platform === "darwin"
      ? ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge", "/Applications/Chromium.app/Contents/MacOS/Chromium"]
      : ["google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "microsoft-edge"].map(name => { try { return execFileSync("which", [name], { encoding: "utf8" }).trim(); } catch { return ""; } });
  return candidates.find(candidate => candidate && existsSync(candidate));
}

// -----------------------------------------------------------------------------------------------
// Test files: demo samples plus every synthetic OOXML fixture from the test suite
// -----------------------------------------------------------------------------------------------

async function collectFiles() {
  const files = new Map();
  const { writeDemoSamples } = await import(pathToFileURL(join(root, "scripts", "demo-samples.mjs")));
  const samples = await mkdtemp(join(tmpdir(), "ufp-regression-samples-"));
  try {
    await writeDemoSamples(samples);
    // Every generated demo sample, in a stable order, so new samples are covered automatically.
    for (const name of (await readdir(samples)).sort()) files.set(name, await readFile(join(samples, name)));
  } finally { await rm(samples, { recursive: true, force: true }); }
  const { ooxmlFixtures, buildFixture } = await import(pathToFileURL(join(root, "tests", "helpers", "ooxml-fixtures.mjs")));
  // Parser-option variants (such as a German-locale parse) are covered by the browser locales instead.
  for (const fixture of ooxmlFixtures) if (fixture.options === undefined) files.set(`${fixture.id}.${fixture.kind}`, buildFixture(fixture));
  return files;
}

const harness = `<!doctype html><html><head><meta charset="utf-8"><title>regression</title>
<style>html,body{margin:0;background:#fff}#viewer{width:${VIEWER.width}px;height:${VIEWER.height}px}</style></head>
<body><div id="viewer"></div><script type="module">
import { FilePreviewer } from "/dist/index.js";
const params = new URL(location.href).searchParams, name = params.get("file"), view = params.get("view");
// Worker mode: documents are parsed by dist/worker.js; a worker that fails to start fails the case (no fallback).
const worker = params.get("worker") === "1" ? { worker: { create: () => new Worker("/dist/worker.js", { type: "module" }) } } : {};
const frames = async count => { for (let i = 0; i < count; i++) await new Promise(r => requestAnimationFrame(r)); };
window.__result = (async () => {
  const container = document.getElementById("viewer");
  const viewer = new FilePreviewer({ container, file: { name, source: "/files/" + encodeURIComponent(name) }, ...worker });
  let error = null;
  try { await viewer.open(); } catch (e) { error = (e.code ?? e.name) + ": " + e.message; }
  await document.fonts.ready; await frames(2);
  const [action, argument] = (view ?? "default").split(":");
  if (action === "zoom") viewer.setZoom(Number(argument) / 100);
  if (action === "sheet") container.querySelectorAll(".ufp-sheet-tab")[Number(argument) - 1]?.click();
  if (action === "slide") for (let i = 1; i < Number(argument); i++) [...container.querySelectorAll("button")].find(b => b.textContent === "Next slide")?.click();
  // Wait until the rendered markup has stopped changing (zoom and tab changes re-layout asynchronously).
  const snapshot = () => container.innerHTML;
  await document.fonts.ready;
  for (let stable = 0, last = snapshot(), tries = 0; stable < 3 && tries < 60; tries++) {
    await frames(1); await new Promise(r => setTimeout(r, 30));
    const now = snapshot(); stable = now === last ? stable + 1 : 0; last = now;
  }
  // The viewer stylesheet, from a <style> element in the container or from a sheet the document adopted (0.3.0),
  // serialized rule by rule so both mechanisms compare equal when the CSS is the same.
  const sheetText = sheet => [...sheet.cssRules].map(rule => rule.cssText).join("\\n");
  const styles = [...[...container.querySelectorAll("style")].map(s => s.sheet), ...document.adoptedStyleSheets].filter(Boolean).map(sheetText).join("\\n");
  const clone = container.cloneNode(true); clone.querySelectorAll("style").forEach(s => s.remove());
  // Object URLs differ on every run; renderer SVG ids (ufp-<kind>-<random>) are renamed in order of appearance.
  let markup = clone.innerHTML.replace(/blob:[^"')\\s]+/g, "blob:normalised");
  const ids = [...new Set([...markup.matchAll(/id="(ufp-[a-z-]+-[a-z0-9]+)"/g)].map(m => m[1]))];
  ids.forEach((id, index) => { markup = markup.split(id).join("ufp-id-" + index); });
  // Screenshots cover the content area only, minus its bottom edge: the toolbar's button edges
  // and the viewer's rounded bottom corners (border-radius 8px) anti-alias slightly differently
  // from run to run after zoom changes. Both are still compared exactly through the markup.
  const area = container.querySelector(".ufp-content").getBoundingClientRect();
  const clip = { x: Math.round(area.left), y: Math.round(area.top), width: Math.round(area.width), height: Math.round(area.height) - 10 };
  return { error, markup, styles, clip, tabs: container.querySelectorAll(".ufp-sheet-tab").length, slides: Number((container.querySelector(".ufp-slide-nav span")?.textContent ?? "").split("/")[1] ?? 0) };
})();
</script></body></html>`;

// -----------------------------------------------------------------------------------------------
// Static server and DevTools Protocol client
// -----------------------------------------------------------------------------------------------

async function startServer(files) {
  const types = { ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".map": "application/json" };
  const server = createServer(async (request, response) => {
    const path = decodeURIComponent(new URL(request.url ?? "/", "http://localhost").pathname);
    try {
      if (path === "/harness.html") { response.writeHead(200, { "content-type": "text/html; charset=utf-8" }); response.end(harness); return; }
      if (path.startsWith("/files/") && files.has(path.slice(7))) { response.writeHead(200, { "content-type": "application/octet-stream" }); response.end(files.get(path.slice(7))); return; }
      const file = resolve(root, `.${path}`);
      if (!path.startsWith("/dist/") || !file.startsWith(join(root, "dist"))) throw new Error("not served");
      const body = await readFile(file);
      response.writeHead(200, { "content-type": types[extname(file)] ?? "application/octet-stream" }); response.end(body);
    } catch { response.writeHead(404).end(); }
  });
  await new Promise(ready => server.listen(0, "127.0.0.1", ready));
  return server;
}

async function launchBrowser(locale) {
  const profile = await mkdtemp(join(tmpdir(), "ufp-regression-profile-"));
  const child = spawn(browserPath, ["--headless=new", "--remote-debugging-port=0", `--user-data-dir=${profile}`, `--lang=${locale}`, `--accept-lang=${locale}`,
    "--no-first-run", "--no-default-browser-check", "--disable-extensions", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1", "--window-size=1280,1000", "about:blank"], { stdio: ["ignore", "ignore", "pipe"] });
  const endpoint = await new Promise((ok, fail) => {
    let log = "";
    const timer = setTimeout(() => fail(new Error("browser did not start within 30 s")), 30_000);
    child.stderr.on("data", chunk => { log += chunk; const match = /DevTools listening on (ws:\/\/\S+)/.exec(log); if (match) { clearTimeout(timer); ok(match[1]); } });
    child.on("exit", code => { clearTimeout(timer); fail(new Error(`browser exited with code ${code}`)); });
  });
  const socket = new WebSocket(endpoint);
  await new Promise((ok, fail) => { socket.addEventListener("open", ok, { once: true }); socket.addEventListener("error", fail, { once: true }); });
  let nextId = 0;
  const pending = new Map();
  socket.addEventListener("message", event => {
    const message = JSON.parse(event.data);
    const waiter = message.id !== undefined && pending.get(message.id);
    if (!waiter) return;
    pending.delete(message.id);
    if (message.error) waiter.fail(new Error(`${message.error.message} (${message.error.code})`)); else waiter.ok(message.result);
  });
  const send = (method, params = {}, sessionId) => new Promise((ok, fail) => { const id = ++nextId; pending.set(id, { ok, fail }); socket.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) })); });
  const close = async () => {
    socket.close(); child.kill();
    await new Promise(done => { if (child.exitCode !== null) done(); else child.once("exit", done); });
    await rm(profile, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }).catch(() => {});
  };
  return { send, close };
}

async function renderCase(browser, port, file, view) {
  const { targetId } = await browser.send("Target.createTarget", { url: "about:blank" });
  try {
    const { sessionId } = await browser.send("Target.attachToTarget", { targetId, flatten: true });
    await browser.send("Emulation.setDeviceMetricsOverride", { width: 1280, height: 1000, deviceScaleFactor: 1, mobile: false }, sessionId);
    await browser.send("Page.navigate", { url: `http://127.0.0.1:${port}/harness.html?file=${encodeURIComponent(file)}&view=${encodeURIComponent(view)}${options.worker ? "&worker=1" : ""}` }, sessionId);
    const deadline = Date.now() + CASE_TIMEOUT_MS;
    let result;
    while (!result) {
      if (Date.now() > deadline) throw new Error(`timed out rendering ${file} (${view})`);
      const evaluated = await browser.send("Runtime.evaluate", { expression: "window.__result", awaitPromise: true, returnByValue: true }, sessionId).catch(() => undefined);
      if (evaluated?.result?.value) result = evaluated.result.value; else await new Promise(wait => setTimeout(wait, 100));
    }
    // Retake until two consecutive screenshots match, so a capture never lands mid-repaint.
    const capture = async () => (await browser.send("Page.captureScreenshot", { format: "png", clip: { ...result.clip, scale: 1 } }, sessionId)).data;
    let previous = await capture(), png;
    for (let attempt = 0; attempt < 10 && png === undefined; attempt++) {
      await new Promise(wait => setTimeout(wait, 100));
      const next = await capture();
      if (next === previous) png = next; else previous = next;
    }
    if (png === undefined) throw new Error(`screenshot of ${file} (${view}) did not stabilise`);
    return { ...result, png: Buffer.from(png, "base64") };
  } finally { await browser.send("Target.closeTarget", { targetId }).catch(() => {}); }
}

// -----------------------------------------------------------------------------------------------
// Record and compare
// -----------------------------------------------------------------------------------------------

const sha256 = data => createHash("sha256").update(data).digest("hex");

/**
 * Decodes an 8-bit RGB or RGBA PNG (as Chromium writes screenshots) to RGB pixels. Screenshots
 * are compared by pixels, not file bytes: the same pixels can be compressed differently.
 */
function decodePng(png) {
  let at = 8, width = 0, height = 0, channels = 0;
  const idat = [];
  while (at < png.length) {
    const length = png.readUInt32BE(at), type = png.toString("ascii", at + 4, at + 8), data = png.subarray(at + 8, at + 8 + length);
    if (type === "IHDR") {
      width = data.readUInt32BE(0); height = data.readUInt32BE(4);
      if (data[8] !== 8 || (data[9] !== 2 && data[9] !== 6) || data[12] !== 0) throw new Error(`unsupported screenshot PNG (depth ${data[8]}, colour type ${data[9]}, interlace ${data[12]})`);
      channels = data[9] === 6 ? 4 : 3;
    }
    if (type === "IDAT") idat.push(data);
    at += 12 + length;
  }
  const raw = inflateSync(Buffer.concat(idat)), stride = width * channels, pixels = Buffer.alloc(width * height * channels), rgb = Buffer.alloc(width * height * 3);
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)], line = y * (stride + 1) + 1, row = y * stride;
    for (let x = 0; x < stride; x++) {
      const a = x >= channels ? pixels[row + x - channels] : 0, b = y ? pixels[row - stride + x] : 0, c = x >= channels && y ? pixels[row - stride + x - channels] : 0;
      const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
      const predictor = filter === 1 ? a : filter === 2 ? b : filter === 3 ? (a + b) >> 1 : filter === 4 ? (pa <= pb && pa <= pc ? a : pb <= pc ? b : c) : 0;
      pixels[row + x] = (raw[line + x] + predictor) & 0xff;
    }
  }
  for (let index = 0; index < width * height; index++) pixels.copy(rgb, index * 3, index * channels, index * channels + 3);
  return { width, height, rgb };
}

/** Describes how two screenshots differ: pixel count, largest channel change and bounding box. */
function describePixelDifference(beforeFile, afterFile) {
  return Promise.all([readFile(beforeFile), readFile(afterFile)]).then(([before, after]) => {
    const a = decodePng(before), b = decodePng(after);
    if (a.width !== b.width || a.height !== b.height) return `size ${a.width}x${a.height} became ${b.width}x${b.height}`;
    let count = 0, delta = 0, left = Infinity, top = Infinity, right = -1, bottom = -1;
    for (let index = 0; index < a.width * a.height; index++) {
      const d = Math.max(Math.abs(a.rgb[index * 3] - b.rgb[index * 3]), Math.abs(a.rgb[index * 3 + 1] - b.rgb[index * 3 + 1]), Math.abs(a.rgb[index * 3 + 2] - b.rgb[index * 3 + 2]));
      if (!d) continue;
      const x = index % a.width, y = Math.floor(index / a.width);
      count++; delta = Math.max(delta, d); left = Math.min(left, x); right = Math.max(right, x); top = Math.min(top, y); bottom = Math.max(bottom, y);
    }
    return `${count} pixels differ (largest channel change ${delta}) within x ${left}-${right}, y ${top}-${bottom}`;
  });
}
const safe = text => text.replace(/[^A-Za-z0-9._-]+/g, "_");

async function run() {
  const files = await collectFiles(), server = await startServer(files), port = server.address().port;
  const output = join(options.dir, mode === "record" ? "baseline" : "current");
  await rm(output, { recursive: true, force: true });
  const manifest = { browser: "", locales: options.locales, viewer: VIEWER, mode: options.worker ? "worker" : "main-thread", commit: gitCommit(), created: new Date().toISOString(), cases: {} };
  try {
    for (const locale of options.locales) {
      const browser = await launchBrowser(locale);
      try {
        manifest.browser = (await browser.send("Browser.getVersion")).product;
        for (const file of files.keys()) {
          const views = ["default", "zoom:150"];
          for (let index = 0; index < views.length; index++) {
            const view = views[index], result = await renderCase(browser, port, file, view);
            if (view === "default") {
              for (let tab = 2; tab <= result.tabs; tab++) views.push(`sheet:${tab}`);
              for (let slide = 2; slide <= result.slides; slide++) views.push(`slide:${slide}`);
            }
            const key = `${locale}/${file}/${view}`, base = join(output, safe(locale), safe(file), safe(view));
            await mkdir(dirname(base), { recursive: true });
            await Promise.all([writeFile(`${base}.html`, result.markup), writeFile(`${base}.css`, result.styles), writeFile(`${base}.png`, result.png)]);
            manifest.cases[key] = { error: result.error, markup: sha256(result.markup), styles: sha256(result.styles), screenshot: sha256(decodePng(result.png).rgb), files: base };
          }
        }
      } finally { await browser.close(); }
    }
  } finally { server.close(); }
  await writeFile(join(output, "manifest.json"), `${JSON.stringify(manifest, null, 2)}\n`);
  const count = Object.keys(manifest.cases).length;
  if (mode === "record") { console.log(`Recorded ${count} cases (${manifest.mode} mode) with ${manifest.browser} at commit ${manifest.commit} into ${output}`); return 0; }
  return compare(manifest, output, count);
}

async function compare(current, output, count) {
  const baselinePath = join(options.dir, "baseline", "manifest.json");
  if (!existsSync(baselinePath)) { console.error(`No baseline at ${baselinePath}; run "record" first.`); return 2; }
  const baseline = JSON.parse(await readFile(baselinePath, "utf8"));
  if (baseline.browser !== current.browser) console.warn(`warning: baseline used ${baseline.browser}, this run used ${current.browser}; screenshots may differ for that reason alone.`);
  const problems = [];
  for (const key of new Set([...Object.keys(baseline.cases), ...Object.keys(current.cases)])) {
    const before = baseline.cases[key], after = current.cases[key];
    if (!before || !after) { problems.push(`${key}: ${before ? "missing in this run" : "new case (not in baseline)"}`); continue; }
    const changed = ["error", "markup", "styles", "screenshot"].filter(field => before[field] !== after[field]);
    if (!changed.length) continue;
    const pixels = changed.includes("screenshot") ? `\n    screenshot: ${await describePixelDifference(`${before.files}.png`, `${after.files}.png`)}` : "";
    problems.push(`${key}: ${changed.join(", ")} differ${pixels}\n    baseline: ${before.files}.*\n    current:  ${after.files}.*`);
  }
  if (!problems.length) { console.log(`All ${count} cases (${current.mode} mode) identical to the baseline recorded at commit ${baseline.commit} (${baseline.mode ?? "main-thread"} mode).`); return 0; }
  console.error(`${problems.length} of ${count} cases differ from the baseline recorded at commit ${baseline.commit}:\n${problems.map(problem => `  ${problem}`).join("\n")}`);
  console.error(`Results: ${output}`);
  return 1;
}

function gitCommit() {
  try {
    const commit = execFileSync("git", ["rev-parse", "--short", "HEAD"], { cwd: root, encoding: "utf8" }).trim();
    const dirty = execFileSync("git", ["status", "--porcelain", "--", "src"], { cwd: root, encoding: "utf8" }).trim();
    return dirty ? `${commit} (with uncommitted src changes)` : commit;
  } catch { return "unknown"; }
}

process.exitCode = await run().catch(error => { console.error(error instanceof Error ? error.stack : error); return 2; });
