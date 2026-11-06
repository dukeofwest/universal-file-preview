// Browser security regression check for the built viewer: embedded SVG, font-name CSS injection
// and viewer containment.
//
// Renders the malicious fixtures from tests/helpers/security-fixtures.mjs in headless Edge or
// Chrome, driven through the Chrome DevTools Protocol with Node's built-in WebSocket (no npm
// dependencies, like scripts/browser-regression.mjs). Two local servers are used: the host serves
// the viewer, and an "attacker" server records every request it receives, from any tab. Nothing
// leaves 127.0.0.1.
//
//   npm run build
//   node scripts/browser-security.mjs [--browser <path>] [--report <file.json>]
//
// Exit code 1 if any check fails. A machine-readable report (browser version, environment, every
// check and its result) is written to --report, by default <os temp>/ufp-browser-security/report.json.
//
// Checks:
// - positive controls: an unsanitized SVG blob runs script when opened, and injected cssText
//   loads the attacker URL, so the harness can detect both attacks;
// - for every fixture: each image blob has an allowlisted type, SVG blobs contain no active or
//   external content and run no script when opened as a document, the expected number of
//   sanitized SVG images is shown, no element is position:fixed or refers to the attacker;
// - the SVG sanitizer table (namespace confusion, CSS-escaped url() values, entities): the
//   output is clean, legitimate content is kept, invalid input is rejected, and the sanitized
//   images load in <img> and run no script when opened as a document;
// - containment, the XLSX input-message tooltip at all four viewer edges (identical with and
//   without containment), frozen panes, scrolling, and rendering under the recommended CSP;
// - no uncaught exception in any page or tab, and no request reached the attacker server after
//   the controls, counted once all tabs have closed; both servers still answer at the end.
import { spawn, execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { createServer } from "node:http";
import { tmpdir } from "node:os";
import { dirname, extname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const STEP_TIMEOUT_MS = 60_000, TOTAL_TIMEOUT_MS = 10 * 60_000;
const options = { browser: process.env.UFP_BROWSER, report: join(tmpdir(), "ufp-browser-security", "report.json") };
const args = process.argv.slice(2);
for (let index = 0; index < args.length; index += 2) {
  if (args[index] === "--browser" && args[index + 1]) options.browser = args[index + 1];
  else if (args[index] === "--report" && args[index + 1]) options.report = resolve(args[index + 1]);
  else { console.error("usage: node scripts/browser-security.mjs [--browser <path>] [--report <file.json>]"); process.exit(2); }
}
if (!existsSync(join(root, "dist", "index.js"))) { console.error("error: dist/ is missing; run npm run build first"); process.exit(2); }
const browserPath = options.browser ?? findBrowser();
if (!browserPath) { console.error("error: no Edge or Chrome found; set UFP_BROWSER or pass --browser"); process.exit(2); }

function findBrowser() {
  const candidates = process.platform === "win32"
    ? [`${process.env["ProgramFiles(x86)"]}\\Microsoft\\Edge\\Application\\msedge.exe`, `${process.env.ProgramFiles}\\Microsoft\\Edge\\Application\\msedge.exe`, `${process.env.ProgramFiles}\\Google\\Chrome\\Application\\chrome.exe`, `${process.env.LOCALAPPDATA}\\Google\\Chrome\\Application\\chrome.exe`]
    : process.platform === "darwin"
      ? ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge", "/Applications/Chromium.app/Contents/MacOS/Chromium"]
      : ["google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "microsoft-edge"].map(name => { try { return execFileSync("which", [name], { encoding: "utf8" }).trim(); } catch { return ""; } });
  return candidates.find(candidate => candidate && existsSync(candidate));
}
function gitState() {
  try {
    const commit = execFileSync("git", ["rev-parse", "HEAD"], { cwd: root, encoding: "utf8" }).trim();
    const dirty = execFileSync("git", ["status", "--porcelain", "--untracked-files=no"], { cwd: root, encoding: "utf8" }).trim() !== "";
    return { commit, dirty };
  } catch { return { commit: "unknown", dirty: true }; }
}

// -----------------------------------------------------------------------------------------------
// Servers and fixtures
// -----------------------------------------------------------------------------------------------

const hits = [];
const attacker = createServer((request, response) => { hits.push(request.url); response.writeHead(200, { "access-control-allow-origin": "*", "content-type": "image/gif" }); response.end(); });
await new Promise(done => attacker.listen(0, "127.0.0.1", done));
const ATTACKER = `http://127.0.0.1:${attacker.address().port}`, ATTACKER_HOST = ATTACKER.slice("http://".length);
const LIVENESS = ["/liveness-node", "/liveness-page"];

const { securityFixtures, svgSanitizerCases, evilSvg } = await import(pathToFileURL(join(root, "tests", "helpers", "security-fixtures.mjs")));
const { makeZip } = await import(pathToFileURL(join(root, "tests", "helpers", "zip.mjs")));
const { ooxmlFixtures, buildFixture } = await import(pathToFileURL(join(root, "tests", "helpers", "ooxml-fixtures.mjs")));
const fixtures = securityFixtures(ATTACKER), svgCases = svgSanitizerCases(ATTACKER);

const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const workbook = (sheet, styles) => makeZip({
  "xl/workbook.xml": `<workbook ${S}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  "xl/_rels/workbook.xml.rels": `<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/>${styles ? '<Relationship Id="rId2" Type="x/styles" Target="styles.xml"/>' : ""}</Relationships>`,
  ...(styles ? { "xl/styles.xml": styles } : {}),
  "xl/worksheets/sheet1.xml": `<worksheet ${S}>${sheet}</worksheet>`,
});
const column = index => String.fromCharCode(65 + Math.floor(index / 26) - 1).replace("@", "") + String.fromCharCode(65 + index % 26);
/** Fixtures the viewer must refuse; every other file must render. */
const refusedFiles = new Set(["shared-attack.xlsx"]);
/** `cells` cells (100 per row) all referring to one shared string of `length` characters. */
const sharedStringBook = (cells, length) => makeZip({
  "xl/workbook.xml": `<workbook ${S}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  "xl/_rels/workbook.xml.rels": '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
  "xl/sharedStrings.xml": `<sst ${S}><si><t>${"y".repeat(length)}</t></si></sst>`,
  "xl/worksheets/sheet1.xml": `<worksheet ${S}><sheetData>${Array.from({ length: Math.ceil(cells / 100) }, (_, row) => `<row r="${row + 1}">${Array.from({ length: Math.min(100, cells - row * 100) }, (_, col) => `<c r="${column(col)}${row + 1}" t="s"><v>0</v></c>`).join("")}</row>`).join("")}</sheetData></worksheet>`,
}, { deflate: true });
// A grid with frozen first row and column and an input message on every cell (tooltip and frozen-pane checks).
const bigRows = Array.from({ length: 200 }, (_, row) => `<row r="${row + 1}">${Array.from({ length: 40 }, (_, col) => `<c r="${column(col)}${row + 1}" t="inlineStr"><is><t>r${row + 1}c${col + 1}</t></is></c>`).join("")}</row>`).join("");
const viewFiles = {
  "tooltip.xlsx": workbook(`<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetData>${bigRows}</sheetData><dataValidations count="1"><dataValidation type="none" allowBlank="1" showInputMessage="1" promptTitle="Input" prompt="An input message long enough to need several lines of text in the box." sqref="A1:AN200"/></dataValidations>`),
  "scroll.xlsx": workbook(`<sheetData>${Array.from({ length: 400 }, (_, index) => `<row r="${index + 1}"><c r="A${index + 1}"><v>${index + 1}</v></c></row>`).join("")}</sheetData>`),
  // XLSX text limits: the 0.2.1 attack (100,000 cells sharing one 1,000,000-character string) must be refused; the
  // largest accepted case (1,023 cells sharing one 32,767-character string, at the model text budget: each cell
  // carries the string as its value and its display text) must render.
  "shared-attack.xlsx": sharedStringBook(100_000, 1_000_000), // refused (refusedFiles)
  "shared-at-budget.xlsx": sharedStringBook(1023, 32_767),
  // Renderer work limits: thousands of overlapping full-size merges and whole-sheet column ranges.
  "limits.xlsx": workbook(`<cols>${'<col min="1" max="16384" width="9"/>'.repeat(20_000)}</cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>limits</t></is></c></row></sheetData><mergeCells count="5000">${'<mergeCell ref="A1:SF5001"/>'.repeat(5000)}</mergeCells>`),
  // A sheet image whose blob URL must be released when the grid re-renders.
  "image.xlsx": makeZip({
    "xl/workbook.xml": `<workbook ${S}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    "xl/worksheets/sheet1.xml": `<worksheet ${S}><sheetData>${Array.from({ length: 300 }, (_, index) => `<row r="${index + 1}"><c r="A${index + 1}"><v>${index}</v></c></row>`).join("")}</sheetData><drawing r:id="rIdDraw"/></worksheet>`,
    "xl/worksheets/_rels/sheet1.xml.rels": '<Relationships><Relationship Id="rIdDraw" Type="x/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
    "xl/drawings/drawing1.xml": '<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="urn:r"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:blipFill><a:blip r:embed="rIdImg"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor></xdr:wsDr>',
    "xl/drawings/_rels/drawing1.xml.rels": '<Relationships><Relationship Id="rIdImg" Type="x/image" Target="../media/image1.png"/></Relationships>',
    "xl/media/image1.png": Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==", "base64"),
  }),
  "scroll.docx": makeZip({ "word/document.xml": `<w:document xmlns:w="urn:w"><w:body>${Array.from({ length: 200 }, (_, index) => `<w:p><w:r><w:t>Paragraph ${index + 1}</w:t></w:r></w:p>`).join("")}</w:body></w:document>` }),
};
// The policy recommended in docs/security.md ("Deployment for untrusted files"); keep the two in sync.
// The recommended policy (docs/security.md): no 'unsafe-inline' for styles. The nonce page allows one nonce, for the
// fallback <style> element of browsers without constructable stylesheets (styleNonce).
const CSP = "default-src 'self'; script-src 'self'; img-src 'self' blob: data:; style-src 'self'; connect-src 'self'; font-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'self'";
const STYLE_NONCE = "ufp-csp-nonce-test";
// The recommended policy with Trusted Types enforced and no policy allowed: the viewer must use no markup sink.
const CSP_TRUSTED_TYPES = `${CSP}; require-trusted-types-for 'script'; trusted-types 'none'`;
const CSP_NONCE = CSP.replace("style-src 'self'", `style-src 'self' 'nonce-${STYLE_NONCE}'`);
const cspFiles = Object.fromEntries(ooxmlFixtures.filter(fixture => fixture.options === undefined).map(fixture => [`${fixture.id}.${fixture.kind}`, buildFixture(fixture)]));
// Pattern fills are drawn with data: SVG backgrounds, which img-src must allow.
cspFiles["pattern.xlsx"] = workbook('<sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>pattern</t></is></c></row></sheetData>',
  `<styleSheet ${S}><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="darkGray"><fgColor rgb="FF2F5597"/><bgColor rgb="FFFFFFFF"/></patternFill></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf fontId="0" fillId="0"/><xf fontId="0" fillId="2" applyFill="1"/></cellXfs></styleSheet>`);
const files = { ...Object.fromEntries(Object.entries(fixtures).map(([name, fixture]) => [name, fixture.bytes])), ...viewFiles, ...cspFiles };

const page = `<!doctype html><meta charset="utf-8"><body style="margin:0"><div id="host" style="position:absolute;left:60px;top:60px;width:900px;height:600px"></div><script type="module">
import { FilePreviewer } from "/dist/index.js";
const host = document.getElementById("host");
const ATTACKER = ${JSON.stringify(ATTACKER)};
const frame = () => new Promise(done => requestAnimationFrame(() => requestAnimationFrame(done)));
localStorage.setItem("token", "host-secret");
let viewer;
async function show(name) {
  viewer?.destroy(); host.replaceChildren();
  const bytes = new Uint8Array(await (await fetch("/fixture/" + encodeURIComponent(name))).arrayBuffer());
  viewer = new FilePreviewer({ container: host, file: { name, source: bytes } });
  let error; try { await viewer.open(); } catch (e) { error = String(e?.message ?? e); }
  await new Promise(done => setTimeout(done, 600));
  return error;
}
window.inspect = async (name) => {
  const error = await show(name), rootElement = host.querySelector(".ufp-root");
  const urls = [...host.querySelectorAll("img")].map(image => image.src).concat([...host.querySelectorAll("image")].map(image => image.getAttribute("href") ?? ""));
  const blobs = [];
  for (const url of urls.filter(url => url.startsWith("blob:"))) {
    const blob = await (await fetch(url)).blob();
    blobs.push({ url, type: blob.type, text: blob.type === "image/svg+xml" ? await blob.text() : "" });
  }
  const fixed = [], attackerStyles = [], families = [];
  for (const element of rootElement.querySelectorAll("*")) {
    const style = getComputedStyle(element);
    if (style.position === "fixed") fixed.push(element.className.baseVal ?? element.className);
    for (const property of ["background-image", "font-family", "mask-image", "list-style-image", "border-image-source", "content", "cursor"]) if (style.getPropertyValue(property).includes(ATTACKER)) attackerStyles.push(property);
    if (element.matches(".ufp-grid-cell, .ufp-pptx-text span, .ufp-word-page span")) families.push(style.fontFamily);
  }
  const rect = rootElement.getBoundingClientRect();
  const outside = [...rootElement.querySelectorAll("*")].filter(element => { const box = element.getBoundingClientRect(); return box.width && box.height && (box.right < rect.left || box.bottom < rect.top); }).length;
  return { error, blobs, fixed, attackerStyles, families: [...new Set(families)], outside };
};
// Positive control for the CSS attack, performed without the viewer's defences.
window.controlCss = async () => {
  const element = document.createElement("div");
  element.style.cssText = 'width:10px;height:10px;font:11px "Calibri\\n;background-image:url(' + ATTACKER + '/control-css);x:"';
  document.body.append(element);
  await new Promise(done => setTimeout(done, 500));
  const injected = getComputedStyle(element).backgroundImage; element.remove();
  return injected;
};
window.controlSvgBlob = async () => URL.createObjectURL(new Blob([await (await fetch("/control.svg")).text()], { type: "image/svg+xml" }));
window.sanitizeAll = async (cases) => {
  const { sanitizeSvg } = await import("/dist/renderers/image-url.js");
  return cases.map(entry => { try { return { name: entry.name, output: new TextDecoder().decode(sanitizeSvg(new TextEncoder().encode(entry.svg))) }; } catch (e) { return { name: entry.name, error: String(e?.message ?? e) }; } });
};
/** Old (0.2.x, DOMParser) and new (DOM-free) sanitizer on the same input: normalized trees and rendered pixels. */
window.svgDifferential = async (cases) => {
  const { sanitizeSvg } = await import("/dist/renderers/image-url.js");
  const { referenceSanitizeSvg } = await import("/reference-sanitizer.js");
  const run = (sanitizer, svg) => { try { return { output: new TextDecoder().decode(sanitizer(new TextEncoder().encode(svg))) }; } catch (e) { return { error: String(e?.message ?? e) }; } };
  // Namespace and local name, attributes sorted (namespace declarations left out), adjacent text merged.
  const canonical = source => {
    const walk = node => {
      const children = []; let run = "";
      for (const child of node.childNodes) {
        if (child.nodeType === 3) { run += child.data; continue; }
        if (run) { children.push(run); run = ""; }
        children.push(child.nodeType === 1 ? walk(child) : "#node" + child.nodeType);
      }
      if (run) children.push(run);
      const attributes = [...node.attributes].filter(a => a.namespaceURI !== "http://www.w3.org/2000/xmlns/").map(a => [a.namespaceURI, a.localName, a.value]).sort((a, b) => (a[1] + " " + a[0]).localeCompare(b[1] + " " + b[0]));
      return [node.namespaceURI, node.localName, attributes, children];
    };
    return JSON.stringify(walk(new DOMParser().parseFromString(source, "image/svg+xml").documentElement));
  };
  const pixels = async source => {
    const url = URL.createObjectURL(new Blob([source], { type: "image/svg+xml" }));
    try {
      const image = new Image();
      await new Promise(done => { image.onload = image.onerror = done; image.src = url; });
      if (!image.naturalWidth) return "not drawn";
      const canvas = document.createElement("canvas"); canvas.width = 160; canvas.height = 160;
      const context = canvas.getContext("2d"); context.drawImage(image, 0, 0, 160, 160);
      const data = context.getImageData(0, 0, 160, 160).data; let hash = 0, ink = 0;
      for (let i = 0; i < data.length; i++) { hash = (Math.imul(hash, 31) + data[i]) | 0; if (i % 4 === 3 && data[i]) ink++; }
      return hash + "/" + ink;
    } finally { URL.revokeObjectURL(url); }
  };
  const results = [];
  for (const entry of cases) {
    const before = run(referenceSanitizeSvg, entry.svg), after = run(sanitizeSvg, entry.svg);
    const result = { name: entry.name, oldError: before.error, newError: after.error, oldOutput: before.output, newOutput: after.output };
    if (before.output && after.output) {
      const [oldTree, newTree] = [canonical(before.output), canonical(after.output)];
      const [oldPixels, newPixels] = [await pixels(before.output), await pixels(after.output)];
      Object.assign(result, { sameTree: oldTree === newTree, samePixels: oldPixels === newPixels, pixels: newPixels, ...(oldTree === newTree ? {} : { oldTree, newTree }) });
    }
    results.push(result);
  }
  return results;
};
window.svgUrl = (text) => URL.createObjectURL(new Blob([text], { type: "image/svg+xml" }));
window.loadImage = (url) => new Promise(done => { const image = new Image(); image.onload = () => { done({ ok: true, width: image.naturalWidth }); image.remove(); }; image.onerror = () => { done({ ok: false }); image.remove(); }; image.src = url; document.body.append(image); });
window.containment = async () => {
  await show("scroll.docx");
  const rootElement = host.querySelector(".ufp-root"), style = getComputedStyle(rootElement), rect = rootElement.getBoundingClientRect();
  // Integrator UI with position:fixed placed inside the viewer: positioned and clipped by the viewer.
  const probe = document.createElement("div");
  probe.style.cssText = "position:fixed;left:0;top:0;width:100vw;height:100vh;background:#c00;z-index:2147483647";
  host.querySelector(".ufp-content").append(probe);
  const box = probe.getBoundingClientRect(), outsideHit = document.elementFromPoint(10, 10) === probe, insideHit = document.elementFromPoint(rect.left + 20, rect.top + 20) === probe;
  probe.remove();
  return { contain: style.contain, isolation: style.isolation, overflow: style.overflow, root: rect.toJSON(), probe: box.toJSON(), outsideHit, insideHit };
};
const intersect = (a, b) => ({ left: Math.max(a.left, b.left), top: Math.max(a.top, b.top), right: Math.min(a.right, b.right), bottom: Math.min(a.bottom, b.bottom) });
const area = r => Math.max(0, r.right - r.left) * Math.max(0, r.bottom - r.top);
window.tooltipEdges = async () => {
  await show("tooltip.xlsx");
  const rootElement = host.querySelector(".ufp-root"), grid = host.querySelector(".ufp-grid-scroll"), body = host.querySelector(".ufp-grid-body");
  const g = grid.getBoundingClientRect(), inset = 40;
  const points = { "top-left": [g.left + 60, g.top + 30], "top-right": [g.right - inset, g.top + 30], "bottom-left": [g.left + 60, g.bottom - inset], "bottom-right": [g.right - inset, g.bottom - inset] };
  const measure = () => {
    const tip = host.querySelector(".ufp-input-message");
    if (!tip) return null;
    const r = tip.getBoundingClientRect(), clip = intersect(intersect(r, g), rootElement.getBoundingClientRect());
    const visible = area(r) ? area(clip) / area(r) : 0, x = (clip.left + clip.right) / 2, y = (clip.top + clip.bottom) / 2;
    return { rect: [r.left, r.top, r.width, r.height].map(Math.round), visible: Number(visible.toFixed(3)), topmost: visible > 0 && tip.contains(document.elementFromPoint(x, y)) };
  };
  const results = {};
  for (const [edge, [x, y]] of Object.entries(points)) {
    body.dispatchEvent(new MouseEvent("click", { clientX: x, clientY: y, bubbles: true }));
    await frame();
    const contained = measure();
    rootElement.style.contain = "none"; rootElement.style.isolation = "auto"; await frame();
    const uncontained = measure();
    rootElement.style.removeProperty("contain"); rootElement.style.removeProperty("isolation"); await frame();
    results[edge] = { contained, uncontained };
  }
  return results;
};
const cellRect = (label) => { const cell = [...host.querySelectorAll(".ufp-grid-cell")].find(element => (element.getAttribute("aria-label") ?? "").startsWith(label + ":")); return cell ? cell.getBoundingClientRect().toJSON() : null; };
window.frozenPanes = async () => {
  await show("tooltip.xlsx");
  const grid = host.querySelector(".ufp-grid-scroll");
  const before = { corner: cellRect("A1"), top: cellRect("B1"), left: cellRect("A3"), body: cellRect("C3") };
  grid.scrollLeft = 600; grid.scrollTop = 1500; grid.dispatchEvent(new Event("scroll"));
  await new Promise(done => setTimeout(done, 600));
  const rowOne = [...host.querySelectorAll(".ufp-grid-cell")].filter(element => /^[A-Z]+1:/.test(element.getAttribute("aria-label") ?? "") && !(element.getAttribute("aria-label") ?? "").startsWith("A1:")).map(element => element.getBoundingClientRect().top);
  const columnA = [...host.querySelectorAll(".ufp-grid-cell")].filter(element => /^A\\d+:/.test(element.getAttribute("aria-label") ?? "") && !(element.getAttribute("aria-label") ?? "").startsWith("A1:")).map(element => element.getBoundingClientRect().left);
  return { scrollLeft: grid.scrollLeft, scrollTop: grid.scrollTop, before, after: { corner: cellRect("A1"), body: cellRect("C3") }, rowOneTops: [...new Set(rowOne.map(Math.round))], columnALefts: [...new Set(columnA.map(Math.round))] };
};
window.scrolling = async () => {
  await show("scroll.xlsx");
  const grid = host.querySelector(".ufp-grid-scroll");
  grid.scrollTop = 4000; grid.dispatchEvent(new Event("scroll"));
  await new Promise(done => setTimeout(done, 600));
  const gridTop = grid.scrollTop, lateRow = [...host.querySelectorAll(".ufp-grid-cell")].some(cell => /^A(1[5-9]\\d|2\\d\\d)\\b/.test(cell.getAttribute("aria-label") ?? ""));
  await show("scroll.docx");
  const content = host.querySelector(".ufp-content"); content.scrollTop = 500;
  return { gridTop, lateRow, contentTop: content.scrollTop, contentScrollable: content.scrollHeight > content.clientHeight };
};
window.textLimits = async () => {
  const result = {};
  for (const name of ["shared-attack.xlsx", "shared-at-budget.xlsx"]) {
    const heap = performance.memory?.usedJSHeapSize ?? 0, started = performance.now(), error = await show(name);
    result[name] = { error: error ?? null, ms: Math.round(performance.now() - started - 600), domTextChars: host.textContent.length, heapGrowthMiB: Math.round(((performance.memory?.usedJSHeapSize ?? 0) - heap) / 1048576), status: host.querySelector(".ufp-status")?.textContent ?? "" };
  }
  return result;
};
window.rendererLimits = async () => {
  const started = performance.now(), error = await show("limits.xlsx");
  return { error, ms: Math.round(performance.now() - started - 600), cells: host.querySelectorAll(".ufp-grid-cell").length };
};
window.imageRelease = async () => {
  await show("image.xlsx");
  const grid = host.querySelector(".ufp-grid-scroll"), first = host.querySelector(".ufp-sheet-image")?.src;
  grid.scrollTop = 4000; grid.dispatchEvent(new Event("scroll")); await new Promise(done => setTimeout(done, 400));
  grid.scrollTop = 0; grid.dispatchEvent(new Event("scroll")); await new Promise(done => setTimeout(done, 400));
  const second = host.querySelector(".ufp-sheet-image")?.src;
  let revoked = false; try { await fetch(first); } catch { revoked = true; }
  return { first: !!first, second: !!second, rerendered: first !== second, revoked };
};
window.liveness = async () => { await fetch(ATTACKER + "/liveness-page", { mode: "no-cors" }); return true; };
window.pwned = () => localStorage.getItem("ufp-pwned");
window.resetPwned = () => localStorage.removeItem("ufp-pwned");
window.ready = true;
</script>`;

const cspScript = `import { FilePreviewer } from "/dist/index.js";
let violations = [];
document.addEventListener("securitypolicyviolation", event => violations.push(event.violatedDirective + " " + (event.blockedURI || "inline") + " " + (event.sourceFile || "")));
const sized = element => { element.style.setProperty("width", "900px"); element.style.setProperty("height", "600px"); return element; };
/** True when the viewer stylesheet applies to the viewer root (it sets isolation: isolate on .ufp-root). */
const styled = root => !!root && getComputedStyle(root).isolation === "isolate";
window.renderAll = async (names) => {
  const host = sized(document.getElementById("host")), results = [];
  for (const name of names) {
    host.replaceChildren(); violations = [];
    const bytes = new Uint8Array(await (await fetch("/fixture/" + encodeURIComponent(name))).arrayBuffer());
    const viewer = new FilePreviewer({ container: host, file: { name, source: bytes } });
    let error; try { await viewer.open(); } catch (e) { error = String(e?.message ?? e); }
    await new Promise(done => setTimeout(done, 300));
    const patterns = [...host.querySelectorAll(".ufp-cell-fill")].filter(element => getComputedStyle(element).backgroundImage.startsWith('url("data:image/svg+xml')).length;
    const images = [...host.querySelectorAll("img")].map(image => image.complete && image.naturalWidth > 0);
    results.push({ name, error, patterns, images: images.length, decoded: images.filter(Boolean).length, violations, rendered: !!host.querySelector(".ufp-content")?.children.length, styled: styled(host.querySelector(".ufp-root")), styleElements: host.querySelectorAll("style").length });
    viewer.destroy();
  }
  return results;
};
/** Stylesheet mechanics: one adopted sheet per document or shadow root, shared and removed with the last viewer. */
window.styleChecks = async () => {
  violations = [];
  const file = { name: "a.txt", source: new TextEncoder().encode("text") };
  const make = () => { const element = sized(document.createElement("div")); document.body.append(element); return element; };
  const before = document.adoptedStyleSheets.length, first = make(), second = make();
  const one = new FilePreviewer({ container: first, file }), two = new FilePreviewer({ container: second, file });
  const result = { adoptedWithTwo: document.adoptedStyleSheets.length - before, styleElements: document.querySelectorAll("style").length, styledBoth: styled(first.querySelector(".ufp-root")) && styled(second.querySelector(".ufp-root")) };
  one.destroy(); result.adoptedAfterOne = document.adoptedStyleSheets.length - before; result.secondStillStyled = styled(second.querySelector(".ufp-root"));
  two.destroy(); result.adoptedAfterBoth = document.adoptedStyleSheets.length - before;
  first.remove(); second.remove();
  // Shadow root: the sheet is adopted by the shadow root, not by the document.
  const shadowHost = make(), shadow = shadowHost.attachShadow({ mode: "open" }), inner = sized(document.createElement("div"));
  shadow.append(inner);
  const shadowed = new FilePreviewer({ container: inner, file });
  result.shadowAdopted = shadow.adoptedStyleSheets.length; result.shadowDocumentAdopted = document.adoptedStyleSheets.length - before; result.shadowStyled = styled(shadow.querySelector(".ufp-root"));
  shadowed.destroy(); result.shadowAfterDestroy = shadow.adoptedStyleSheets.length; shadowHost.remove();
  result.violations = violations;
  return result;
};
/** Browsers without constructable stylesheets (simulated): the viewer falls back to a <style> element. */
window.fallbackStyles = async (nonce) => {
  violations = [];
  const original = Object.getOwnPropertyDescriptor(CSSStyleSheet.prototype, "replaceSync");
  Object.defineProperty(CSSStyleSheet.prototype, "replaceSync", { value: undefined, configurable: true, writable: true });
  const element = sized(document.createElement("div")); document.body.append(element);
  let viewer;
  try { viewer = new FilePreviewer({ container: element, file: { name: "a.txt", source: new TextEncoder().encode("text") }, ...(nonce ? { styleNonce: nonce } : {}) }); }
  finally { Object.defineProperty(CSSStyleSheet.prototype, "replaceSync", original); }
  await new Promise(done => setTimeout(done, 100));
  const style = element.querySelector("style");
  const result = { styleElement: !!style, nonce: style?.nonce ?? "", styled: styled(element.querySelector(".ufp-root")), violations };
  viewer.destroy(); result.removed = !element.querySelector("style"); element.remove();
  return result;
};
window.ready = true;`;

const mime = { ".js": "text/javascript", ".css": "text/css", ".map": "application/json" };
const server = createServer(async (request, response) => {
  const url = new URL(request.url, "http://host");
  if (url.pathname === "/") { response.writeHead(200, { "content-type": "text/html" }); return response.end(page); }
  // No style attribute in the page: the policy blocks inline styles; the test script sizes the host through CSSOM.
  if (url.pathname === "/csp.html" || url.pathname === "/csp-nonce.html") { response.writeHead(200, { "content-type": "text/html", "content-security-policy": url.pathname === "/csp.html" ? CSP : CSP_NONCE }); return response.end('<!doctype html><meta charset="utf-8"><div id="host"></div><script type="module" src="/csp-test.js"></script>'); }
  if (url.pathname === "/reference-sanitizer.js") { response.writeHead(200, { "content-type": "text/javascript" }); return response.end(await readFile(join(root, "tests", "helpers", "svg-sanitizer-reference.mjs"))); }
  if (url.pathname === "/csp-tt.html") { response.writeHead(200, { "content-type": "text/html", "content-security-policy": CSP_TRUSTED_TYPES }); return response.end('<!doctype html><meta charset="utf-8"><div id="host"></div><script type="module" src="/csp-test.js"></script>'); }
  if (url.pathname === "/csp-test.js") { response.writeHead(200, { "content-type": "text/javascript" }); return response.end(cspScript); }
  if (url.pathname === "/control.svg") { response.writeHead(200, { "content-type": "text/plain" }); return response.end(evilSvg(ATTACKER)); }
  if (url.pathname === "/liveness") { response.writeHead(200); return response.end("ok"); }
  if (url.pathname.startsWith("/fixture/")) {
    const body = files[decodeURIComponent(url.pathname.slice("/fixture/".length))];
    response.writeHead(body ? 200 : 404, { "content-type": "application/octet-stream" }); return response.end(body);
  }
  if (url.pathname.startsWith("/dist/") && !url.pathname.includes("..")) {
    try { const body = await readFile(join(root, url.pathname)); response.writeHead(200, { "content-type": mime[extname(url.pathname)] ?? "application/octet-stream" }); return response.end(body); } catch { /* fall through */ }
  }
  response.writeHead(404); response.end();
});
await new Promise(done => server.listen(0, "127.0.0.1", done));
const HOST = `http://127.0.0.1:${server.address().port}`;

// -----------------------------------------------------------------------------------------------
// Browser (CDP)
// -----------------------------------------------------------------------------------------------

const started = Date.now();
const report = { tool: "scripts/browser-security.mjs", startedAt: new Date(started).toISOString(), node: process.version, platform: `${process.platform} ${process.arch}`, ...gitState(), browserPath, host: HOST, attacker: ATTACKER, steps: [], exceptions: [], attackerHits: hits, failures: [] };
const failures = report.failures;
let step = "startup";
const check = (condition, message) => { if (!condition) failures.push(`${step}: ${message}`); return condition; };
const record = (data) => { report.steps.push({ step, ...data }); };

const profile = await mkdtemp(join(tmpdir(), "ufp-security-profile-"));
const child = spawn(browserPath, ["--headless=new", "--remote-debugging-port=0", `--user-data-dir=${profile}`, "--no-first-run", "--no-default-browser-check", "--disable-extensions", "--disable-gpu", "about:blank"], { stdio: ["ignore", "ignore", "pipe"] });
let browserExited = false;
child.on("exit", () => { browserExited = true; });
const totalTimer = setTimeout(() => { failures.push(`total timeout of ${TOTAL_TIMEOUT_MS} ms exceeded during ${step}`); void finish(true); }, TOTAL_TIMEOUT_MS);

const endpoint = await new Promise((done, fail) => {
  let log = "";
  child.stderr.on("data", chunk => { log += chunk; const match = /DevTools listening on (ws:\/\/\S+)/.exec(log); if (match) done(match[1]); });
  child.on("exit", code => fail(new Error(`browser exited with code ${code}`)));
});
const socket = new WebSocket(endpoint);
await new Promise((done, fail) => { socket.addEventListener("open", done, { once: true }); socket.addEventListener("error", fail, { once: true }); });
let nextId = 0;
const pending = new Map(), sessionLabels = new Map(), openTargets = new Set();
socket.addEventListener("message", event => {
  const message = JSON.parse(event.data);
  if (message.method === "Runtime.exceptionThrown") {
    const details = message.params.exceptionDetails, label = sessionLabels.get(message.sessionId) ?? "unknown";
    report.exceptions.push({ step: label === "control" ? "control" : step, tab: label, text: details.exception?.description ?? details.text, url: details.url ?? "" });
    return;
  }
  if (message.method === "Inspector.targetCrashed") { failures.push(`${step}: a tab crashed (${sessionLabels.get(message.sessionId) ?? "unknown"})`); return; }
  const waiter = message.id !== undefined && pending.get(message.id);
  if (!waiter) return;
  pending.delete(message.id);
  if (message.error) waiter.fail(new Error(message.error.message)); else waiter.done(message.result);
});
socket.addEventListener("close", () => { for (const waiter of pending.values()) waiter.fail(new Error("the DevTools connection closed (browser crash?)")); pending.clear(); });
function withTimeout(promise, label, ms = STEP_TIMEOUT_MS) {
  let timer;
  return Promise.race([promise, new Promise((_, fail) => { timer = setTimeout(() => fail(new Error(`${label} timed out after ${ms} ms`)), ms); })]).finally(() => clearTimeout(timer));
}
const send = (method, params = {}, sessionId) => withTimeout(new Promise((done, fail) => {
  if (browserExited) return fail(new Error("the browser has exited"));
  const id = ++nextId; pending.set(id, { done, fail });
  socket.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
}), method);
const wait = ms => new Promise(done => setTimeout(done, ms));
async function evaluate(sessionId, expression) {
  const result = await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true }, sessionId);
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
  return result.result.value;
}
/** Opens a tab with exception reporting enabled before it navigates, so early errors are captured. */
async function openTab(url, label) {
  const { targetId } = await send("Target.createTarget", { url: "about:blank" });
  const { sessionId } = await send("Target.attachToTarget", { targetId, flatten: true });
  sessionLabels.set(sessionId, label); openTargets.add(targetId);
  await send("Runtime.enable", {}, sessionId); await send("Inspector.enable", {}, sessionId);
  await send("Page.navigate", { url }, sessionId);
  return { targetId, sessionId };
}
async function closeTab(tab) { openTargets.delete(tab.targetId); await send("Target.closeTarget", { targetId: tab.targetId }).catch(() => {}); }
async function waitReady(tab) {
  for (let attempt = 0; attempt < 100 && !(await evaluate(tab.sessionId, "!!window.ready").catch(() => false)); attempt++) await wait(100);
  if (!(await evaluate(tab.sessionId, "!!window.ready"))) throw new Error("the test page did not load");
}
/** Opens a blob URL as a top-level document (as "Open image in new tab" does) and returns its title. */
async function openAsDocument(url, label) {
  const tab = await openTab(url, label);
  await wait(1000);
  const title = await evaluate(tab.sessionId, "document.title").catch(error => `evaluate failed: ${error.message}`);
  await closeTab(tab);
  return title;
}
/** Fails the current step for uncaught exceptions raised while it ran. */
function checkExceptions() {
  const raised = report.exceptions.filter(entry => entry.step === step);
  check(!raised.length, `uncaught exception(s): ${raised.map(entry => `[${entry.tab}] ${entry.text.split("\n")[0]}`).join("; ")}`);
}

const allowedTypes = new Set(["image/png", "image/jpeg", "image/gif", "image/bmp", "image/webp", "image/svg+xml"]);
const escape = text => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const unsafeSvg = new RegExp(`<script|\\son[a-z]+\\s*=|javascript:|${escape(ATTACKER_HOST)}|<foreignObject|<image|<use|<style|<animate|<set[\\s>]|xlink|xml:base|xmlns(?::[\\w.-]+)?="(?!http://www\\.w3\\.org/2000/svg")|\\\\|url\\(\\s*[^#\\s]|<!--|<\\?|CDATA|xhtml|style=|class=`, "i");
let mainTab;

async function run() {
  const version = await send("Browser.getVersion");
  report.browser = { product: version.product, userAgent: version.userAgent, protocolVersion: version.protocolVersion, revision: version.revision };
  mainTab = await openTab(`${HOST}/`, "main");
  await waitReady(mainTab);
  console.log(`Browser ${version.product}; host ${HOST}; attacker ${ATTACKER}`);

  // Positive controls: without the defences, the harness detects both attacks.
  step = "control";
  const controlTitle = await openAsDocument(await evaluate(mainTab.sessionId, "window.controlSvgBlob()"), "control");
  check(controlTitle === "SCRIPT-RAN", `a raw SVG blob should run script when opened (title ${JSON.stringify(controlTitle)}); the harness cannot detect script in SVG`);
  const controlBackground = await evaluate(mainTab.sessionId, "window.controlCss()");
  await wait(300);
  check(controlBackground.includes(ATTACKER) && hits.includes("/control-css"), "injected cssText should load the attacker URL; the harness cannot detect CSS injection");
  await evaluate(mainTab.sessionId, "window.resetPwned()");
  await wait(1000);
  const afterControls = hits.length;
  record({ controlTitle, controlHits: hits.slice(0, afterControls) });
  console.log(`  controls: raw SVG blob title=${JSON.stringify(controlTitle)}; injected cssText requested ${hits.includes("/control-css") ? "/control-css" : "nothing"}`);

  for (const [name, fixture] of Object.entries(fixtures)) {
    step = name;
    const before = failures.length, mark = hits.length, result = await evaluate(mainTab.sessionId, `window.inspect(${JSON.stringify(name)})`);
    check(!result.error, `preview failed: ${result.error}`);
    for (const blob of result.blobs) {
      check(allowedTypes.has(blob.type), `blob type ${JSON.stringify(blob.type)} is not allowlisted`);
      if (blob.type === "image/svg+xml") check(!unsafeSvg.test(blob.text), `SVG blob still contains active or external content: ${blob.text.slice(0, 200)}`);
      blob.openedTitle = await openAsDocument(blob.url, `${name} blob`);
      check(blob.openedTitle !== "SCRIPT-RAN", `script ran when a ${blob.type} blob was opened as a document`);
    }
    const svgCount = result.blobs.filter(blob => blob.type === "image/svg+xml").length;
    check(svgCount === fixture.expectSvgImages, `expected ${fixture.expectSvgImages} sanitized SVG image(s), found ${svgCount}`);
    check(!result.fixed.length, `position:fixed elements: ${result.fixed.join(", ")}`);
    check(!result.attackerStyles.length, `computed styles refer to the attacker: ${result.attackerStyles.join(", ")}`);
    check(!result.outside, `${result.outside} element(s) lie entirely above or left of the viewer`);
    await wait(300);
    check(hits.length === mark, `attacker server received ${hits.slice(mark).join(", ")}`);
    check(!(await evaluate(mainTab.sessionId, "window.pwned()")), "the payload wrote to the host origin's localStorage");
    checkExceptions();
    await evaluate(mainTab.sessionId, "window.resetPwned()");
    record({ blobs: result.blobs.map(({ text, url, ...blob }) => blob), families: result.families, passed: failures.length === before });
    console.log(`  ${failures.length === before ? "ok  " : "FAIL"} ${name}: ${result.blobs.map(blob => blob.type).join(", ") || "no images"}${result.families.length ? `; fonts ${result.families.join(" | ")}` : ""}`);
  }

  // The SVG sanitizer on its own: namespace confusion, escaped url() values, entities, invalid roots.
  step = "svg sanitizer";
  {
    const before = failures.length, mark = hits.length;
    const outputs = await evaluate(mainTab.sessionId, `window.sanitizeAll(${JSON.stringify(svgCases)})`);
    const rows = [];
    for (const [index, entry] of svgCases.entries()) {
      const output = outputs[index], row = { name: entry.name, rejected: !!output.error };
      if (entry.invalid) check(!!output.error, `${entry.name}: should be rejected, produced ${String(output.output).slice(0, 120)}`);
      else if (check(!output.error, `${entry.name}: rejected unexpectedly (${output.error})`)) {
        check(!unsafeSvg.test(output.output), `${entry.name}: output still contains active or external content: ${output.output}`);
        for (const text of entry.keep ?? []) check(output.output.includes(text), `${entry.name}: legitimate content ${JSON.stringify(text)} was removed: ${output.output}`);
        const url = await evaluate(mainTab.sessionId, `window.svgUrl(${JSON.stringify(output.output)})`);
        const image = await evaluate(mainTab.sessionId, `window.loadImage(${JSON.stringify(url)})`);
        check(image.ok, `${entry.name}: the sanitized SVG does not load as an image`);
        row.title = await openAsDocument(url, `sanitizer ${entry.name}`);
        check(row.title !== "SCRIPT-RAN", `${entry.name}: script ran when the sanitized SVG was opened as a document`);
        row.output = output.output;
      }
      rows.push(row);
    }
    await wait(500);
    check(hits.length === mark, `attacker server received ${hits.slice(mark).join(", ")}`);
    check(!(await evaluate(mainTab.sessionId, "window.pwned()")), "a sanitized SVG wrote to the host origin's localStorage");
    checkExceptions();
    record({ cases: rows });
    console.log(`  ${failures.length === before ? "ok  " : "FAIL"} svg sanitizer: ${svgCases.length} cases (${rows.filter(row => row.rejected).length} rejected as invalid)`);
  }

  step = "svg differential";
  {
    // Every sanitizer case, the SVG embedded in the media fixtures and the demo illustration, through the 0.2.x
    // DOMParser sanitizer and the DOM-free one: inputs with a DOCTYPE must now be refused (approved change);
    // every other input must be accepted or refused by both, with the same tree and the same pixels.
    const before = failures.length;
    const { writeDemoSamples } = await import(pathToFileURL(join(root, "scripts", "demo-samples.mjs")));
    const { mkdtemp, rm } = await import("node:fs/promises"), { tmpdir } = await import("node:os");
    const demo = await mkdtemp(join(tmpdir(), "ufp-svg-demo-"));
    let illustration;
    try { await writeDemoSamples(demo); illustration = await readFile(join(demo, "sample-illustration.svg"), "utf8"); } finally { await rm(demo, { recursive: true, force: true }); }
    const inputs = [...svgCases.map(entry => ({ name: entry.name, svg: entry.svg })), { name: "embedded media SVG (evilSvg)", svg: evilSvg(ATTACKER) }, { name: "demo sample-illustration.svg", svg: illustration }];
    const results = await evaluate(mainTab.sessionId, `window.svgDifferential(${JSON.stringify(inputs)})`);
    let identical = 0, refusedByBoth = 0, doctypes = 0;
    for (const [index, result] of results.entries()) {
      if (/<!DOCTYPE/i.test(inputs[index].svg)) { doctypes++; check(!!result.newError, `${result.name}: an SVG with a DOCTYPE was not refused`); continue; }
      check(!!result.oldError === !!result.newError, `${result.name}: old sanitizer ${result.oldError ? "refused" : "accepted"}, new sanitizer ${result.newError ? "refused (" + result.newError + ")" : "accepted"}`);
      if (result.oldError && result.newError) { refusedByBoth++; continue; }
      if (result.oldError || result.newError) continue;
      check(result.sameTree, `${result.name}: the sanitized trees differ:\n  old ${result.oldTree}\n  new ${result.newTree}`);
      check(result.samePixels, `${result.name}: the sanitized images render differently`);
      check(!unsafeSvg.test(result.newOutput), `${result.name}: new output contains active or external content: ${result.newOutput}`);
      if (result.sameTree && result.samePixels) identical++;
    }
    checkExceptions();
    record({ results: results.map(({ oldOutput, ...rest }) => rest) });
    console.log(`  ${failures.length === before ? "ok  " : "FAIL"} svg differential: ${inputs.length} inputs; ${identical} with identical trees and pixels, ${refusedByBoth} refused by both, ${doctypes} with a DOCTYPE now refused`);
  }

  step = "containment";
  {
    const before = failures.length, contained = await evaluate(mainTab.sessionId, "window.containment()");
    // Browsers report "layout paint style" as the equivalent keyword "content".
    check(contained.contain === "content" || contained.contain === "strict" || ["layout", "paint", "style"].every(value => contained.contain.includes(value)), `.ufp-root contain is ${JSON.stringify(contained.contain)}`);
    check(contained.isolation === "isolate" && contained.overflow === "hidden", `isolation ${contained.isolation}, overflow ${contained.overflow}`);
    check(contained.probe.left >= contained.root.left && contained.probe.top >= contained.root.top && !contained.outsideHit, `a position:fixed element escaped the viewer (${JSON.stringify(contained)})`);
    checkExceptions();
    record(contained);
    console.log(`  ${failures.length === before ? "ok  " : "FAIL"} containment: contain=${contained.contain}; fixed-position element inside the viewer placed at ${contained.probe.left},${contained.probe.top} (viewer at ${contained.root.left},${contained.root.top}), ${Math.round(contained.probe.width)}x${Math.round(contained.probe.height)}, clipped to the viewer`);
  }

  step = "input-message tooltip";
  {
    const before = failures.length, edges = await evaluate(mainTab.sessionId, "window.tooltipEdges()");
    for (const [edge, { contained, uncontained }] of Object.entries(edges)) {
      check(contained !== null, `${edge}: no input message was shown`);
      check(JSON.stringify(contained) === JSON.stringify(uncontained), `${edge}: the tooltip differs with and without containment (${JSON.stringify(contained)} vs ${JSON.stringify(uncontained)})`);
    }
    check(edges["top-left"]?.contained?.visible === 1, "top-left: the input message should be fully visible");
    checkExceptions();
    record({ edges });
    console.log(`  ${failures.length === before ? "ok  " : "FAIL"} input-message tooltip: visible fraction ${Object.entries(edges).map(([edge, value]) => `${edge} ${value.contained?.visible}`).join(", ")} (identical without containment)`);
  }

  step = "frozen panes";
  {
    const before = failures.length, frozen = await evaluate(mainTab.sessionId, "window.frozenPanes()");
    check(frozen.scrollLeft > 0 && frozen.scrollTop > 0, `the grid did not scroll (${frozen.scrollLeft}, ${frozen.scrollTop})`);
    check(JSON.stringify(frozen.before.corner) === JSON.stringify(frozen.after.corner), "the frozen corner cell A1 moved");
    check(frozen.rowOneTops.length === 1 && frozen.rowOneTops[0] === Math.round(frozen.before.top.top), `frozen row 1 is not pinned (tops ${frozen.rowOneTops}, expected ${Math.round(frozen.before.top?.top)})`);
    check(frozen.columnALefts.length === 1 && frozen.columnALefts[0] === Math.round(frozen.before.left.left), `frozen column A is not pinned (lefts ${frozen.columnALefts}, expected ${Math.round(frozen.before.left?.left)})`);
    check(JSON.stringify(frozen.before.body) !== JSON.stringify(frozen.after.body), "unfrozen cells did not move");
    checkExceptions();
    record(frozen);
    console.log(`  ${failures.length === before ? "ok  " : "FAIL"} frozen panes: scrolled to ${frozen.scrollLeft},${frozen.scrollTop}; row 1 and column A stay pinned`);
  }

  step = "scrolling";
  {
    const before = failures.length, scroll = await evaluate(mainTab.sessionId, "window.scrolling()");
    check(scroll.gridTop > 0 && scroll.lateRow, `grid did not scroll or render later rows (${JSON.stringify(scroll)})`);
    check(scroll.contentScrollable && scroll.contentTop > 0, `document view did not scroll (${JSON.stringify(scroll)})`);
    checkExceptions();
    record(scroll);
    console.log(`  ${failures.length === before ? "ok  " : "FAIL"} scrolling: grid scrollTop ${scroll.gridTop}, later rows rendered ${scroll.lateRow}; document scrollTop ${scroll.contentTop}`);
  }

  step = "xlsx text limits";
  {
    const before = failures.length, text = await evaluate(mainTab.sessionId, "window.textLimits()");
    const attack = text["shared-attack.xlsx"], atBudget = text["shared-at-budget.xlsx"];
    check(/32,767/.test(attack.error ?? ""), `the 0.2.1 shared-string attack was not refused (${JSON.stringify(attack)})`);
    check(attack.ms < 3000 && attack.domTextChars < 10_000, `the refused attack took ${attack.ms} ms or left ${attack.domTextChars} characters in the page`);
    check(!atBudget.error, `the largest accepted shared-string workbook failed: ${atBudget.error}`);
    // Residual bound, as for ODS: each visible cell may hold up to 32,767 characters (recorded, not hidden).
    check(atBudget.ms < 5000, `the largest accepted shared-string workbook took ${atBudget.ms} ms`);
    checkExceptions();
    record(text);
    console.log(`  ${failures.length === before ? "ok  " : "FAIL"} xlsx text limits: attack refused in ${attack.ms} ms (${attack.domTextChars} chars in the page); largest accepted file rendered in ${atBudget.ms} ms with ${(atBudget.domTextChars / 1e6).toFixed(1)} M chars in the page`);
  }

  step = "renderer limits";
  {
    const before = failures.length, limits = await evaluate(mainTab.sessionId, "window.rendererLimits()"), images = await evaluate(mainTab.sessionId, "window.imageRelease()");
    check(!limits.error && limits.cells > 0, `the crafted sheet did not render (${limits.error ?? "no cells"})`);
    check(limits.ms < 5000, `5,000 overlapping merges and 20,000 column ranges took ${limits.ms} ms`);
    check(images.first && images.second && images.rerendered, `the sheet image was not re-rendered (${JSON.stringify(images)})`);
    check(images.revoked, "the previous image blob URL is still alive after the grid re-rendered");
    checkExceptions();
    record({ limits, images });
    console.log(`  ${failures.length === before ? "ok  " : "FAIL"} renderer limits: crafted merges and columns rendered in ${limits.ms} ms; previous image URL revoked on re-render: ${images.revoked}`);
  }

  step = "csp";
  {
    const before = failures.length, cspTab = await openTab(`${HOST}/csp.html`, "csp");
    await waitReady(cspTab);
    const results = await evaluate(cspTab.sessionId, `window.renderAll(${JSON.stringify(Object.keys(files))})`);
    for (const result of results) {
      // Files that must be refused (the shared-string attack) are expected to fail under any policy.
      if (refusedFiles.has(result.name)) check(!!result.error, `${result.name} should be refused`);
      else check(!result.error && result.rendered, `${result.name} did not render (${result.error ?? "empty"})`);
      check(result.styled && result.styleElements === 0, `${result.name}: the viewer stylesheet did not apply without 'unsafe-inline' (styled ${result.styled}, ${result.styleElements} <style> elements)`);
      // The sanitizer parses SVG without the DOM, so even malicious SVG causes no policy reports (0.2.x reported
      // blocked inline handlers and styles while parsing it with DOMParser).
      check(!result.violations.length, `${result.name} caused violations: ${[...new Set(result.violations)].join("; ")}`);
    }
    const patterns = results.find(result => result.name === "pattern.xlsx")?.patterns ?? 0;
    // Some fixtures carry deliberately truncated images; what matters is that blob: images decode under the policy.
    const images = results.reduce((sum, result) => sum + result.decoded, 0);
    check(patterns > 0, "the pattern fill was not drawn, so the data: image path did not run");
    check(images > 0, "no image was decoded, so the blob: image path did not run");
    // Stylesheet mechanics under the same policy.
    const sheets = await evaluate(cspTab.sessionId, "window.styleChecks()");
    check(sheets.adoptedWithTwo === 1 && sheets.styleElements === 0 && sheets.styledBoth, `two viewers should share one adopted stylesheet (${JSON.stringify(sheets)})`);
    check(sheets.adoptedAfterOne === 1 && sheets.secondStillStyled && sheets.adoptedAfterBoth === 0, `the shared stylesheet should stay until the last viewer is destroyed (${JSON.stringify(sheets)})`);
    check(sheets.shadowAdopted === 1 && sheets.shadowDocumentAdopted === 0 && sheets.shadowStyled && sheets.shadowAfterDestroy === 0, `a viewer in a shadow root should adopt the stylesheet there (${JSON.stringify(sheets)})`);
    check(!sheets.violations.length, `stylesheet checks caused violations: ${sheets.violations.join("; ")}`);
    // Without constructable stylesheets the fallback <style> element is blocked by the recommended policy ...
    const blocked = await evaluate(cspTab.sessionId, "window.fallbackStyles()");
    check(blocked.styleElement && !blocked.styled && blocked.violations.some(violation => violation.startsWith("style-src")), `the fallback <style> element should be blocked without a nonce (${JSON.stringify(blocked)})`);
    check(blocked.removed, "the fallback <style> element was not removed by destroy()");
    checkExceptions();
    await closeTab(cspTab);
    // ... and applies with styleNonce under a policy that allows that nonce.
    const nonceTab = await openTab(`${HOST}/csp-nonce.html`, "csp-nonce");
    await waitReady(nonceTab);
    const allowed = await evaluate(nonceTab.sessionId, `window.fallbackStyles(${JSON.stringify(STYLE_NONCE)})`);
    check(allowed.styleElement && allowed.styled && allowed.nonce === STYLE_NONCE && !allowed.violations.length, `the fallback <style> element should apply with styleNonce (${JSON.stringify(allowed)})`);
    checkExceptions();
    await closeTab(nonceTab);
    // Every file again with Trusted Types enforced and no policy: no markup sink is used anywhere.
    const trustedTab = await openTab(`${HOST}/csp-tt.html`, "csp-trusted-types");
    await waitReady(trustedTab);
    const trusted = await evaluate(trustedTab.sessionId, `window.renderAll(${JSON.stringify(Object.keys(files))})`);
    for (const result of trusted) {
      if (refusedFiles.has(result.name)) check(!!result.error, `Trusted Types: ${result.name} should be refused`);
      else check(!result.error && result.rendered && result.styled, `Trusted Types: ${result.name} did not render (${result.error ?? "empty"})`);
      check(!result.violations.length, `Trusted Types: ${result.name} caused violations: ${[...new Set(result.violations)].join("; ")}`);
    }
    const trustedImages = trusted.reduce((sum, result) => sum + result.decoded, 0);
    check(trustedImages > 0, "no image was decoded under Trusted Types");
    checkExceptions();
    await closeTab(trustedTab);
    const handlerReports = results.reduce((sum, result) => sum + result.violations.length, 0);
    record({ policy: CSP, results, sheets, fallback: { blocked, allowed }, trustedTypes: { policy: CSP_TRUSTED_TYPES, results: trusted } });
    console.log(`  ${failures.length === before ? "ok  " : "FAIL"} csp: ${results.length} files rendered and styled under the recommended policy (no 'unsafe-inline'); ${patterns} pattern fill(s), ${images} decoded image(s); ${handlerReports} policy reports; Trusted Types enforced: ${trusted.length} files rendered, ${trustedImages} decoded image(s), no violations`);
    console.log(`  ${failures.length === before ? "ok  " : "FAIL"} stylesheet: one shared adopted sheet per root (shadow root included), removed with the last viewer; fallback <style> blocked without a nonce, applied with styleNonce`);
  }

  // Final accounting, after every tab opened by the tests has closed.
  step = "final";
  await evaluate(mainTab.sessionId, "window.liveness()").catch(error => failures.push(`final: the page could not reach the attacker server: ${error.message}`));
  await fetch(`${ATTACKER}/liveness-node`).catch(error => failures.push(`final: the attacker server is not answering: ${error.message}`));
  await fetch(`${HOST}/liveness`).catch(error => failures.push(`final: the host server is not answering: ${error.message}`));
  for (const targetId of [...openTargets]) if (targetId !== mainTab.targetId) await closeTab({ targetId });
  await wait(2000);
  const late = hits.slice(afterControls).filter(path => !LIVENESS.includes(path));
  check(LIVENESS.every(path => hits.includes(path)), `the attacker server did not record the liveness requests (${hits.slice(-5).join(", ")})`);
  check(!late.length, `attacker server received ${late.length} request(s) after the controls: ${late.join(", ")}`);
  const unattributed = report.exceptions.filter(entry => entry.step !== "control" && !report.steps.some(recorded => recorded.step === entry.step));
  check(!unattributed.length, `uncaught exception(s) outside the checks: ${unattributed.map(entry => entry.text.split("\n")[0]).join("; ")}`);
  record({ attackerRequestsAfterControls: late.length, controlExceptions: report.exceptions.filter(entry => entry.step === "control").length });
  console.log(`  ${late.length ? "FAIL" : "ok  "} attacker server: ${afterControls} request(s) during the controls, ${late.length} afterwards (liveness confirmed from Node and from the page)`);
}

let finished = false;
async function finish(timedOut = false) {
  if (finished) return;
  finished = true;
  clearTimeout(totalTimer);
  try { socket.close(); } catch { /* already closed */ }
  child.kill(); attacker.close(); server.close();
  await wait(500);
  await rm(profile, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }).catch(() => {});
  report.durationMs = Date.now() - started;
  report.passed = !failures.length && !timedOut;
  await mkdir(dirname(options.report), { recursive: true });
  await writeFile(options.report, `${JSON.stringify(report, null, 2)}\n`);
  if (failures.length) {
    console.error(`\n${failures.length} check(s) failed:`);
    for (const failure of failures) console.error(`  - ${failure}`);
  } else console.log("\nAll browser security checks passed.");
  console.log(`Report: ${options.report}`);
  process.exit(report.passed ? 0 : 1);
}

try { await run(); } catch (error) { failures.push(`${step}: ${error.message}`); }
await finish();
