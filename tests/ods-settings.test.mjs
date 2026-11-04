// settings.xml view hints (implementation specification M14): best-effort, never affecting cell content.
import test from "node:test";
import assert from "node:assert/strict";
import { parseOds } from "../dist/parsers/odf/ods.js";
import { PreviewError } from "../dist/errors.js";
import { makeOdfPackage, odfXml, spreadsheetContent, spreadsheetStyles } from "./helpers/odf.mjs";

const PREFIX_SETS = [["canonical prefixes", {}], ["renamed prefixes", { office: "o", config: "cfg", table: "t", text: "x", style: "s", fo: "f", svg: "g", manifest: "mf" }]];
const SHEETS = ["Summary", "Data", "Hidden"].map(name => `<table:table table:name="${name}"${name === "Hidden" ? ' table:style-name="taHidden"' : ""}><table:table-row><table:table-cell office:value-type="string"><text:p>${name}</text:p></table:table-cell></table:table-row></table:table>`).join("");
const AUTOMATIC = '<style:style style:name="taHidden" style:family="table"><style:table-properties table:display="false"/></style:style>';

const configItem = (name, type, value) => `<config:config-item config:name="${name}" config:type="${type}">${value}</config:config-item>`;
/** One LibreOffice-shaped per-sheet entry; `items` overrides or adds config items. */
function tableEntry(name, items = {}) {
  const defaults = { CursorPositionX: ["int", 0], CursorPositionY: ["int", 0], HorizontalSplitMode: ["short", 0], VerticalSplitMode: ["short", 0], HorizontalSplitPosition: ["int", 0], VerticalSplitPosition: ["int", 0], ActiveSplitRange: ["short", 2], PositionLeft: ["int", 0], PositionRight: ["int", 0], PositionTop: ["int", 0], PositionBottom: ["int", 0], ZoomType: ["short", 0], ZoomValue: ["int", 100], PageViewZoomValue: ["int", 60], ShowGrid: ["boolean", "true"], HasColumnRowHeaders: ["boolean", "true"] };
  const merged = { ...defaults, ...items };
  return `<config:config-item-map-entry config:name="${name}">${Object.entries(merged).filter(([, value]) => value !== undefined).map(([key, [type, value]]) => configItem(key, type, value)).join("")}</config:config-item-map-entry>`;
}
function settingsXml({ tables = "", view = {}, extraSets = "" } = {}, prefixes = {}) {
  const viewItems = { ViewId: ["string", "view1"], ActiveTable: ["string", "Summary"], HorizontalScrollbarWidth: ["int", 1000], ZoomType: ["short", 0], ZoomValue: ["int", 100], ShowGrid: ["boolean", "true"], HasColumnRowHeaders: ["boolean", "true"], ...view };
  const items = Object.entries(viewItems).filter(([, value]) => value !== undefined).map(([key, [type, value]]) => configItem(key, type, value)).join("");
  return odfXml(`<?xml version="1.0" encoding="UTF-8"?><office:document-settings office:version="1.3"><office:settings>
    <config:config-item-set config:name="ooo:view-settings">${configItem("VisibleAreaTop", "int", 0)}
      <config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>
        <config:config-item-map-named config:name="Tables">${tables}</config:config-item-map-named>${items}
      </config:config-item-map-entry></config:config-item-map-indexed>
    </config:config-item-set>
    <config:config-item-set config:name="ooo:configuration-settings">${configItem("ShowGrid", "boolean", "false")}${configItem("ShowZeroValues", "boolean", "true")}</config:config-item-set>${extraSets}
  </office:settings></office:document-settings>`, prefixes);
}
function parse(settings, prefixes = {}, options = {}) {
  const bytes = makeOdfPackage({ content: spreadsheetContent({ body: SHEETS, automatic: AUTOMATIC }, prefixes), styles: spreadsheetStyles({}, prefixes), ...(settings !== undefined ? { settings } : {}) });
  return parseOds(bytes, "view.ods", options);
}
const viewOf = sheet => Object.fromEntries(["frozen", "zoom", "showGridLines", "showHeaders"].filter(key => sheet[key] !== undefined).map(key => [key, sheet[key]]));
const ignored = workbook => workbook.diagnostics.find(entry => entry.code === "odf.settings.ignored");

for (const [label, prefixes] of PREFIX_SETS) {
  test(`[${label}] LibreOffice view settings: active sheet, frozen panes, zoom, grid and headers`, () => {
    const tables = [
      tableEntry("Summary", { HorizontalSplitMode: ["short", 2], VerticalSplitMode: ["short", 2], HorizontalSplitPosition: ["int", 1], VerticalSplitPosition: ["int", 3], ZoomValue: ["int", 150] }),
      tableEntry("Data", { VerticalSplitMode: ["short", 2], VerticalSplitPosition: ["int", 1], ShowGrid: ["boolean", "false"], HasColumnRowHeaders: ["boolean", "false"], ZoomValue: ["int", 75] }),
      tableEntry("Hidden"),
    ].join("");
    const workbook = parse(settingsXml({ tables, view: { ActiveTable: ["string", "Data"] } }, prefixes), prefixes);
    assert.equal(workbook.activeSheet, 1, `${label}: ActiveTable`);
    assert.deepEqual(viewOf(workbook.sheets[0]), { frozen: { rows: 3, columns: 1 }, zoom: 1.5, showGridLines: true, showHeaders: true }, label);
    assert.deepEqual(viewOf(workbook.sheets[1]), { frozen: { rows: 1, columns: 0 }, zoom: 0.75, showGridLines: false, showHeaders: false }, label);
    assert.deepEqual(viewOf(workbook.sheets[2]), { showGridLines: true, showHeaders: true }, `${label}: 100% is the default zoom`);
    assert.equal(ignored(workbook), undefined, label);
    assert.deepEqual(workbook.sheets[1].cells.map(cell => cell.value), ["Data"], `${label}: content unchanged`);
  });
}

test("a missing settings.xml is fine and leaves the defaults", () => {
  const workbook = parse(undefined);
  assert.equal(workbook.activeSheet, undefined);
  for (const sheet of workbook.sheets) assert.deepEqual(viewOf(sheet), {});
  assert.deepEqual(workbook.diagnostics.map(entry => entry.code), ["odf.hidden.sheets"], "only the fixture's hidden sheet");
});

test("the first sheet as active sheet is the default and not recorded", () => {
  const workbook = parse(settingsXml({ tables: tableEntry("Summary") }));
  assert.equal("activeSheet" in workbook, false);
  assert.equal(ignored(workbook), undefined);
});

test("only frozen splits (mode 2) freeze panes; split windows and fit zooms are ignored silently", () => {
  const tables = [
    tableEntry("Summary", { HorizontalSplitMode: ["short", 1], HorizontalSplitPosition: ["int", 4800], VerticalSplitMode: ["short", 1], VerticalSplitPosition: ["int", 2400] }),
    tableEntry("Data", { ZoomType: ["short", 1], ZoomValue: ["int", 63] }),
  ].join("");
  const workbook = parse(settingsXml({ tables }));
  assert.deepEqual(viewOf(workbook.sheets[0]), { showGridLines: true, showHeaders: true });
  assert.deepEqual(viewOf(workbook.sheets[1]), { showGridLines: true, showHeaders: true });
  assert.equal(ignored(workbook), undefined);
});

test("sheets without an entry use the view-wide grid and header settings", () => {
  const workbook = parse(settingsXml({ tables: tableEntry("Summary", { ShowGrid: undefined, HasColumnRowHeaders: undefined }), view: { ShowGrid: ["boolean", "false"], HasColumnRowHeaders: ["boolean", "false"] } }));
  for (const sheet of workbook.sheets) assert.deepEqual(viewOf(sheet), { showGridLines: false, showHeaders: false }, sheet.name);
});

test("unreadable values are ignored one by one, reported as info, and never change content", () => {
  const tables = [
    tableEntry("Summary", { ZoomValue: ["int", "huge"], HorizontalSplitMode: ["short", 2], HorizontalSplitPosition: ["int", 2] }),
    tableEntry("Data", { ZoomValue: ["int", 5000], ShowGrid: ["boolean", "maybe"] }),
    tableEntry("Hidden", { VerticalSplitMode: ["short", 2], VerticalSplitPosition: ["int", -1] }),
  ].join("");
  const workbook = parse(settingsXml({ tables, view: { ActiveTable: ["string", "No such sheet"] } }));
  assert.equal(workbook.activeSheet, undefined);
  assert.deepEqual(viewOf(workbook.sheets[0]), { frozen: { rows: 0, columns: 2 }, showGridLines: true, showHeaders: true }, "the readable values of an entry still apply");
  assert.deepEqual(viewOf(workbook.sheets[1]), { showGridLines: true, showHeaders: true }, "an unreadable entry value falls back to the view-wide one");
  assert.deepEqual(viewOf(workbook.sheets[2]), { showGridLines: true, showHeaders: true });
  assert.deepEqual(ignored(workbook), { code: "odf.settings.ignored", severity: "info", message: "View settings could not be read.", count: 5, part: "settings.xml" });
});

test("a settings.xml that cannot be parsed, or is not a settings document, is ignored", () => {
  for (const [label, settings] of [
    ["not well-formed", "<office:document-settings"],
    ["DTD", '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e "x">]><x/>'],
    ["unbound prefix", "<config:x/>"],
    ["wrong root", odfXml('<office:document-styles office:version="1.3"/>')],
    ["no office:settings", odfXml('<office:document-settings office:version="1.3"/>')],
  ]) {
    const workbook = parse(settings);
    assert.equal(ignored(workbook)?.severity, "info", label);
    assert.deepEqual(workbook.sheets.map(sheet => sheet.name), ["Summary", "Data", "Hidden"], `${label}: the spreadsheet still opens`);
    for (const sheet of workbook.sheets) assert.deepEqual(viewOf(sheet), {}, label);
  }
  const noViews = parse(odfXml('<office:document-settings office:version="1.3"><office:settings><config:config-item-set config:name="ooo:configuration-settings"/></office:settings></office:document-settings>'));
  assert.equal(ignored(noViews), undefined, "settings without view settings are simply absent");
});

test("cancellation while reading settings still stops the preview", () => {
  const controller = new AbortController();
  controller.abort();
  assert.throws(() => parse(settingsXml({ tables: tableEntry("Summary") }), {}, { signal: controller.signal }), error => error instanceof PreviewError && error.code === "aborted");
});
