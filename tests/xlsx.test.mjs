import test from "node:test";
import assert from "node:assert/strict";
import { inflateSync } from "node:zlib";
import { makeZip } from "./helpers/zip.mjs";
import { parseXlsx } from "../dist/parsers/ooxml.js";
import { formatValue, formatGeneral } from "../dist/parsers/xlsx-number-format.js";
import { evaluateFormula } from "../dist/parsers/xlsx-formula.js";

const workbookXml = (sheets, extra = "") => `<workbook xmlns:r="urn:r"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${sheets}</sheets>${extra}</workbook>`;
const rels = ids => `<Relationships>${ids.map(([id, target, type = "worksheet"]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`).join("")}</Relationships>`;
const theme = `<a:theme xmlns:a="urn:a"><a:themeElements><a:clrScheme name="t"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="735773"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Corbel"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;

test("XLSX number formats follow Excel sections, scaling, dates and locale separators", () => {
  const cases = [
    [1234.5678, "#,##0.00", "1,234.57", "1.234,57"],
    [-1234.5, "#,##0.00;[Red](#,##0.00)", "(1,234.50)", "(1.234,50)"],
    [0.256, "0.0%", "25.6%", "25,6%"],
    [12345.678, "0.00E+00", "1.23E+04", "1,23E+04"],
    [1.5, "# ?/?", "1 1/2", "1 1/2"],
    [1500000, '#,##0,,"M"', "2M", "2M"],
    [46323, "dddd, d. mmmm yyyy", "Wednesday, 28. October 2026", "Mittwoch, 28. Oktober 2026"],
    [46323.5625, "dd.mm.yyyy hh:mm", "28.10.2026 13:30", "28.10.2026 13:30"],
    [1.75, "[h]:mm", "42:00", "42:00"],
    [5551234567, "[<=9999999]###-####;(###) ###-####", "(555) 123-4567", "(555) 123-4567"],
    [0, '#,##0;-#,##0;"zero"', "zero", "zero"],
    [5, "00000", "00005", "00005"],
  ];
  for (const [value, format, en, de] of cases) {
    assert.equal(formatValue(value, format, { locale: "en-US" }).text, en, `${format} en-US`);
    assert.equal(formatValue(value, format, { locale: "de-DE" }).text, de, `${format} de-DE`);
  }
  assert.equal(formatValue(-1234.5, "#,##0.00;[Red](#,##0.00)").color, "#ff0000");
  const accounting = formatValue(-1234.5, '_-* #,##0.00 €_-;-* #,##0.00 €_-;_-* "-"?? €_-;_-@_-', { locale: "de-DE" });
  assert.deepEqual([accounting.text, accounting.repeat], ["-1.234,50 € ", { index: 1, char: " " }]);
  assert.equal(formatValue("text", '"Note: "@').text, "Note: text");
  assert.equal(formatGeneral(1 / 3), "0.333333333");
  assert.equal(formatGeneral(123456789012345, "de-DE"), "1,23457E+14");
});

test("XLSX formula evaluator handles precedence, ranges, names and relative references", () => {
  const values = { "1:1": 4, "1:2": 6, "1:3": "x", "2:1": 10, "4:8": 5 };
  const context = (row, column) => ({
    sheet: "S", row, column,
    cell: (_sheet, r, c) => values[`${r}:${c}`] ?? null,
    name: name => ({ InPlan: "S!A$4=MEDIAN(S!A$4,S!$A1,S!$A1+S!$B1-1)" })[name],
  });
  const at = context(1, 1);
  assert.equal(evaluateFormula("-2^2", at), 4);
  assert.equal(evaluateFormula("2+3*4&\"!\"", at), "14!");
  assert.equal(evaluateFormula("SUM(A1:B2)", at), 20);
  assert.equal(evaluateFormula("MEDIAN(1,9,5)", at), 5);
  assert.equal(evaluateFormula("MOD(COLUMN(),2)=0", context(1, 4)), true);
  assert.equal(evaluateFormula("IF(C1=\"X\",\"yes\",\"no\")", at), "yes");
  assert.equal(evaluateFormula("COUNTIF(A1:B2,\">5\")", at), 2);
  // Conditional-format style evaluation: A1-relative formula shifted to the target cell.
  assert.equal(evaluateFormula("A1>5", context(2, 1), { row: 1, column: 1 }), true);
  // Defined names resolve their relative references against the cell being evaluated (H1 -> H$4 and $A1/$B1).
  assert.equal(evaluateFormula("InPlan", context(1, 8)), true);
  assert.equal(evaluateFormula("1/0", at).code, "#DIV/0!");
});

test("XLSX parser resolves cell styles, rich text, number formats and sheet layout", () => {
  const entries = {
    "xl/workbook.xml": workbookXml(`<sheet name="Data" r:id="rId1"/><sheet name="Secret" state="hidden" r:id="rId2"/>`, `<workbookPr date1904="0"/>`),
    "xl/_rels/workbook.xml.rels": rels([["rId1", "worksheets/sheet1.xml"], ["rId2", "worksheets/sheet2.xml"], ["rId3", "theme/theme1.xml", "theme"], ["rId4", "styles.xml", "styles"], ["rId5", "sharedStrings.xml", "sharedStrings"]]),
    "xl/theme/theme1.xml": theme,
    "xl/styles.xml": `<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00 &quot;EUR&quot;"/></numFmts>
      <fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><scheme val="minor"/></font><font><b/><i/><u val="double"/><sz val="14"/><color theme="1" tint="0.249977111117893"/><name val="Corbel"/></font></fonts>
      <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill></fills>
      <borders count="2"><border/><border diagonalDown="1"><left style="thin"><color indexed="64"/></left><bottom style="double"><color rgb="FFFF0000"/></bottom><diagonal style="thin"><color auto="1"/></diagonal></border></borders>
      <cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="164" fontId="1" fillId="2" borderId="1" applyNumberFormat="1"><alignment horizontal="center" vertical="top" wrapText="1" indent="1" textRotation="90"/></xf><xf numFmtId="14" fontId="0" fillId="0" borderId="0"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`,
    "xl/sharedStrings.xml": `<sst><si><t>Plain</t></si><si><r><t xml:space="preserve">Rich </t></r><r><rPr><b/><color rgb="FFFF0000"/><sz val="12"/><rFont val="Arial"/></rPr><t>red</t></r><rPh sb="0" eb="1"><t>ignored</t></rPh></si><si><t>Line_x000D_break</t></si></sst>`,
    "xl/worksheets/sheet1.xml": `<worksheet><sheetPr><tabColor rgb="FF00B050"/></sheetPr><sheetViews><sheetView showGridLines="0" zoomScale="80" workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" state="frozen"/></sheetView></sheetViews>
      <sheetFormatPr defaultRowHeight="15" baseColWidth="10"/><cols><col min="1" max="1" width="24.5703125" customWidth="1"/><col min="3" max="4" width="0" hidden="1"/></cols>
      <sheetData><row r="1" ht="30" customHeight="1"><c r="A1" s="1"><v>1234.5</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" s="2"><v>46323</v></c><c r="D1" s="3"><v>0.256</v></c></row><row r="2" hidden="1"><c r="A2" t="s"><v>2</v></c><c r="B2" t="b"><v>1</v></c><c r="C2" t="e"><v>#N/A</v></c><c r="D2"><f>A1*2</f></c></row></sheetData>
      <mergeCells count="1"><mergeCell ref="E1:F3"/></mergeCells></worksheet>`,
    "xl/worksheets/sheet2.xml": `<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`,
  };
  const model = parseXlsx(makeZip(entries), "styled.xlsx");
  const [sheet, hidden] = model.sheets;
  assert.equal(hidden.hidden, true);
  assert.equal(model.defaultFont.family, "Calibri");
  const cell = ref => sheet.cells.find(candidate => candidate.ref === ref);
  const style = model.styles[cell("A1").style];
  assert.deepEqual(style.font, { family: "Corbel", size: 14, bold: true, italic: true, underline: "double", color: "#404040" });
  assert.deepEqual(style.fill, { type: "pattern", pattern: "solid", foreground: "#dae3f3" });
  assert.deepEqual(style.borders, { left: { style: "thin", color: "#000000" }, bottom: { style: "double", color: "#ff0000" }, diagonal: { style: "thin", color: "#000000" }, diagonalDown: true });
  assert.deepEqual(style.alignment, { horizontal: "center", vertical: "top", wrapText: true, indent: 1, textRotation: 90 });
  assert.equal(cell("A1").display, "1,234.50 EUR");
  assert.equal(cell("C1").display, "10/28/2026");
  assert.equal(cell("D1").display, "26%");
  assert.equal(cell("B2").display, "TRUE");
  assert.equal(cell("C2").display, "#N/A");
  assert.equal(cell("A2").value, "Line\rbreak");
  assert.equal(cell("B1").value, "Rich red");
  assert.deepEqual(cell("B1").runs, [{ text: "Rich " }, { text: "red", font: { family: "Arial", size: 12, bold: true, color: "#ff0000" } }]);
  // Uncached formula results are computed.
  assert.deepEqual([cell("D2").value, cell("D2").display], ["2469", "2469"]);
  assert.equal(sheet.tabColor, "#00b050");
  assert.deepEqual(sheet.frozen, { rows: 2, columns: 1 });
  assert.equal(sheet.showGridLines, false);
  assert.equal(sheet.zoom, 0.8);
  assert.equal(sheet.defaultColumnWidth, -10);
  assert.deepEqual(sheet.columns, [{ min: 1, max: 1, width: 24.5703125 }, { min: 3, max: 4, width: 0, hidden: true }]);
  assert.deepEqual(sheet.rows, { 1: { height: 30 }, 2: { hidden: true } });
  assert.deepEqual(sheet.merges, [{ top: 1, left: 5, bottom: 3, right: 6 }]);
  // German regional settings change separators and the built-in short date.
  const german = parseXlsx(makeZip(entries), "styled.xlsx", { locale: "de-DE" }).sheets[0].cells;
  assert.deepEqual([german[0].display, german[2].display, german.find(c => c.ref === "B2").display], ["1.234,50 EUR", "28.10.2026", "WAHR"]);
});

test("XLSX parser evaluates conditional formats with defined names, priorities, and dxf fill semantics", () => {
  const entries = {
    "xl/workbook.xml": workbookXml(`<sheet name="Plan" r:id="rId1"/>`, `<definedNames><definedName name="InPlan">Plan!A$1=MEDIAN(Plan!A$1,Plan!$A1,Plan!$A1+Plan!$B1-1)</definedName></definedNames>`),
    "xl/_rels/workbook.xml.rels": rels([["rId1", "worksheets/sheet1.xml"], ["rId3", "theme/theme1.xml", "theme"], ["rId4", "styles.xml", "styles"]]),
    "xl/theme/theme1.xml": theme,
    "xl/styles.xml": `<styleSheet><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs>
      <dxfs count="3"><dxf><fill><patternFill patternType="solid"><fgColor auto="1"/><bgColor theme="7"/></patternFill></fill></dxf><dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="lightUp"><fgColor theme="7"/><bgColor auto="1"/></patternFill></fill><border><bottom style="thin"><color theme="0"/></bottom></border></dxf></dxfs></styleSheet>`,
    "xl/worksheets/sheet1.xml": `<worksheet><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c><c r="C1"><v>1</v></c><c r="D1"><v>2</v></c><c r="E1"><v>3</v></c><c r="F1"><v>4</v></c><c r="G1"><v>5</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><v>10</v></c><c r="D2"><v>50</v></c><c r="E2"><v>100</v></c></row></sheetData>
      <conditionalFormatting sqref="C1:G1"><cfRule type="expression" dxfId="0" priority="1" stopIfTrue="1"><formula>InPlan</formula></cfRule><cfRule type="expression" dxfId="2" priority="2"><formula>TRUE</formula></cfRule></conditionalFormatting>
      <conditionalFormatting sqref="C2:E2"><cfRule type="cellIs" dxfId="1" priority="3" operator="greaterThan"><formula>40</formula></cfRule><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFFFFFF"/><color rgb="FF000000"/></colorScale></cfRule><cfRule type="dataBar" priority="5"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting></worksheet>`,
  };
  const sheet = parseXlsx(makeZip(entries), "cf.xlsx").sheets[0];
  const cf = sheet.conditional;
  // Row 1: period numbers 1..5 in C1:G1, plan starts at 2 for 3 periods -> D1:F1 are in plan; dxf solid fills use bgColor.
  assert.deepEqual(cf.D1.style.fill, { type: "pattern", pattern: "solid", foreground: "#735773" });
  assert.equal(cf.D1.style.borders, undefined, "stopIfTrue prevents the lower-priority rule");
  assert.deepEqual(cf.C1.style.fill, { type: "pattern", pattern: "lightUp", foreground: "#735773", background: "#ffffff" });
  assert.deepEqual(cf.C1.style.borders.bottom, { style: "thin", color: "#ffffff" });
  assert.deepEqual(cf.G1.style.fill.pattern, "lightUp");
  // cellIs wins over the colour scale for fill; colour scale applies elsewhere; data bars scale min..max.
  assert.deepEqual(cf.E2.style.fill, { type: "pattern", pattern: "solid", foreground: "#ffc7ce" });
  assert.deepEqual(cf.E2.style.font, { bold: true, color: "#9c0006" });
  assert.equal(cf.C2.style.fill.foreground, "#ffffff");
  assert.equal(cf.C2.dataBar.end, 0.1);
  assert.equal(cf.E2.dataBar.end, 0.9);
});

test("XLSX parser applies built-in table styles, filter buttons, notes, links and drawings", () => {
  const png = Buffer.from("89504e470d0a1a0a0000000d4948445200000001000000010806000000", "hex");
  const entries = {
    "xl/workbook.xml": workbookXml(`<sheet name="T" r:id="rId1"/>`),
    "xl/_rels/workbook.xml.rels": rels([["rId1", "worksheets/sheet1.xml"], ["rId3", "theme/theme1.xml", "theme"]]),
    "xl/theme/theme1.xml": theme,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns:r="urn:r"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Q1</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>North</t></is></c><c r="B2"><v>120</v></c></row></sheetData>
      <hyperlinks><hyperlink ref="A2" r:id="rIdLink"/><hyperlink ref="B2" r:id="rIdBad"/></hyperlinks><drawing r:id="rIdDraw"/><tableParts count="1"><tablePart r:id="rIdTable"/></tableParts></worksheet>`,
    "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships><Relationship Id="rIdTable" Type="table" Target="../tables/table1.xml"/><Relationship Id="rIdLink" Type="hyperlink" Target="https://example.com/" TargetMode="External"/><Relationship Id="rIdBad" Type="hyperlink" Target="javascript:alert(1)" TargetMode="External"/><Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rIdDraw" Type="drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    "xl/tables/table1.xml": `<table ref="A1:B2" headerRowCount="1"><autoFilter ref="A1:B2"/><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>`,
    "xl/comments1.xml": `<comments><commentList><comment ref="B2"><text><r><t>Check this</t></r></text></comment></commentList></comments>`,
    "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="urn:r"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:blipFill><a:blip r:embed="rIdImg"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="476250"/><xdr:sp><xdr:spPr><a:prstGeom prst="roundRect"/><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></xdr:spPr><xdr:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1200" b="1"/><a:t>Note</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:oneCellAnchor></xdr:wsDr>`,
    "xl/drawings/_rels/drawing1.xml.rels": `<Relationships><Relationship Id="rIdImg" Type="image" Target="../media/image1.png"/></Relationships>`,
    "xl/media/image1.png": png,
  };
  const sheet = parseXlsx(makeZip(entries), "table.xlsx").sheets[0];
  const [table] = sheet.tables;
  assert.deepEqual(table.range, { top: 1, left: 1, bottom: 2, right: 2 });
  assert.deepEqual(table.elements.headerRow, { fill: { type: "pattern", pattern: "solid", foreground: "#4472c4" }, font: { color: "#ffffff", bold: true } });
  assert.deepEqual(table.elements.firstRowStripe.fill, { type: "pattern", pattern: "solid", foreground: "#dae3f3" });
  assert.deepEqual(sheet.filterButtons, ["A1", "B1"]);
  assert.deepEqual(sheet.hyperlinks, { A2: "https://example.com/" });
  assert.deepEqual(sheet.comments, { B2: "Check this" });
  assert.equal(sheet.drawings.length, 2);
  assert.deepEqual([sheet.drawings[0].kind, sheet.drawings[0].mimeType, sheet.drawings[0].anchor.to.column], ["image", "image/png", 5]);
  const [shape, text] = sheet.drawings[1].elements;
  assert.deepEqual([sheet.drawings[1].kind, sheet.drawings[1].frame], ["shapes", { left: 0, top: 0, width: 952500, height: 476250 }]);
  assert.deepEqual([shape.kind, shape.geometry, shape.fill, shape.width], ["shape", "roundRect", { color: "#2f5597" }, 952500]);
  assert.deepEqual([text.kind, text.verticalAlign, text.paragraphs[0].alignment, text.paragraphs[0].runs[0].text, text.paragraphs[0].runs[0].bold, text.paragraphs[0].runs[0].fontSize], ["text", "center", "center", "Note", true, 12]);
});

test("XLSX drawings keep groups, 3D cameras, text warps, reflections and hidden filter buttons", () => {
  const sp = (id, xfrm, body = "", props = "") => `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="S${id}"/></xdr:nvSpPr><xdr:spPr>${xfrm}${props}</xdr:spPr>${body}</xdr:sp>`;
  const xfrm = (x, y, cx, cy, rot = "") => `<a:xfrm${rot}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
  const fill = `<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:ln><a:noFill/></a:ln>`;
  const group = `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="1" name="G" descr="Timeline"/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="2000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></xdr:grpSpPr>`
    + sp(2, xfrm(0, 0, 500, 400), "", `<a:prstGeom prst="rect"/>${fill}<a:scene3d><a:camera prst="perspectiveRelaxed"><a:rot lat="17973601" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>`)
    + sp(3, xfrm(500, 0, 100, 100), "", `<a:prstGeom prst="ellipse"/>${fill}<a:scene3d><a:camera prst="perspectiveRelaxed"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>`)
    + sp(4, xfrm(0, 500, 400, 200), `<xdr:txBody><a:bodyPr anchor="ctr"><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp></a:bodyPr><a:p><a:pPr algn="ctr"/><a:fld id="{1}" type="TxLink"><a:rPr sz="1200"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Franklin Gothic Medium"/></a:rPr><a:t>Arched</a:t></a:fld></a:p></xdr:txBody>`, `<a:prstGeom prst="rect"/><a:noFill/>`)
    + sp(5, xfrm(500, 500, 400, 200, ` rot="20655491"`), `<xdr:txBody><a:bodyPr wrap="none"/><a:p><a:r><a:rPr sz="1400"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst><a:outerShdw blurRad="38100" dist="19050" dir="2700000"><a:schemeClr val="dk1"><a:alpha val="40000"/></a:schemeClr></a:outerShdw></a:effectLst></a:rPr><a:t>2027</a:t></a:r></a:p></xdr:txBody>`, `<a:prstGeom prst="rect"/><a:noFill/><a:effectLst><a:reflection blurRad="6350" stA="52000" endA="300" endPos="35000" dir="5400000" sy="-100000" algn="bl"/></a:effectLst>`)
    + `</xdr:grpSp>`;
  const entries = {
    "xl/workbook.xml": workbookXml(`<sheet name="T" r:id="rId1"/>`),
    "xl/_rels/workbook.xml.rels": rels([["rId1", "worksheets/sheet1.xml"], ["rId3", "theme/theme1.xml", "theme"]]),
    "xl/theme/theme1.xml": theme,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns:r="urn:r"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>A</t></is></c><c r="B1" t="inlineStr"><is><t>B</t></is></c></row></sheetData><drawing r:id="rIdDraw"/><tableParts count="1"><tablePart r:id="rIdTable"/></tableParts></worksheet>`,
    "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships><Relationship Id="rIdTable" Type="table" Target="../tables/table1.xml"/><Relationship Id="rIdDraw" Type="drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    "xl/tables/table1.xml": `<table ref="A1:B2" headerRowCount="1"><autoFilter ref="A1:B2"><filterColumn colId="1" hiddenButton="1"/></autoFilter><tableStyleInfo name="TableStyleLight1"/></table>`,
    "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>${group}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
  };
  const sheet = parseXlsx(makeZip(entries), "drawing.xlsx").sheets[0];
  assert.deepEqual(sheet.filterButtons, ["A1"]);
  const [drawing] = sheet.drawings;
  assert.deepEqual([drawing.kind, drawing.name, drawing.frame], ["shapes", "Timeline", { left: 100, top: 200, width: 2000, height: 1000 }]);
  const [swoosh, ellipse, arched, label] = drawing.elements;
  // Group children are scaled by ext/chExt and offset into the frame.
  assert.deepEqual([swoosh.left, swoosh.top, swoosh.width, swoosh.height, swoosh.fill.color], [100, 200, 1000, 400, "#595959"]);
  assert.equal(Math.round(swoosh.camera.latitude * 100) / 100, 299.56);
  assert.ok(swoosh.camera.distance > 0 && swoosh.camera.lit);
  // A preset camera without a rotation uses the preset's own tilt.
  assert.equal(ellipse.camera.latitude, 309.6);
  assert.deepEqual([arched.kind, arched.warp, arched.paragraphs[0].runs[0].text, arched.paragraphs[0].runs[0].fontFamily], ["text", "textArchUp", "Arched", "Franklin Gothic Medium"]);
  assert.deepEqual([label.rotation.toFixed(2), label.wrap, label.reflection.startOpacity, label.reflection.endPosition], ["344.26", false, 0.52, 0.35]);
  const shadow = label.paragraphs[0].runs[0].shadow;
  assert.deepEqual([shadow.x.toFixed(3), shadow.y.toFixed(3), shadow.blur, shadow.color, shadow.opacity], ["1.061", "1.061", 3, "#000000", 0.4]);
});

test("XLSX slicers, sparklines, TIFF pictures and input messages", () => {
  // 2x1 uncompressed 8-bit grayscale TIFF: black, white.
  const tags = [[256, 2], [257, 1], [258, 8], [259, 1], [262, 1], [273, 0], [277, 1], [278, 1], [279, 2]], dataOffset = 8 + 2 + tags.length * 12 + 4;
  const tiff = Buffer.alloc(dataOffset + 2);
  tiff.write("II", 0, "ascii"); tiff.writeUInt16LE(42, 2); tiff.writeUInt32LE(8, 4); tiff.writeUInt16LE(tags.length, 8);
  tags.forEach(([tag, value], index) => {
    const offset = 10 + index * 12, long = tag === 273 || tag === 279;
    tiff.writeUInt16LE(tag, offset); tiff.writeUInt16LE(long ? 4 : 3, offset + 2); tiff.writeUInt32LE(1, offset + 4);
    if (long) tiff.writeUInt32LE(tag === 273 ? dataOffset : value, offset + 8); else tiff.writeUInt16LE(value, offset + 8);
  });
  tiff[dataOffset + 1] = 255;
  const office = "http://schemas.microsoft.com/office/2007/relationships";
  const dxf = inner => `<dxf>${inner}</dxf>`, edges = color => `<border><left style="thin"><color rgb="${color}"/></left><right style="thin"><color rgb="${color}"/></right><top style="thin"><color rgb="${color}"/></top><bottom style="thin"><color rgb="${color}"/></bottom></border>`;
  const entries = {
    "xl/workbook.xml": workbookXml(`<sheet name="Data" r:id="rId1"/>`),
    "xl/_rels/workbook.xml.rels": `<Relationships><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="styles" Target="styles.xml"/><Relationship Id="rId3" Type="${office}/slicerCache" Target="slicerCaches/slicerCache1.xml"/></Relationships>`,
    "xl/styles.xml": `<styleSheet><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>`
      + `<dxfs count="2">${dxf(edges("FF00B050"))}${dxf(`<font><b/></font><border><bottom style="thin"><color rgb="FF00B050"/></bottom></border>`)}</dxfs>`
      + `<tableStyles count="1"><tableStyle name="Slicer S" pivot="0" table="0"><tableStyleElement type="wholeTable" dxfId="0"/><tableStyleElement type="headerRow" dxfId="1"/></tableStyle></tableStyles>`
      + `<extLst><ext uri="{46F421CA-312F-682f-3DD2-61675219B42D}"><x14:dxfs count="2">${dxf(`<font><color rgb="FFFFFFFF"/></font><fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/><bgColor rgb="FF006100"/></patternFill></fill>`)}${dxf(`<font><color rgb="FF000000"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFC0C0C0"/></patternFill></fill>`)}</x14:dxfs></ext>`
      + `<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"><x14:slicerStyle name="Slicer S"><x14:slicerStyleElements><x14:slicerStyleElement type="selectedItemWithData" dxfId="0"/><x14:slicerStyleElement type="unselectedItemWithData" dxfId="1"/></x14:slicerStyleElements></x14:slicerStyle></x14:slicerStyles></ext></extLst></styleSheet>`,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns:r="urn:r"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Wert</t></is></c></row>`
      + `<row r="2"><c r="A2" t="inlineStr"><is><t>Beta</t></is></c><c r="B2"><v>1</v></c></row><row r="3" hidden="1"><c r="A3" t="inlineStr"><is><t>Alpha</t></is></c><c r="B3"><v>2</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t>Beta</t></is></c><c r="B4"><v>3</v></c></row></sheetData>`
      + `<dataValidations count="1"><dataValidation allowBlank="1" showInputMessage="1" promptTitle="Hinweis" prompt="Namen eingeben" sqref="A1 A5:A6"/></dataValidations>`
      + `<drawing r:id="rIdDraw"/><tableParts count="1"><tablePart r:id="rIdTable"/></tableParts>`
      + `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups><x14:sparklineGroup type="column" displayEmptyCellsAs="gap" minAxisType="custom" manualMin="0" high="1"><x14:colorSeries rgb="FF376092"/><x14:colorHigh rgb="FFFF0000"/><x14:sparklines><x14:sparkline><xm:f>'Data'!B2:B4</xm:f><xm:sqref>C2</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst></worksheet>`,
    "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships><Relationship Id="rIdTable" Type="table" Target="../tables/table1.xml"/><Relationship Id="rIdDraw" Type="drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rIdSlicer" Type="${office}/slicer" Target="../slicers/slicer1.xml"/></Relationships>`,
    "xl/tables/table1.xml": `<table id="7" ref="A1:B4" headerRowCount="1"><autoFilter ref="A1:B4"><filterColumn colId="0"><filters><filter val="Beta"/></filters></filterColumn></autoFilter><tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Wert"/></tableColumns><tableStyleInfo name="TableStyleLight1"/></table>`,
    "xl/slicers/slicer1.xml": `<slicers><slicer name="Name" cache="Slicer_Name" caption="Namen" columnCount="2" style="Slicer S" rowHeight="241300"/></slicers>`,
    "xl/slicerCaches/slicerCache1.xml": `<slicerCacheDefinition name="Slicer_Name" sourceName="Name"><extLst><x:ext><x15:tableSlicerCache tableId="7" column="1"/></x:ext></extLst></slicerCacheDefinition>`,
    "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="urn:r"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`
      + `<mc:AlternateContent xmlns:mc="urn:mc"><mc:Choice Requires="sle15"><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Name"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="urn:sle" name="Name"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>`
      + `<mc:Fallback><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="0" name=""/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr><xdr:txBody><a:p><a:r><a:t>Not supported</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`
      + `<xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="190500" cy="95250"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="Bild"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImg"/></xdr:blipFill></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`,
    "xl/drawings/_rels/drawing1.xml.rels": `<Relationships><Relationship Id="rIdImg" Type="image" Target="../media/image1.tiff"/></Relationships>`,
    "xl/media/image1.tiff": tiff,
  };
  const sheet = parseXlsx(makeZip(entries), "slicer.xlsx").sheets[0];
  const [slicer, picture] = sheet.drawings;
  // Table slicer instead of the fallback text box: sorted items, selection from the table filter, the custom style's item and frame parts.
  assert.deepEqual([slicer.kind, slicer.caption, slicer.columnCount, slicer.itemHeight, slicer.filtered], ["slicer", "Namen", 2, 241300, true]);
  assert.deepEqual(slicer.items, [{ text: "Alpha", selected: false, hasData: true }, { text: "Beta", selected: true, hasData: true }]);
  assert.deepEqual([slicer.style.selectedItemWithData.fill.foreground, slicer.style.selectedItemWithData.font.color, slicer.style.unselectedItemWithData.fill.foreground], ["#006100", "#ffffff", "#c0c0c0"]);
  assert.deepEqual([slicer.style.header.font.bold, slicer.style.wholeTable.borders.left.color], [true, "#00b050"]);
  // TIFF pictures are converted to PNG.
  assert.equal(picture.mimeType, "image/png");
  const png = Buffer.from(picture.bytes), idat = png.indexOf("IDAT");
  assert.deepEqual([...inflateSync(png.subarray(idat + 4, idat + 4 + png.readUInt32BE(idat - 4)))], [0, 0, 0, 0, 255, 255, 255, 255, 255]);
  assert.deepEqual(sheet.sparklines, [{ ref: "C2", type: "column", values: [1, 2, 3], color: "#376092", lineWeight: 0.75, emptyCells: "gap", highColor: "#ff0000", min: 0 }]);
  assert.deepEqual(sheet.inputMessages, [{ ranges: [{ top: 1, left: 1, bottom: 1, right: 1 }, { top: 5, left: 1, bottom: 6, right: 1 }], title: "Hinweis", text: "Namen eingeben" }]);
});
