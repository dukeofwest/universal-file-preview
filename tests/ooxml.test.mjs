import test from "node:test";
import assert from "node:assert/strict";
import { makeZip } from "./helpers/zip.mjs";
import { parseDocx, parseXlsx, parsePptx } from "../dist/parsers/ooxml.js";

test("DOCX parser builds paragraphs, basic text formatting, and tables", () => {
  const xml = `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`;
  const model = parseDocx(makeZip({ "word/document.xml": xml }), "sample.docx");
  assert.equal(model.blocks.length, 2);
  assert.equal(model.blocks[0].value.runs[0].text, "Hello");
  assert.equal(model.blocks[0].value.runs[0].bold, true);
  assert.equal(model.blocks[1].value.rows[0].cells[0].paragraphs[0].runs[0].text, "Cell");
});

test("DOCX parser reads wrapped text, inherited styles, page layout, and headers", () => {
  const entries = {
    "word/document.xml": `<w:document xmlns:w="urn:w" xmlns:r="urn:r"><w:body><w:sdt><w:sdtContent><w:p><w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:after="240"/></w:pPr><w:hyperlink r:id="external"><w:r><w:t>Invoice Summary</w:t></w:r></w:hyperlink></w:p></w:sdtContent></w:sdt><w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="8" w:color="586f80"/><w:insideH w:val="single" w:sz="4" w:color="aaaaaa"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:shd w:fill="586f80"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="FFFFFF"/></w:rPr><w:t>RECHNUNGS-NR. 1002</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>DATUM 10.11.22</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
    "word/styles.xml": `<w:styles xmlns:w="urn:w"><w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style></w:styles>`,
    "word/_rels/document.xml.rels": `<Relationships><Relationship Id="rIdHeader" Type="header" Target="header1.xml"/><Relationship Id="external" Target="https://example.test" TargetMode="External"/></Relationships>`,
    "word/header1.xml": `<w:hdr xmlns:w="urn:w"><w:sdt><w:sdtContent><w:p><w:r><w:t>Sono Jain, CPA</w:t></w:r></w:p></w:sdtContent></w:sdt></w:hdr>`,
  };
  const model = parseDocx(makeZip(entries), "invoice.docx");
  assert.equal(model.blocks.length, 2);
  assert.equal(model.blocks[0].value.runs[0].text, "Invoice Summary");
  assert.equal(model.blocks[0].value.runs[0].bold, true);
  assert.equal(model.blocks[0].value.runs[0].fontSize, 18);
  assert.equal(model.blocks[0].value.alignment, "center");
  assert.equal(model.blocks[0].value.spacingAfter, 12);
  assert.equal(model.blocks[1].value.rows[0].cells[0].paragraphs[0].runs[0].text, "RECHNUNGS-NR. 1002");
  assert.equal(model.blocks[1].value.rows[0].cells[0].background, "#586f80");
  assert.equal(model.blocks[1].value.rows[0].cells[0].width, 50);
  assert.equal(model.page.width, 612);
  assert.equal(model.headers[0].value.runs[0].text, "Sono Jain, CPA");
});

test("DOCX parser applies theme defaults and conditional table styles without losing table layout", () => {
  const entries = {
    "word/document.xml": `<w:document xmlns:w="urn:w"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="BlueTable"/><w:tblW w:w="2500" w:type="pct"/><w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Header label</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Header value</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Body label</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Body value</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="600"/></w:sectPr></w:body></w:document>`,
    "word/styles.xml": `<w:styles xmlns:w="urn:w"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:sz w:val="20"/></w:rPr></w:style><w:style w:type="table" w:styleId="TableNormal"><w:tblPr><w:tblCellMar><w:top w:w="0"/><w:left w:w="108"/><w:bottom w:w="0"/><w:right w:w="108"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="BlueTable" w:basedOn="TableNormal"><w:rPr><w:color w:val="123456"/></w:rPr><w:tblPr><w:jc w:val="right"/><w:tblCellMar><w:left w:w="100"/></w:tblCellMar><w:tblBorders><w:insideH w:val="single" w:sz="4" w:color="aaaaaa"/></w:tblBorders></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:color w:val="FFFFFF"/><w:caps/></w:rPr><w:tcPr><w:shd w:fill="577188"/></w:tcPr></w:tblStylePr></w:style></w:styles>`,
    "word/theme/theme1.xml": `<a:theme xmlns:a="urn:a"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Cambria"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`,
  };
  const model = parseDocx(makeZip(entries), "styled.docx"), table = model.blocks[0].value;
  assert.equal(model.defaultRun.fontFamily, "Cambria");
  assert.equal(model.page.headerDistance, 36);
  assert.equal(model.page.footerDistance, 30);
  assert.equal(table.widthPercent, 50);
  assert.equal(table.alignment, "right");
  assert.equal(table.cellMargins.left, 5);
  assert.equal(table.rows[0].cells[0].background, "#577188");
  assert.equal(table.rows[0].cells[0].paragraphs[0].runs[0].caps, true);
  assert.equal(table.rows[0].cells[0].paragraphs[0].runs[0].fontSize, 10);
  assert.equal(table.rows[0].cells[0].paragraphs[0].runs[0].color, "#FFFFFF");
  assert.equal(table.rows[1].cells[0].paragraphs[0].runs[0].color, "#123456");
});

test("DOCX parser resolves basic numbering and hard page breaks", () => {
  const entries = {
    "word/document.xml": `<w:document xmlns:w="urn:w"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>continued</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>`,
    "word/numbering.xml": `<w:numbering xmlns:w="urn:w"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1)"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  };
  const model = parseDocx(makeZip(entries), "numbered.docx");
  assert.equal(model.blocks[0].value.list.label, "a)");
  assert.equal(model.blocks[0].value.runs[1].pageBreak, true);
  assert.equal(model.blocks[1].value.list.label, "b)");
});

test("XLSX parser resolves worksheet relationship and cached values", () => {
  const entries = {
    "xl/workbook.xml": `<workbook xmlns:r="urn:r"><sheets><sheet name="Main" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`,
    "xl/sharedStrings.xml": `<sst><si><t>Shared text</t></si></sst>`,
    "xl/worksheets/sheet1.xml": `<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>`,
  };
  const model = parseXlsx(makeZip(entries), "sample.xlsx");
  assert.equal(model.sheets[0].name, "Main");
  assert.deepEqual(model.sheets[0].cells.map(cell => [cell.ref, cell.value, cell.formula]), [["A1", "Shared text", undefined], ["B1", "2", "1+1"]]);
});

test("PPTX parser follows presentation slide order", () => {
  const entries = {
    "ppt/presentation.xml": `<p:presentation xmlns:r="urn:r"><p:sldIdLst><p:sldId r:id="rId1"/><p:sldId r:id="rId2"/></p:sldIdLst></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": `<Relationships><Relationship Id="rId1" Target="slides/slide2.xml"/><Relationship Id="rId2" Target="slides/slide1.xml"/></Relationships>`,
    "ppt/slides/slide2.xml": `<p:sld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>First</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:sld>`,
    "ppt/slides/slide1.xml": `<p:sld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Second</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:sld>`,
  };
  const model = parsePptx(makeZip(entries), "sample.pptx");
  assert.deepEqual(model.slides.map(slide => slide.title), ["First", "Second"]);
});

test("PPTX parser resolves theme style references, background-filled shapes, and linear-light tints", () => {
  const theme = `<a:theme><a:themeElements><a:clrScheme><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="FF0000"/></a:accent1></a:clrScheme>`
    + `<a:fmtScheme><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
  const entries = {
    "ppt/presentation.xml": `<p:presentation xmlns:r="urn:r"><p:sldMasterIdLst><p:sldMasterId r:id="rId9"/></p:sldMasterIdLst><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": `<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId9" Target="slideMasters/slideMaster1.xml"/></Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster><p:cSld><p:spTree/></p:cSld><p:clrMap bg1="lt1" tx1="dk1"/></p:sldMaster>`,
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": `<Relationships><Relationship Id="rId1" Target="../theme/theme1.xml"/></Relationships>`,
    "ppt/theme/theme1.xml": theme,
    "ppt/slides/slide1.xml": `<p:sld><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr></a:solidFill></p:bgPr></p:bg><p:spTree>`
      + `<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="50"/></a:xfrm><a:prstGeom prst="homePlate"><a:avLst><a:gd name="adj" fmla="val 40000"/></a:avLst></a:prstGeom></p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp>`
      + `<p:sp useBgFill="1"><p:spPr><a:xfrm><a:off x="10" y="10"/><a:ext cx="20" cy="20"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`
      + `</p:spTree></p:cSld></p:sld>`,
  };
  const slide = parsePptx(makeZip(entries), "styled.pptx").slides[0];
  assert.deepEqual(slide.background, { color: "#bcbcbc" });
  const [styled, backgroundFilled] = slide.elements.filter(element => element.kind === "shape");
  assert.equal(styled.geometry, "homePlate");
  assert.deepEqual(styled.adjustments, { adj: 40000 });
  assert.deepEqual(styled.fill, { color: "#ff0000" });
  assert.equal(styled.line.color, "#bc0000");
  assert.equal(backgroundFilled.useBackgroundFill, true);
  assert.deepEqual(backgroundFilled.fill, { color: "#bcbcbc" });
});
