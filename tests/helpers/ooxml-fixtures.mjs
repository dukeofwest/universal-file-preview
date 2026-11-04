// Synthetic OOXML fixtures used by the M0 characterization snapshots.
//
// Each entry map is copied verbatim from tests/ooxml.test.mjs and tests/xlsx.test.mjs so the
// snapshots cover exactly the inputs the existing assertions exercise. The original test files
// are intentionally left unchanged; if a fixture is edited there, mirror the edit here.
import { makeZip } from "./zip.mjs";

// From tests/xlsx.test.mjs (module-level helpers).
const workbookXml = (sheets, extra = "") => `<workbook xmlns:r="urn:r"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${sheets}</sheets>${extra}</workbook>`;
const rels = ids => `<Relationships>${ids.map(([id, target, type = "worksheet"]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`).join("")}</Relationships>`;
const theme = `<a:theme xmlns:a="urn:a"><a:themeElements><a:clrScheme name="t"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="735773"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Corbel"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;

// ---------------------------------------------------------------------------------------------
// tests/ooxml.test.mjs
// ---------------------------------------------------------------------------------------------

const docxBasic = () => ({
  "word/document.xml": `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`,
});

const docxInvoice = () => ({
  "word/document.xml": `<w:document xmlns:w="urn:w" xmlns:r="urn:r"><w:body><w:sdt><w:sdtContent><w:p><w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:after="240"/></w:pPr><w:hyperlink r:id="external"><w:r><w:t>Invoice Summary</w:t></w:r></w:hyperlink></w:p></w:sdtContent></w:sdt><w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="8" w:color="586f80"/><w:insideH w:val="single" w:sz="4" w:color="aaaaaa"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:shd w:fill="586f80"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="FFFFFF"/></w:rPr><w:t>RECHNUNGS-NR. 1002</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>DATUM 10.11.22</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
  "word/styles.xml": `<w:styles xmlns:w="urn:w"><w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style></w:styles>`,
  "word/_rels/document.xml.rels": `<Relationships><Relationship Id="rIdHeader" Type="header" Target="header1.xml"/><Relationship Id="external" Target="https://example.test" TargetMode="External"/></Relationships>`,
  "word/header1.xml": `<w:hdr xmlns:w="urn:w"><w:sdt><w:sdtContent><w:p><w:r><w:t>Sono Jain, CPA</w:t></w:r></w:p></w:sdtContent></w:sdt></w:hdr>`,
});

const docxStyled = () => ({
  "word/document.xml": `<w:document xmlns:w="urn:w"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="BlueTable"/><w:tblW w:w="2500" w:type="pct"/><w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Header label</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Header value</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Body label</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Body value</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="600"/></w:sectPr></w:body></w:document>`,
  "word/styles.xml": `<w:styles xmlns:w="urn:w"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:sz w:val="20"/></w:rPr></w:style><w:style w:type="table" w:styleId="TableNormal"><w:tblPr><w:tblCellMar><w:top w:w="0"/><w:left w:w="108"/><w:bottom w:w="0"/><w:right w:w="108"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="BlueTable" w:basedOn="TableNormal"><w:rPr><w:color w:val="123456"/></w:rPr><w:tblPr><w:jc w:val="right"/><w:tblCellMar><w:left w:w="100"/></w:tblCellMar><w:tblBorders><w:insideH w:val="single" w:sz="4" w:color="aaaaaa"/></w:tblBorders></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:color w:val="FFFFFF"/><w:caps/></w:rPr><w:tcPr><w:shd w:fill="577188"/></w:tcPr></w:tblStylePr></w:style></w:styles>`,
  "word/theme/theme1.xml": `<a:theme xmlns:a="urn:a"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Cambria"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`,
});

const docxNumbered = () => ({
  "word/document.xml": `<w:document xmlns:w="urn:w"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>continued</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>`,
  "word/numbering.xml": `<w:numbering xmlns:w="urn:w"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1)"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
});

const xlsxCached = () => ({
  "xl/workbook.xml": `<workbook xmlns:r="urn:r"><sheets><sheet name="Main" r:id="rId1"/></sheets></workbook>`,
  "xl/_rels/workbook.xml.rels": `<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`,
  "xl/sharedStrings.xml": `<sst><si><t>Shared text</t></si></sst>`,
  "xl/worksheets/sheet1.xml": `<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>`,
});

const pptxOrder = () => ({
  "ppt/presentation.xml": `<p:presentation xmlns:r="urn:r"><p:sldIdLst><p:sldId r:id="rId1"/><p:sldId r:id="rId2"/></p:sldIdLst></p:presentation>`,
  "ppt/_rels/presentation.xml.rels": `<Relationships><Relationship Id="rId1" Target="slides/slide2.xml"/><Relationship Id="rId2" Target="slides/slide1.xml"/></Relationships>`,
  "ppt/slides/slide2.xml": `<p:sld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>First</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:sld>`,
  "ppt/slides/slide1.xml": `<p:sld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Second</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:sld>`,
});

const pptxTheme = () => {
  const pptTheme = `<a:theme><a:themeElements><a:clrScheme><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="FF0000"/></a:accent1></a:clrScheme>`
    + `<a:fmtScheme><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
  return {
    "ppt/presentation.xml": `<p:presentation xmlns:r="urn:r"><p:sldMasterIdLst><p:sldMasterId r:id="rId9"/></p:sldMasterIdLst><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": `<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId9" Target="slideMasters/slideMaster1.xml"/></Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster><p:cSld><p:spTree/></p:cSld><p:clrMap bg1="lt1" tx1="dk1"/></p:sldMaster>`,
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": `<Relationships><Relationship Id="rId1" Target="../theme/theme1.xml"/></Relationships>`,
    "ppt/theme/theme1.xml": pptTheme,
    "ppt/slides/slide1.xml": `<p:sld><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr></a:solidFill></p:bgPr></p:bg><p:spTree>`
      + `<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="50"/></a:xfrm><a:prstGeom prst="homePlate"><a:avLst><a:gd name="adj" fmla="val 40000"/></a:avLst></a:prstGeom></p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp>`
      + `<p:sp useBgFill="1"><p:spPr><a:xfrm><a:off x="10" y="10"/><a:ext cx="20" cy="20"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`
      + `</p:spTree></p:cSld></p:sld>`,
  };
};

// ---------------------------------------------------------------------------------------------
// tests/xlsx.test.mjs
// ---------------------------------------------------------------------------------------------

const xlsxStyled = () => ({
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
});

const xlsxConditional = () => ({
  "xl/workbook.xml": workbookXml(`<sheet name="Plan" r:id="rId1"/>`, `<definedNames><definedName name="InPlan">Plan!A$1=MEDIAN(Plan!A$1,Plan!$A1,Plan!$A1+Plan!$B1-1)</definedName></definedNames>`),
  "xl/_rels/workbook.xml.rels": rels([["rId1", "worksheets/sheet1.xml"], ["rId3", "theme/theme1.xml", "theme"], ["rId4", "styles.xml", "styles"]]),
  "xl/theme/theme1.xml": theme,
  "xl/styles.xml": `<styleSheet><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs>
      <dxfs count="3"><dxf><fill><patternFill patternType="solid"><fgColor auto="1"/><bgColor theme="7"/></patternFill></fill></dxf><dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="lightUp"><fgColor theme="7"/><bgColor auto="1"/></patternFill></fill><border><bottom style="thin"><color theme="0"/></bottom></border></dxf></dxfs></styleSheet>`,
  "xl/worksheets/sheet1.xml": `<worksheet><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c><c r="C1"><v>1</v></c><c r="D1"><v>2</v></c><c r="E1"><v>3</v></c><c r="F1"><v>4</v></c><c r="G1"><v>5</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><v>10</v></c><c r="D2"><v>50</v></c><c r="E2"><v>100</v></c></row></sheetData>
      <conditionalFormatting sqref="C1:G1"><cfRule type="expression" dxfId="0" priority="1" stopIfTrue="1"><formula>InPlan</formula></cfRule><cfRule type="expression" dxfId="2" priority="2"><formula>TRUE</formula></cfRule></conditionalFormatting>
      <conditionalFormatting sqref="C2:E2"><cfRule type="cellIs" dxfId="1" priority="3" operator="greaterThan"><formula>40</formula></cfRule><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFFFFFF"/><color rgb="FF000000"/></colorScale></cfRule><cfRule type="dataBar" priority="5"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting></worksheet>`,
});

const xlsxTable = () => {
  const png = Buffer.from("89504e470d0a1a0a0000000d4948445200000001000000010806000000", "hex");
  return {
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
};

const xlsxDrawing = () => {
  const sp = (id, xfrm, body = "", props = "") => `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="S${id}"/></xdr:nvSpPr><xdr:spPr>${xfrm}${props}</xdr:spPr>${body}</xdr:sp>`;
  const xfrm = (x, y, cx, cy, rot = "") => `<a:xfrm${rot}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
  const fill = `<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:ln><a:noFill/></a:ln>`;
  const group = `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="1" name="G" descr="Timeline"/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="2000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></xdr:grpSpPr>`
    + sp(2, xfrm(0, 0, 500, 400), "", `<a:prstGeom prst="rect"/>${fill}<a:scene3d><a:camera prst="perspectiveRelaxed"><a:rot lat="17973601" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>`)
    + sp(3, xfrm(500, 0, 100, 100), "", `<a:prstGeom prst="ellipse"/>${fill}<a:scene3d><a:camera prst="perspectiveRelaxed"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>`)
    + sp(4, xfrm(0, 500, 400, 200), `<xdr:txBody><a:bodyPr anchor="ctr"><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp></a:bodyPr><a:p><a:pPr algn="ctr"/><a:fld id="{1}" type="TxLink"><a:rPr sz="1200"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Franklin Gothic Medium"/></a:rPr><a:t>Arched</a:t></a:fld></a:p></xdr:txBody>`, `<a:prstGeom prst="rect"/><a:noFill/>`)
    + sp(5, xfrm(500, 500, 400, 200, ` rot="20655491"`), `<xdr:txBody><a:bodyPr wrap="none"/><a:p><a:r><a:rPr sz="1400"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst><a:outerShdw blurRad="38100" dist="19050" dir="2700000"><a:schemeClr val="dk1"><a:alpha val="40000"/></a:schemeClr></a:outerShdw></a:effectLst></a:rPr><a:t>2027</a:t></a:r></a:p></xdr:txBody>`, `<a:prstGeom prst="rect"/><a:noFill/><a:effectLst><a:reflection blurRad="6350" stA="52000" endA="300" endPos="35000" dir="5400000" sy="-100000" algn="bl"/></a:effectLst>`)
    + `</xdr:grpSp>`;
  return {
    "xl/workbook.xml": workbookXml(`<sheet name="T" r:id="rId1"/>`),
    "xl/_rels/workbook.xml.rels": rels([["rId1", "worksheets/sheet1.xml"], ["rId3", "theme/theme1.xml", "theme"]]),
    "xl/theme/theme1.xml": theme,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns:r="urn:r"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>A</t></is></c><c r="B1" t="inlineStr"><is><t>B</t></is></c></row></sheetData><drawing r:id="rIdDraw"/><tableParts count="1"><tablePart r:id="rIdTable"/></tableParts></worksheet>`,
    "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships><Relationship Id="rIdTable" Type="table" Target="../tables/table1.xml"/><Relationship Id="rIdDraw" Type="drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    "xl/tables/table1.xml": `<table ref="A1:B2" headerRowCount="1"><autoFilter ref="A1:B2"><filterColumn colId="1" hiddenButton="1"/></autoFilter><tableStyleInfo name="TableStyleLight1"/></table>`,
    "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>${group}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
  };
};

const xlsxSlicer = () => {
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
  return {
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
};

/**
 * Every synthetic OOXML fixture in the existing suite. `kind` selects the parser;
 * `options` is passed to parseXlsx exactly as the original test does (or `undefined`).
 */
export const ooxmlFixtures = [
  { id: "docx-basic", source: "tests/ooxml.test.mjs", kind: "docx", file: "sample.docx", entries: docxBasic },
  { id: "docx-invoice", source: "tests/ooxml.test.mjs", kind: "docx", file: "invoice.docx", entries: docxInvoice },
  { id: "docx-styled", source: "tests/ooxml.test.mjs", kind: "docx", file: "styled.docx", entries: docxStyled },
  { id: "docx-numbered", source: "tests/ooxml.test.mjs", kind: "docx", file: "numbered.docx", entries: docxNumbered },
  { id: "xlsx-cached", source: "tests/ooxml.test.mjs", kind: "xlsx", file: "sample.xlsx", entries: xlsxCached },
  { id: "pptx-order", source: "tests/ooxml.test.mjs", kind: "pptx", file: "sample.pptx", entries: pptxOrder },
  { id: "pptx-theme", source: "tests/ooxml.test.mjs", kind: "pptx", file: "styled.pptx", entries: pptxTheme },
  { id: "xlsx-styled", source: "tests/xlsx.test.mjs", kind: "xlsx", file: "styled.xlsx", entries: xlsxStyled },
  { id: "xlsx-styled-de", source: "tests/xlsx.test.mjs", kind: "xlsx", file: "styled.xlsx", entries: xlsxStyled, options: { locale: "de-DE" } },
  { id: "xlsx-conditional", source: "tests/xlsx.test.mjs", kind: "xlsx", file: "cf.xlsx", entries: xlsxConditional },
  { id: "xlsx-table", source: "tests/xlsx.test.mjs", kind: "xlsx", file: "table.xlsx", entries: xlsxTable },
  { id: "xlsx-drawing", source: "tests/xlsx.test.mjs", kind: "xlsx", file: "drawing.xlsx", entries: xlsxDrawing },
  { id: "xlsx-slicer", source: "tests/xlsx.test.mjs", kind: "xlsx", file: "slicer.xlsx", entries: xlsxSlicer },
];

/** Builds the fixture's package bytes exactly as the original tests do (stored entries). */
export function buildFixture(fixture) { return makeZip(fixture.entries()); }
