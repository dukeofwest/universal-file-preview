import { deflateRawSync } from "node:zlib";
import { mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

const contentTypes = overrides => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${overrides}
</Types>`;

const relationships = relationships => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`;

const rootOfficeRelationship = type => relationships(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${type}"/>`);

const docx = {
  "[Content_Types].xml": contentTypes(`
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`),
  "_rels/.rels": rootOfficeRelationship("word/document.xml"),
  "word/_rels/document.xml.rels": relationships(`
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`),
  "word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
      <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:pPr><w:spacing w:after="180"/></w:pPr><w:rPr><w:b/><w:color w:val="173B67"/><w:sz w:val="36"/></w:rPr></w:style>
      <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:spacing w:before="240" w:after="100"/></w:pPr><w:rPr><w:b/><w:color w:val="245EA8"/><w:sz w:val="28"/></w:rPr></w:style>
    </w:styles>`,
  "word/numbering.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    </w:numbering>`,
  "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Quarterly product brief</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:color w:val="68778B"/></w:rPr><w:t>A compact DOCX sample with headings, emphasis, spacing, and a table.</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Highlights</w:t></w:r></w:p>
      <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Files are parsed locally in your browser.</w:t></w:r></w:p>
      <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:rPr><w:b/><w:color w:val="245EA8"/></w:rPr><w:t>Formatting is part of the preview.</w:t></w:r></w:p>
      <w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>
        <w:top w:val="single" w:sz="6" w:color="B8C6D8"/><w:left w:val="single" w:sz="6" w:color="B8C6D8"/><w:bottom w:val="single" w:sz="6" w:color="B8C6D8"/><w:right w:val="single" w:sz="6" w:color="B8C6D8"/><w:insideH w:val="single" w:sz="4" w:color="D9E1EA"/><w:insideV w:val="single" w:sz="4" w:color="D9E1EA"/>
      </w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>
        <w:tr><w:tc><w:tcPr><w:shd w:fill="245EA8"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Area</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:shd w:fill="245EA8"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Status</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:tc><w:p><w:r><w:t>Image preview</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Ready to try</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:tc><w:p><w:r><w:t>Office documents</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Partial renderer support</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/></w:sectPr>
    </w:body></w:document>`,
};

const xlsx = {
  "[Content_Types].xml": contentTypes(`
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`),
  "_rels/.rels": rootOfficeRelationship("xl/workbook.xml"),
  "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Overview" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  "xl/_rels/workbook.xml.rels": relationships(`
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`),
  "xl/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts>
      <fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Aptos Display"/></font></fonts>
      <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF245EA8"/></patternFill></fill></fills>
      <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD9E1EA"/></left><right style="thin"><color rgb="FFD9E1EA"/></right><top style="thin"><color rgb="FFD9E1EA"/></top><bottom style="thin"><color rgb="FFD9E1EA"/></bottom><diagonal/></border></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/></cellXfs>
    </styleSheet>`,
  "xl/sharedStrings.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="8" uniqueCount="8">
      <si><t>Product</t></si><si><t>Units</t></si><si><t>Growth</t></si><si><t>Preview toolkit</t></si><si><t>1,240</t></si><si><t>Images</t></si><si><t>Office Open XML</t></si><si><t>Partial rendering</t></si>
    </sst>`,
  "xl/worksheets/sheet1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C4"/><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>
      <cols><col min="1" max="1" width="26" customWidth="1"/><col min="2" max="3" width="15" customWidth="1"/></cols>
      <sheetData>
        <row r="1" ht="25" customHeight="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="B1" t="s" s="1"><v>1</v></c><c r="C1" t="s" s="1"><v>2</v></c></row>
        <row r="2"><c r="A2" t="s"><v>3</v></c><c r="B2"><v>1240</v></c><c r="C2" s="2"><v>0.184</v></c></row>
        <row r="3"><c r="A3" t="s"><v>5</v></c><c r="B3"><v>3</v></c><c r="C3" s="2"><v>0.12</v></c></row>
        <row r="4"><c r="A4" t="s"><v>6</v></c><c r="B4" t="s"><v>7</v></c></row>
      </sheetData>
    </worksheet>`,
};

const pptx = {
  "[Content_Types].xml": contentTypes(`
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
    <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
    <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`),
  "_rels/.rels": rootOfficeRelationship("ppt/presentation.xml"),
  "ppt/presentation.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId2"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500" type="wide"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
  "ppt/_rels/presentation.xml.rels": relationships(`
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>`),
  "ppt/slideMasters/slideMaster1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" dk1="dk1" lt1="lt1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`,
  "ppt/slideMasters/_rels/slideMaster1.xml.rels": relationships(`
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>`),
  "ppt/slideLayouts/slideLayout1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></p:clrMapOvr></p:sldLayout>`,
  "ppt/slideLayouts/_rels/slideLayout1.xml.rels": relationships(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`),
  "ppt/theme/theme1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="UFP sample"><a:themeElements><a:clrScheme name="UFP"><a:dk1><a:srgbClr val="18324F"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="3C4858"/></a:dk2><a:lt2><a:srgbClr val="F3F6FA"/></a:lt2><a:accent1><a:srgbClr val="3478F6"/></a:accent1><a:accent2><a:srgbClr val="19A987"/></a:accent2><a:accent3><a:srgbClr val="F1A33C"/></a:accent3><a:accent4><a:srgbClr val="8B72D9"/></a:accent4><a:accent5><a:srgbClr val="50A6C2"/></a:accent5><a:accent6><a:srgbClr val="78A55A"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Aptos"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme><a:fmtScheme name="UFP"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
  "ppt/slides/_rels/slide1.xml.rels": relationships(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`),
  "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="F3F6FA"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="685800"/><a:ext cx="7772400" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:pPr algn="l"><a:lnSpc><a:spcPct val="110000"/></a:lnSpc></a:pPr><a:r><a:rPr lang="en-US" sz="3200" b="1"><a:solidFill><a:srgbClr val="18324F"/></a:solidFill><a:latin typeface="Aptos Display"/></a:rPr><a:t>A clearer view of your files</a:t></a:r><a:endParaRPr lang="en-US" sz="3200"/></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1905000"/><a:ext cx="6800000" cy="700000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:pPr/><a:r><a:rPr lang="en-US" sz="1800"><a:solidFill><a:srgbClr val="52657A"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>A sample slide with a theme, typography, shapes, and a clean layout.</a:t></a:r><a:endParaRPr lang="en-US" sz="1800"/></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="4" name="Accent bar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="3000000"/><a:ext cx="4000000" cy="250000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="3478F6"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="5" name="Details"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="3600000"/><a:ext cx="7000000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1500"><a:solidFill><a:srgbClr val="3478F6"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>Angular  ·  React  ·  Vue  ·  Browser-side previews</a:t></a:r><a:endParaRPr lang="en-US" sz="1500"/></a:p></p:txBody></p:sp>
    </p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
};

const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="675" viewBox="0 0 1200 675" role="img" aria-labelledby="title description">
  <title id="title">Universal File Preview sample</title><desc id="description">A blue and mint abstract file preview illustration.</desc>
  <defs><linearGradient id="bg" x1="0" x2="1" y1="0" y2="1"><stop stop-color="#eef5ff"/><stop offset="1" stop-color="#f1fbf7"/></linearGradient></defs>
  <rect width="1200" height="675" rx="34" fill="url(#bg)"/><circle cx="940" cy="180" r="210" fill="#d9e8ff"/><circle cx="1030" cy="210" r="100" fill="#c0f1df"/>
  <rect x="160" y="115" width="560" height="445" rx="28" fill="#fff" stroke="#d7e1ed" stroke-width="3"/><rect x="210" y="170" width="230" height="20" rx="10" fill="#a8b7c8"/><rect x="210" y="225" width="430" height="16" rx="8" fill="#e0e7ef"/><rect x="210" y="265" width="365" height="16" rx="8" fill="#e0e7ef"/>
  <rect x="210" y="330" width="460" height="150" rx="18" fill="#f5f8fc" stroke="#e1e8f0" stroke-width="2"/><rect x="236" y="355" width="190" height="34" rx="8" fill="#3478f6"/><rect x="450" y="355" width="190" height="34" rx="8" fill="#3478f6"/><rect x="236" y="402" width="190" height="52" rx="8" fill="#fff" stroke="#dfe7ef"/><rect x="450" y="402" width="190" height="52" rx="8" fill="#fff" stroke="#dfe7ef"/>
  <path d="M850 397h152v32H850z" fill="#3478f6"/><path d="M873 323h107v74H873z" fill="#fff" stroke="#c9d8ea" stroke-width="3"/><path d="M897 348h58M897 364h42" stroke="#8c9eb2" stroke-width="7" stroke-linecap="round"/><text x="160" y="615" fill="#40536a" font-family="Arial,sans-serif" font-size="24" font-weight="600">Sample SVG · decoded by the browser</text>
</svg>`;

// OpenDocument spreadsheet sample: EUR amounts and percentages with saved display text, and
// totals whose formula results are saved (shown as saved; never recalculated by the viewer).
const EURO = String.fromCharCode(0x20ac);
const odfNamespaces = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"', 'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"', 'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"', 'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"', 'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"',
  'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"',
].join(" ");
const odsText = (text, style = "") => `<table:table-cell${style ? ` table:style-name="${style}"` : ""} office:value-type="string" calcext:value-type="string"><text:p>${text}</text:p></table:table-cell>`;
const odsFloat = (value, text, style = "ce3", formula = "") => `<table:table-cell table:style-name="${style}"${formula ? ` table:formula="${formula}"` : ""} office:value-type="float" office:value="${value}" calcext:value-type="float"><text:p>${text}</text:p></table:table-cell>`;
const odsMoney = (value, text, style = "ce3", formula = "") => `<table:table-cell table:style-name="${style}"${formula ? ` table:formula="${formula}"` : ""} office:value-type="currency" office:currency="EUR" office:value="${value}" calcext:value-type="currency"><text:p>${text} ${EURO}</text:p></table:table-cell>`;
const odsPercent = (value, text, style = "ce3") => `<table:table-cell table:style-name="${style}" office:value-type="percentage" office:value="${value}" calcext:value-type="percentage"><text:p>${text}</text:p></table:table-cell>`;
const ods = {
  mimetype: "application/vnd.oasis.opendocument.spreadsheet",
  "META-INF/manifest.xml": `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">
  <manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`,
  "meta.xml": `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta ${odfNamespaces} office:version="1.3"><office:meta><meta:generator>UniversalFilePreview/demo</meta:generator></office:meta></office:document-meta>`,
  "styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${odfNamespaces} office:version="1.3">
  <office:font-face-decls><style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'"/></office:font-face-decls>
  <office:styles>
    <style:default-style style:family="table-cell"><style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt" fo:language="en" fo:country="GB"/></style:default-style>
    <style:style style:name="Default" style:family="table-cell"/>
  </office:styles>
</office:document-styles>`,
  "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${odfNamespaces} office:version="1.3">
  <office:automatic-styles>
    <style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="4.6cm"/></style:style>
    <style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="3.2cm"/></style:style>
    <style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.62cm"/></style:style>
    <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#2365c7" fo:border="0.06pt solid #1a4f9c" style:vertical-align="middle"/><style:text-properties fo:color="#ffffff" fo:font-weight="bold"/></style:style>
    <style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:border-bottom="0.06pt solid #c8cdd3" style:vertical-align="middle"/></style:style>
    <style:style style:name="ce3" style:family="table-cell" style:parent-style-name="ce2"/>
    <style:style style:name="ce4" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:border-top="1.5pt solid #20242a" fo:background-color="#eef3fb" style:vertical-align="middle"/><style:text-properties fo:font-weight="bold"/></style:style>
  </office:automatic-styles>
  <office:body><office:spreadsheet>
    <table:table table:name="Quarterly results">
      <table:table-column table:style-name="co1"/><table:table-column table:style-name="co2" table:number-columns-repeated="3"/>
      <table:table-row table:style-name="ro1">${odsText("Product", "ce1")}${odsText("Units", "ce1")}${odsText("Revenue", "ce1")}${odsText("Growth", "ce1")}</table:table-row>
      <table:table-row table:style-name="ro1">${odsText("Preview toolkit", "ce2")}${odsFloat(1240, "1,240")}${odsMoney(18600, "18,600.00")}${odsPercent(0.184, "18.4%")}</table:table-row>
      <table:table-row table:style-name="ro1">${odsText("Images", "ce2")}${odsFloat(310, "310")}${odsMoney(4650, "4,650.00")}${odsPercent(0.12, "12.0%")}</table:table-row>
      <table:table-row table:style-name="ro1">${odsText("OpenDocument", "ce2")}${odsFloat(95, "95")}${odsMoney(1425.5, "1,425.50")}${odsPercent(0.075, "7.5%")}</table:table-row>
      <table:table-row table:style-name="ro1">${odsText("Total", "ce4")}${odsFloat(1645, "1,645", "ce4", "of:=SUM([.B2:.B4])")}${odsMoney(24675.5, "24,675.50", "ce4", "of:=SUM([.C2:.C4])")}<table:table-cell table:style-name="ce4"/></table:table-row>
    </table:table>
    <table:table table:name="Notes">
      <table:table-column table:style-name="co1" table:number-columns-repeated="2"/>
      <table:table-row>${odsText("Prepared")}<table:table-cell office:value-type="date" office:date-value="2026-10-28" calcext:value-type="date"><text:p>28 Oct 2026</text:p></table:table-cell></table:table-row>
      <table:table-row>${odsText("Source")}${odsText("OpenDocument spreadsheet, rendered in the browser")}</table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>`,
};

function makeZip(entries) {
  const localParts = [];
  const centralParts = [];
  let localOffset = 0;
  for (const [name, value] of Object.entries(entries)) {
    const nameBytes = Buffer.from(name);
    const raw = Buffer.from(value);
    // OpenDocument requires its "mimetype" entry to be stored uncompressed (ODF 1.3 Part 2 §3.3).
    const method = name === "mimetype" ? 0 : 8;
    const compressed = method ? deflateRawSync(raw) : raw;
    const crc = crc32(raw);
    const local = Buffer.alloc(30 + nameBytes.length + compressed.length);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x800, 6);
    local.writeUInt16LE(method, 8); local.writeUInt32LE(crc, 14); local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(raw.length, 22); local.writeUInt16LE(nameBytes.length, 26); nameBytes.copy(local, 30); compressed.copy(local, 30 + nameBytes.length);
    localParts.push(local);
    const central = Buffer.alloc(46 + nameBytes.length);
    central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6); central.writeUInt16LE(0x800, 8);
    central.writeUInt16LE(method, 10); central.writeUInt32LE(crc, 16); central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(raw.length, 24); central.writeUInt16LE(nameBytes.length, 28); central.writeUInt32LE(localOffset, 42); nameBytes.copy(central, 46);
    centralParts.push(central);
    localOffset += local.length;
  }
  const body = Buffer.concat(localParts);
  const central = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(Object.keys(entries).length, 8); end.writeUInt16LE(Object.keys(entries).length, 10);
  end.writeUInt32LE(central.length, 12); end.writeUInt32LE(body.length, 16);
  return Buffer.concat([body, central, end]);
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

export async function writeDemoSamples(outputDirectory) {
  const directory = resolve(outputDirectory);
  await mkdir(directory, { recursive: true });
  await Promise.all([
    writeFile(resolve(directory, "sample-illustration.svg"), svg, "utf8"),
    writeFile(resolve(directory, "sample-report.docx"), makeZip(docx)),
    writeFile(resolve(directory, "sample-workbook.xlsx"), makeZip(xlsx)),
    writeFile(resolve(directory, "sample-deck.pptx"), makeZip(pptx)),
    writeFile(resolve(directory, "sample-spreadsheet.ods"), makeZip(ods)),
  ]);
}
