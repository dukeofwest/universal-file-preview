import { PreviewError } from "../errors.js";
import { ZipArchive } from "../archive/zip-reader.js";
import { parseXml, elements, firstElement, textContent, type XmlNode } from "../xml/xml-parser.js";
import type { Paragraph, Presentation, PptxElement, PptxFill, PptxImageElement, PptxLine, PptxShapeElement, PptxTableElement, PptxTextElement, PptxShadow, Table, TableBorder, TableCell, TableCellMargins, TextRun, WordBlock, WordDocument, WordDrawing, WordDrawingPath, WordPageLayout } from "../model.js";

export { parseXlsx } from "./xlsx.js";

const decodeXml = (archive: ZipArchive, path: string): XmlNode => parseXml(new TextDecoder("utf-8", { fatal: true }).decode(archive.read(path)));
const child = (node: XmlNode | undefined, name: string) => node ? firstElement(node, name) : undefined;
const attrs = (node: XmlNode | undefined) => node?.attributes ?? {};

export function parseDocx(bytes: Uint8Array, name: string): WordDocument {
  const zip = ZipArchive.open(bytes), document = decodeXml(zip, "word/document.xml");
  const body = child(document, "body");
  if (!body) throw new PreviewError("DOCX document body is missing.", "invalid-file");
  const relationships = parseRelationships(zip, "word/_rels/document.xml.rels", "word");
  const images = readWordImages(zip, relationships);
  const styleContext = readWordStyles(zip);
  const numbering = readWordNumbering(zip, styleContext.themeFonts);
  const context: WordParseContext = { relationships, images, styleContext, themeColors: readThemeColors(zip), tableStyles: styleContext.tableStyles, numbering, numberCounters: new Map() };
  const blocks = parseWordBlocks(body, context);
  const page = parsePageLayout(body);
  const sections = descendants(body, "sectPr"), section = sections.at(-1);
  const differentFirstPageHeaderFooter = sections.length === 1 && (onOff(child(section, "titlePg")) ?? false);
  const headerReferences = differentFirstPageHeaderFooter && section ? elements(section, "headerReference") : descendants(body, "headerReference");
  const footerReferences = differentFirstPageHeaderFooter && section ? elements(section, "footerReference") : descendants(body, "footerReference");
  const referenceIds = (references: XmlNode[], type?: string) => references
    .filter(reference => type === undefined || (attribute(reference, "type") ?? "default") === type)
    .map(reference => attribute(reference, "id"))
    .filter((id): id is string => !!id);
  const defaultHeaderIds = referenceIds(headerReferences, differentFirstPageHeaderFooter ? "default" : undefined);
  const defaultFooterIds = referenceIds(footerReferences, differentFirstPageHeaderFooter ? "default" : undefined);
  const headers = readHeaderFooter(zip, defaultHeaderIds, relationships, context);
  const footers = readHeaderFooter(zip, defaultFooterIds, relationships, context);
  const firstPageHeaders = differentFirstPageHeaderFooter ? readHeaderFooter(zip, referenceIds(headerReferences, "first"), relationships, context) : [];
  const firstPageFooters = differentFirstPageHeaderFooter ? readHeaderFooter(zip, referenceIds(footerReferences, "first"), relationships, context) : [];
  return {
    kind: "word", blocks, title: name,
    defaultRun: { ...styleContext.defaults.run, ...styleContext.defaultStyle.run },
    defaultParagraph: { ...styleContext.defaults.paragraph, ...styleContext.defaultStyle.paragraph },
    ...(page ? { page } : {}),
    ...(headers.length ? { headers } : {}), ...(footers.length ? { footers } : {}),
    ...(differentFirstPageHeaderFooter ? { differentFirstPageHeaderFooter: true, ...(firstPageHeaders.length ? { firstPageHeaders } : {}), ...(firstPageFooters.length ? { firstPageFooters } : {}) } : {}),
  };
}

type WordStyle = { paragraph: Partial<Paragraph>; run: Partial<TextRun>; basedOn?: string };
type WordTableStyleRegion = WordStyle & { background?: string; verticalAlign?: TableCell["verticalAlign"]; borders?: Table["borders"]; cellBorders?: Table["borders"]; margins?: TableCellMargins };
type WordTableStyle = WordTableStyleRegion & { alignment?: Table["alignment"]; cellMargins?: TableCellMargins; rowBandSize?: number; regions: Map<string, WordTableStyleRegion> };
type WordNumberingLevel = { format: string; text: string; indentLeft?: number; hanging?: number; tabPosition?: number; labelStyle?: Partial<TextRun> };
type WordParseContext = {
  relationships: Map<string, string>;
  images: Map<string, { bytes: Uint8Array; mimeType: string }>;
  styleContext: { defaults: WordStyle; styles: Map<string, WordStyle>; defaultStyle: WordStyle; tableStyles: Map<string, WordTableStyle>; themeFonts: Record<string, string> };
  themeColors: Record<string, string>;
  tableStyles: Map<string, WordTableStyle>;
  tableParagraph?: Partial<Paragraph>;
  tableRun?: Partial<TextRun>;
  numbering: Map<string, Map<number, WordNumberingLevel>>;
  numberCounters: Map<string, number[]>;
};

function parseWordBlocks(node: XmlNode, context: WordParseContext): WordBlock[] {
  const blocks: WordBlock[] = [];
  const visit = (part: XmlNode) => {
    for (const item of part.children) {
      if (typeof item === "string") continue;
      if (item.localName === "p") blocks.push({ type: "paragraph", value: parseWordParagraph(item, context) });
      else if (item.localName === "tbl") blocks.push({ type: "table", value: parseWordTable(item, context) });
      else visit(item); // content controls, custom XML, insertions, and similar wrappers
    }
  };
  visit(node);
  return blocks;
}

function parseWordParagraph(node: XmlNode, context: WordParseContext): Paragraph {
  const properties = child(node, "pPr");
  const style = attribute(child(properties ?? node, "pStyle"), "val");
  const styleProps = style ? resolveWordStyle(style, context.styleContext.styles) : context.styleContext.defaultStyle;
  const defaults = context.styleContext.defaults;
  const directParagraph = readParagraphProperties(properties);
  const paragraph = { ...defaults.paragraph, ...styleProps.paragraph, ...context.tableParagraph, ...directParagraph };
  const directRunDefault = readRunProperties(child(properties ?? node, "rPr"), context.styleContext.themeFonts);
  const baseRun = { ...defaults.run, ...styleProps.run, ...context.tableRun, ...directRunDefault };
  const runs: TextRun[] = [];
  const fieldState: WordFieldState = { instruction: "", active: false, readingResult: false, pageMarkerEmitted: false, simplePageField: false };
  const walk = (part: XmlNode) => {
    for (const item of part.children) {
      if (typeof item === "string") continue;
      if (item.localName === "r") { parseWordRun(item, baseRun, context, runs, fieldState); continue; }
      if (item.localName === "fldSimple") {
        const previousSimplePageField = fieldState.simplePageField;
        const previousMarkerEmitted = fieldState.pageMarkerEmitted;
        fieldState.simplePageField = /^\s*PAGE(?:\s|$)/i.test(attribute(item, "instr") ?? "");
        fieldState.pageMarkerEmitted = false;
        walk(item);
        fieldState.simplePageField = previousSimplePageField;
        fieldState.pageMarkerEmitted = previousMarkerEmitted;
        continue;
      }
      // Text boxes and wrapped runs are common in Word-authored business forms.
      if (item.localName !== "pPr" && item.localName !== "rPr") walk(item);
    }
  };
  walk(node);
  const numPr = child(properties ?? node, "numPr");
  const numId = attribute(child(numPr ?? node, "numId"), "val");
  const level = Number(attribute(child(numPr ?? node, "ilvl"), "val") ?? "0");
  const numbering = numId ? context.numbering.get(numId)?.get(Number.isFinite(level) ? level : 0) : undefined;
  let list: Paragraph["list"];
  if (numbering && numId) {
    const listLevel = Math.max(0, Math.min(8, level)), counters = context.numberCounters.get(numId) ?? [];
    counters.length = Math.max(counters.length, listLevel + 1);
    counters[listLevel] = (counters[listLevel] ?? 0) + 1;
    counters.length = listLevel + 1;
    context.numberCounters.set(numId, counters);
    const label = numbering.format.toLowerCase() === "bullet" ? numbering.text : numbering.text.replace(/%(\d+)/g, (_match, raw: string) => {
      const index = Number(raw) - 1, definition = context.numbering.get(numId)?.get(index);
      return formatListNumber(counters[index] ?? 1, definition?.format ?? numbering.format);
    });
    list = {
      label, level: listLevel, kind: /^(?:bullet|none)$/i.test(numbering.format) ? "bullet" : "ordered",
      ...(numbering.indentLeft !== undefined ? { indentLeft: numbering.indentLeft } : {}),
      ...(numbering.hanging !== undefined ? { hanging: numbering.hanging } : {}),
      ...(numbering.tabPosition !== undefined ? { tabPosition: numbering.tabPosition } : {}),
      ...(numbering.labelStyle ? { labelStyle: numbering.labelStyle } : {}),
    };
  }
  const result: Paragraph = { runs, defaultRun: baseRun, ...paragraph, ...(style ? { style } : {}), ...(style?.toLowerCase().startsWith("heading") ? { kind: "heading" } : {}), ...(list ? { list } : {}) };
  return result;
}
type WordFieldState = { instruction: string; active: boolean; readingResult: boolean; pageMarkerEmitted: boolean; simplePageField: boolean };

function parseWordRun(node: XmlNode, base: Partial<TextRun>, context: WordParseContext, output: TextRun[], fieldState: WordFieldState): void {
  const runStyle = attribute(child(child(node, "rPr") ?? node, "rStyle"), "val");
  const style = runStyle ? resolveWordStyle(runStyle, context.styleContext.styles).run : {};
    const format = { ...base, ...style, ...readRunProperties(child(node, "rPr"), context.styleContext.themeFonts) };
  const emit = (text: string) => output.push({ ...format, text });
  const visit = (part: XmlNode) => {
    for (const item of part.children) {
      if (typeof item === "string") continue;
      if (item.localName === "rPr" || item.localName === "delText") continue;
      if (item.localName === "fldChar") {
        const type = attribute(item, "fldCharType");
        if (type === "begin") { fieldState.instruction = ""; fieldState.active = true; fieldState.readingResult = false; fieldState.pageMarkerEmitted = false; }
        else if (type === "separate") fieldState.readingResult = true;
        else if (type === "end") { fieldState.active = false; fieldState.readingResult = false; fieldState.instruction = ""; fieldState.pageMarkerEmitted = false; }
      } else if (item.localName === "instrText") {
        if (fieldState.active && !fieldState.readingResult) fieldState.instruction += textContent(item);
      } else if (item.localName === "t") {
        const isPageField = fieldState.simplePageField || fieldState.active && fieldState.readingResult && /^\s*PAGE(?:\s|$)/i.test(fieldState.instruction);
        if (isPageField && !fieldState.pageMarkerEmitted) {
          output.push({ ...format, text: "", pageNumber: true });
          fieldState.pageMarkerEmitted = true;
        } else if (!isPageField) emit(textContent(item));
      }
      else if (item.localName === "tab" || item.localName === "ptab") emit("\t");
      else if (item.localName === "br" && attribute(item, "type") === "page") output.push({ ...format, text: "", pageBreak: true });
      else if (item.localName === "br" || item.localName === "cr") emit("\n");
      else if (item.localName === "lastRenderedPageBreak") output.push({ ...format, text: "", pageBreak: true });
      else if (item.localName === "drawing" || item.localName === "pict") {
        const embed = descendants(item, "blip").map(blip => attribute(blip, "embed")).find(Boolean)
          ?? descendants(item, "imagedata").map(image => attribute(image, "id")).find(Boolean);
        if (embed) {
          const image = context.images.get(embed), relationshipPath = context.relationships.get(embed);
          if (image) {
            const extents = descendants(item, "extent")[0], wp = Number(attribute(extents, "cx")), hp = Number(attribute(extents, "cy"));
            const alt = attribute(descendants(item, "docPr")[0], "descr") ?? attribute(descendants(item, "docPr")[0], "name");
            output.push({ ...format, text: "", image: { ...image, ...(alt ? { alt } : {}), ...(wp > 0 ? { width: wp / 9525 } : {}), ...(hp > 0 ? { height: hp / 9525 } : {}) } });
          } else if (relationshipPath) emit("[Image]");
        }
        const drawing = parseWordDrawing(item, context);
        if (drawing) output.push({ ...format, text: "", drawing });
        // Word text boxes contain their own paragraphs, but their text still belongs in the visual flow.
        for (const box of descendants(item, "txbxContent")) visit(box);
      } else if (item.localName === "fldSimple") {
        const previousSimplePageField = fieldState.simplePageField;
        const previousMarkerEmitted = fieldState.pageMarkerEmitted;
        fieldState.simplePageField = /^\s*PAGE(?:\s|$)/i.test(attribute(item, "instr") ?? "");
        fieldState.pageMarkerEmitted = false;
        visit(item);
        fieldState.simplePageField = previousSimplePageField;
        fieldState.pageMarkerEmitted = previousMarkerEmitted;
      } else visit(item);
    }
  };
  visit(node);
}

function parseWordDrawing(node: XmlNode, context: WordParseContext): WordDrawing | undefined {
  const anchor = descendants(node, "anchor")[0] ?? descendants(node, "inline")[0];
  if (!anchor) return undefined;
  const extent = child(anchor, "extent"), widthEmu = Number(attribute(extent, "cx")), heightEmu = Number(attribute(extent, "cy"));
  if (!(widthEmu > 0 && heightEmu > 0)) return undefined;
  const horizontal = child(anchor, "positionH"), vertical = child(anchor, "positionV");
  const horizontalOffset = child(horizontal, "posOffset"), verticalOffset = child(vertical, "posOffset");
  const left = horizontalOffset ? Number(textContent(horizontalOffset)) / 9525 : 0;
  const top = verticalOffset ? Number(textContent(verticalOffset)) / 9525 : 0;
  const group = descendants(anchor, "wgp")[0], groupTransform = child(child(group, "grpSpPr"), "xfrm");
  const groupOffset = child(groupTransform, "off"), groupExtent = child(groupTransform, "ext");
  const childOffset = child(groupTransform, "chOff"), childExtent = child(groupTransform, "chExt");
  const groupX = Number(attribute(groupOffset, "x") ?? 0), groupY = Number(attribute(groupOffset, "y") ?? 0);
  const scaleX = Number(attribute(groupExtent, "cx") ?? widthEmu) / Number(attribute(childExtent, "cx") ?? widthEmu);
  const scaleY = Number(attribute(groupExtent, "cy") ?? heightEmu) / Number(attribute(childExtent, "cy") ?? heightEmu);
  const childX = Number(attribute(childOffset, "x") ?? 0), childY = Number(attribute(childOffset, "y") ?? 0);
  const paths: WordDrawingPath[] = [];
  for (const shape of descendants(anchor, "wsp").concat(descendants(anchor, "sp"))) {
    const properties = child(shape, "spPr"), transform = child(properties, "xfrm"), offset = child(transform, "off"), shapeExtent = child(transform, "ext");
    const shapeX = Number(attribute(offset, "x") ?? 0), shapeY = Number(attribute(offset, "y") ?? 0);
    const shapeWidth = Number(attribute(shapeExtent, "cx") ?? widthEmu), shapeHeight = Number(attribute(shapeExtent, "cy") ?? heightEmu);
    const geometry = child(properties, "custGeom"), preset = child(properties, "prstGeom");
    const pathList = child(geometry, "pathLst"), shapePaths = pathList ? elements(pathList, "path") : [];
    if (shapePaths.length) {
      for (const path of shapePaths) {
        const pathWidth = Number(attribute(path, "w") ?? shapeWidth), pathHeight = Number(attribute(path, "h") ?? shapeHeight);
        const d = drawingPathData(path);
        if (!d || !(pathWidth > 0 && pathHeight > 0)) continue;
        const sx = shapeWidth / pathWidth * scaleX, sy = shapeHeight / pathHeight * scaleY;
        const tx = groupX + (shapeX - childX) * scaleX, ty = groupY + (shapeY - childY) * scaleY;
        const style = drawingShapeStyle(properties, context.themeColors);
        paths.push({ d, transform: `matrix(${sx},0,0,${sy},${tx},${ty})`, ...style });
      }
    } else if (preset) {
      const presetName = attribute(preset, "prst");
      const d = presetName === "ellipse" ? `M ${shapeWidth / 2} 0 A ${shapeWidth / 2} ${shapeHeight / 2} 0 1 1 ${shapeWidth / 2} ${shapeHeight} A ${shapeWidth / 2} ${shapeHeight / 2} 0 1 1 ${shapeWidth / 2} 0 Z`
        : presetName === "line" ? `M 0 0 L ${shapeWidth} ${shapeHeight}` : `M 0 0 H ${shapeWidth} V ${shapeHeight} H 0 Z`;
      const tx = groupX + (shapeX - childX) * scaleX, ty = groupY + (shapeY - childY) * scaleY;
      paths.push({ d, transform: `matrix(${scaleX},0,0,${scaleY},${tx},${ty})`, ...drawingShapeStyle(properties, context.themeColors) });
    }
  }
  if (!paths.length) return undefined;
  return { left, top, width: widthEmu / 9525, height: heightEmu / 9525, behindText: attribute(anchor, "behindDoc") === "1", paths };
}

function drawingPathData(path: XmlNode): string {
  const commands: string[] = [];
  let x = 0, y = 0, startX = 0, startY = 0;
  for (const command of path.children) {
    if (typeof command === "string") continue;
    const coordinates = descendants(command, "pt").map(point => [Number(attribute(point, "x") ?? 0), Number(attribute(point, "y") ?? 0)] as const);
    const points = coordinates.map(([px, py]) => `${px} ${py}`), last = coordinates.at(-1);
    if (command.localName === "moveTo" && points[0]) { commands.push(`M ${points[0]}`); [startX, startY] = coordinates[0]!; }
    else if (command.localName === "lnTo" && points[0]) commands.push(`L ${points[0]}`);
    else if (command.localName === "cubicBezTo" && points.length >= 3) commands.push(`C ${points.slice(0, 3).join(" ")}`);
    else if (command.localName === "quadBezTo" && points.length >= 2) commands.push(`Q ${points.slice(0, 2).join(" ")}`);
    else if (command.localName === "close") { commands.push("Z"); x = startX; y = startY; continue; }
    else if (command.localName === "arcTo") {
      // DrawingML arcs continue from the current point around an ellipse, with angles in 60,000ths of a degree.
      const wR = Number(attribute(command, "wR") ?? 0), hR = Number(attribute(command, "hR") ?? 0);
      const start = Number(attribute(command, "stAng") ?? 0) / 60_000 * Math.PI / 180, sweep = Number(attribute(command, "swAng") ?? 0) / 60_000 * Math.PI / 180;
      const centerX = x - wR * Math.cos(start), centerY = y - hR * Math.sin(start);
      const segments = Math.abs(sweep) >= Math.PI * 2 - 1e-6 ? 2 : 1;
      for (let segment = 1; segment <= segments; segment++) {
        const angle = start + sweep * segment / segments, endX = centerX + wR * Math.cos(angle), endY = centerY + hR * Math.sin(angle);
        commands.push(`A ${wR} ${hR} 0 ${Math.abs(sweep / segments) > Math.PI ? 1 : 0} ${sweep > 0 ? 1 : 0} ${endX} ${endY}`);
        x = endX; y = endY;
      }
      continue;
    }
    if (last) [x, y] = last;
  }
  return commands.join(" ");
}

function drawingShapeStyle(properties: XmlNode | undefined, themeColors: Record<string, string>): Pick<WordDrawingPath, "fill" | "stroke" | "strokeWidth"> {
  const fill = child(properties, "solidFill"), noFill = child(properties, "noFill"), line = child(properties, "ln");
  const fillColor = fill ? drawingColor(fill, themeColors) : undefined;
  const lineColor = line && !child(line, "noFill") ? drawingColor(child(line, "solidFill"), themeColors) : undefined;
  const lineWidth = Number(attribute(line, "w"));
  return { ...(fillColor ? { fill: fillColor } : noFill || fill ? { fill: "none" } : {}), ...(lineColor ? { stroke: lineColor, strokeWidth: lineWidth > 0 ? lineWidth / 9525 : 1 } : {}) };
}

function drawingColor(node: XmlNode | undefined, themeColors: Record<string, string>): string | undefined {
  if (!node) return undefined;
  const color = ["srgbClr", "sysClr", "schemeClr", "prstClr"].map(name => descendants(node, name)[0]).find(Boolean);
  if (!color) return undefined;
  const value = attribute(color, "val"), resolved = color.localName === "sysClr" ? attribute(color, "lastClr") : value;
  if (!resolved) return undefined;
  if (color.localName === "schemeClr") return themeColors[resolved];
  if (/^[0-9a-f]{6}$/i.test(resolved)) return `#${resolved}`;
  return wordColor(resolved);
}

function parseWordTable(node: XmlNode, context: WordParseContext): Table {
  const tblPr = child(node, "tblPr"), widthNode = child(tblPr ?? node, "tblW");
  const styleId = attribute(child(tblPr ?? node, "tblStyle"), "val");
  const style = resolveWordTableStyle(styleId, context.tableStyles);
  const look = child(tblPr ?? node, "tblLook");
  const widthValue = Number(attribute(widthNode, "w"));
  const widthPercent = widthNode && attribute(widthNode, "type") === "pct" && widthValue > 0 ? widthValue / 50 : undefined;
  const width = widthNode && attribute(widthNode, "type") !== "pct" && widthValue > 0 ? twipsToPoints(widthValue) : undefined;
  const columns = elements(child(node, "tblGrid") ?? node, "gridCol").map(col => Number(attribute(col, "w")));
  const tableWidth = columns.reduce((sum, value) => sum + (Number.isFinite(value) ? value : 0), 0);
  const rowNodes = descendantsWithoutNestedTables(node, "tr"), activeMerges = new Map<number, TableCell>();
  const rows = rowNodes.map((row, rowIndex) => {
    const trPr = child(row, "trPr"), height = child(trPr ?? row, "trHeight");
    const nextMerges = new Map<number, TableCell>(); let column = 0;
    const cells = directDescendantsWithoutNestedTables(row, "tc").flatMap((cell): TableCell[] => {
      const tcPr = child(cell, "tcPr"), cellWidth = Number(attribute(child(tcPr ?? cell, "tcW"), "w"));
      const start = column, span = Math.max(1, Math.min(64, Number(attribute(child(tcPr ?? cell, "gridSpan"), "val") ?? "1")));
      column += span;
      const merge = child(tcPr ?? cell, "vMerge"), mergeValue = attribute(merge, "val");
      if (merge && mergeValue !== "restart" && activeMerges.has(start)) {
        const mergedCell = activeMerges.get(start)!; mergedCell.rowSpan = (mergedCell.rowSpan ?? 1) + 1;
        for (let col = start; col < start + span; col++) nextMerges.set(col, mergedCell);
        return [];
      }
      const gridWidth = columns.slice(start, start + span).reduce((sum, value) => sum + (Number.isFinite(value) ? value : 0), 0);
      const widthValue = cellWidth > 0 && tableWidth > 0 ? cellWidth / tableWidth * 100 : gridWidth > 0 && tableWidth > 0 ? gridWidth / tableWidth * 100 : undefined;
      const lastColumn = columns.length > 0 && start + span >= columns.length;
      const firstRowEnabled = tableLookEnabled(look, "firstRow");
      const bandRowIndex = rowIndex - (firstRowEnabled ? 1 : 0);
      const bandSize = Math.max(1, style.rowBandSize ?? 1);
      const horizontalBand = tableLookBandingEnabled(look, "horizontal") && bandRowIndex >= 0
        ? style.regions.get(Math.floor(bandRowIndex / bandSize) % 2 === 0 ? "band1Horz" : "band2Horz")
        : undefined;
      const regions = [
        ...(horizontalBand ? [horizontalBand] : []),
        ...(rowIndex === 0 && firstRowEnabled ? [style.regions.get("firstRow")] : []),
        ...(rowIndex === rowNodes.length - 1 && tableLookEnabled(look, "lastRow") ? [style.regions.get("lastRow")] : []),
        ...(start === 0 && tableLookEnabled(look, "firstColumn") ? [style.regions.get("firstCol")] : []),
        ...(lastColumn && tableLookEnabled(look, "lastColumn") ? [style.regions.get("lastCol")] : []),
      ].filter((region): region is WordTableStyleRegion => !!region);
      const cellBackground = [...regions].reverse().map(region => region.background).find(Boolean) ?? style.background;
      const cellBorders = mergeTableBorders(style.cellBorders, ...regions.map(region => region.borders), readTableBorders(child(tcPr ?? cell, "tcBorders")));
      const cellMargins = { ...style.cellMargins, ...regions.reduce((margins, region) => ({ ...margins, ...region.margins }), {}), ...readTableCellMargins(child(tcPr ?? cell, "tcMar")) };
      const vertical = attribute(child(tcPr ?? cell, "vAlign"), "val");
      const tableParagraph = regions.reduce((properties, region) => ({ ...properties, ...region.paragraph }), style.paragraph);
      const tableRun = regions.reduce((properties, region) => ({ ...properties, ...region.run }), style.run);
      const cellContext = { ...context, tableParagraph, tableRun };
      const paragraphs = parseWordBlocks(cell, cellContext).filter((block): block is Extract<WordBlock, { type: "paragraph" }> => block.type === "paragraph").map(block => block.value);
      const directShading = child(tcPr ?? cell, "shd"), shading = attribute(directShading, "fill");
      const cellBordersWithDirect = cellBorders;
      const background = directShading
        ? shading && shading !== "auto" ? `#${shading}` : undefined
        : cellBackground;
      const tableCell: TableCell = { paragraphs, column: start, ...(span > 1 ? { colSpan: span } : {}), ...(merge ? { rowSpan: 1 } : {}), ...(widthValue ? { width: widthValue } : {}), ...(background ? { background } : {}), ...(vertical === "center" || vertical === "bottom" ? { verticalAlign: vertical } : style.verticalAlign ? { verticalAlign: style.verticalAlign } : {}), ...(Object.keys(cellBordersWithDirect).length ? { borders: cellBordersWithDirect } : {}), ...(Object.keys(cellMargins).length ? { margins: cellMargins } : {}) };
      if (merge) for (let col = start; col < start + span; col++) nextMerges.set(col, tableCell);
      return [tableCell];
    });
    activeMerges.clear(); for (const [col, merged] of nextMerges) activeMerges.set(col, merged);
    const rowHeight = Number(attribute(height, "val"));
    return { cells, ...(rowHeight > 0 ? { height: twipsToPoints(rowHeight) } : {}) };
  });
  const alignment = attribute(child(tblPr ?? node, "jc"), "val") ?? style.alignment;
  const borders = mergeTableBorders(style.borders, readTableBorders(child(tblPr ?? node, "tblBorders")));
  const cellMargins = { ...style.cellMargins, ...readWordTableCellMargins(child(tblPr ?? node, "tblCellMar")) };
  return { rows, ...(width ? { width } : {}), ...(widthPercent ? { widthPercent } : {}), ...(alignment === "center" || alignment === "right" ? { alignment } : {}), ...(Object.keys(borders).length ? { borders } : {}), ...(Object.keys(cellMargins).length ? { cellMargins } : {}) };
}

function readWordStyles(zip: ZipArchive): { defaults: WordStyle; styles: Map<string, WordStyle>; defaultStyle: WordStyle; tableStyles: Map<string, WordTableStyle>; themeFonts: Record<string, string> } {
  const defaults: WordStyle = { paragraph: {}, run: {} }, styles = new Map<string, WordStyle>(), tableStyles = new Map<string, WordTableStyle>();
  const themeFonts = readThemeFonts(zip);
  if (!zip.has("word/styles.xml")) return { defaults, styles, defaultStyle: { paragraph: {}, run: {} }, tableStyles, themeFonts };
  const root = decodeXml(zip, "word/styles.xml"), docDefaults = child(root, "docDefaults");
  Object.assign(defaults.paragraph, readParagraphProperties(child(child(docDefaults ?? root, "pPrDefault"), "pPr")));
  Object.assign(defaults.run, readRunProperties(child(child(docDefaults ?? root, "rPrDefault"), "rPr"), themeFonts));
  let defaultParagraphStyleId: string | undefined;
  for (const style of elements(root, "style")) {
    const id = attribute(style, "styleId"); if (!id) continue;
    const basedOn = attribute(child(style, "basedOn"), "val");
    const type = attribute(style, "type");
    if (type === "table") tableStyles.set(id, readWordTableStyle(style, themeFonts, basedOn));
    else {
      styles.set(id, { ...(basedOn ? { basedOn } : {}), paragraph: readParagraphProperties(child(style, "pPr")), run: readRunProperties(child(style, "rPr"), themeFonts) });
      if (type === "paragraph" && attribute(style, "default") === "1") defaultParagraphStyleId = id;
    }
  }
  return { defaults, styles, defaultStyle: resolveWordStyle(defaultParagraphStyleId, styles), tableStyles, themeFonts };
}

function readThemeFonts(zip: ZipArchive): Record<string, string> {
  if (!zip.has("word/theme/theme1.xml")) return {};
  const theme = decodeXml(zip, "word/theme/theme1.xml"), scheme = child(child(theme, "themeElements"), "fontScheme");
  const family = (group: string) => attribute(child(child(scheme, group), "latin"), "typeface");
  const major = family("majorFont"), minor = family("minorFont");
  return {
    majorAscii: major ?? "", majorHAnsi: major ?? "", majorEastAsia: attribute(child(child(scheme, "majorFont"), "ea"), "typeface") || major || "",
    majorCs: attribute(child(child(scheme, "majorFont"), "cs"), "typeface") || major || "",
    majorBidi: attribute(child(child(scheme, "majorFont"), "cs"), "typeface") || major || "",
    minorAscii: minor ?? "", minorHAnsi: minor ?? "", minorEastAsia: attribute(child(child(scheme, "minorFont"), "ea"), "typeface") || minor || "",
    minorCs: attribute(child(child(scheme, "minorFont"), "cs"), "typeface") || minor || "",
    minorBidi: attribute(child(child(scheme, "minorFont"), "cs"), "typeface") || minor || "",
  };
}

function readThemeColors(zip: ZipArchive): Record<string, string> {
  if (!zip.has("word/theme/theme1.xml")) return {};
  const theme = decodeXml(zip, "word/theme/theme1.xml"), scheme = child(child(theme, "themeElements"), "clrScheme");
  const colors: Record<string, string> = {};
  for (const entry of scheme?.children ?? []) {
    if (typeof entry === "string") continue;
    const color = ["srgbClr", "sysClr"].map(name => child(entry, name)).find(Boolean);
    const value = attribute(color, color?.localName === "sysClr" ? "lastClr" : "val");
    if (value && /^[0-9a-f]{6}$/i.test(value)) colors[entry.localName] = `#${value}`;
  }
  return colors;
}

function readWordTableStyle(node: XmlNode, themeFonts: Record<string, string>, basedOn?: string): WordTableStyle {
  const properties = child(node, "tblPr"), cellProperties = child(node, "tcPr"), vertical = attribute(child(cellProperties ?? node, "vAlign"), "val");
  const alignment = attribute(child(properties ?? node, "jc"), "val");
  const rawRowBandSize = Number(attribute(child(properties ?? node, "tblStyleRowBandSize"), "val"));
  const rowBandSize = Number.isInteger(rawRowBandSize) && rawRowBandSize > 0 ? rawRowBandSize : undefined;
  const regions = new Map<string, WordTableStyleRegion>();
  for (const conditional of elements(node, "tblStylePr")) {
    const type = attribute(conditional, "type"); if (!type) continue;
    const tcPr = child(conditional, "tcPr"), fill = attribute(child(tcPr ?? conditional, "shd"), "fill");
    const regionAlignment = attribute(child(conditional, "jc"), "val");
    const regionVertical = attribute(child(tcPr ?? conditional, "vAlign"), "val");
    const regionBorders = readTableBorders(child(tcPr ?? conditional, "tcBorders"));
    regions.set(type, {
      paragraph: readParagraphProperties(child(conditional, "pPr")),
      run: readRunProperties(child(conditional, "rPr"), themeFonts),
      ...(fill && fill !== "auto" ? { background: `#${fill}` } : {}),
      ...(regionVertical === "center" || regionVertical === "bottom" ? { verticalAlign: regionVertical } : {}),
      ...(Object.keys(regionBorders).length ? { borders: regionBorders } : {}),
      ...(Object.keys(readWordTableCellMargins(child(tcPr ?? conditional, "tcMar"))).length ? { margins: readWordTableCellMargins(child(tcPr ?? conditional, "tcMar")) } : {}),
      ...(regionAlignment === "center" || regionAlignment === "right" ? { paragraph: { ...readParagraphProperties(child(conditional, "pPr")), alignment: regionAlignment } } : {}),
    });
  }
  const fill = attribute(child(cellProperties ?? node, "shd"), "fill");
  const tableBorders = readTableBorders(child(properties ?? node, "tblBorders"));
  const cellBorders = readTableBorders(child(cellProperties ?? node, "tcBorders"));
  return {
    ...(basedOn ? { basedOn } : {}),
    paragraph: readParagraphProperties(child(node, "pPr")),
    run: readRunProperties(child(node, "rPr"), themeFonts),
    ...(fill && fill !== "auto" ? { background: `#${fill}` } : {}),
    ...(vertical === "center" || vertical === "bottom" ? { verticalAlign: vertical } : {}),
    ...(Object.keys(tableBorders).length ? { borders: tableBorders } : {}),
    ...(Object.keys(cellBorders).length ? { cellBorders } : {}),
    ...(Object.keys(readWordTableCellMargins(child(cellProperties ?? node, "tcMar"))).length ? { margins: readWordTableCellMargins(child(cellProperties ?? node, "tcMar")) } : {}),
    alignment: alignment === "center" || alignment === "right" ? alignment : undefined,
    cellMargins: readWordTableCellMargins(child(properties ?? node, "tblCellMar")),
    ...(rowBandSize ? { rowBandSize } : {}),
    regions,
  };
}

function resolveWordTableStyle(id: string | undefined, styles: Map<string, WordTableStyle>, seen = new Set<string>()): WordTableStyle {
  if (!id || seen.has(id)) return { paragraph: {}, run: {}, regions: new Map() };
  seen.add(id);
  const style = styles.get(id); if (!style) return { paragraph: {}, run: {}, regions: new Map() };
  const parent = resolveWordTableStyle(style.basedOn, styles, seen), regions = new Map(parent.regions);
  for (const [name, region] of style.regions) {
    const inherited = regions.get(name);
    regions.set(name, { ...inherited, ...region, paragraph: { ...inherited?.paragraph, ...region.paragraph }, run: { ...inherited?.run, ...region.run }, borders: mergeTableBorders(inherited?.borders, region.borders), margins: { ...inherited?.margins, ...region.margins } });
  }
  return {
    ...parent, ...style,
    paragraph: { ...parent.paragraph, ...style.paragraph }, run: { ...parent.run, ...style.run },
    borders: mergeTableBorders(parent.borders, style.borders), cellBorders: mergeTableBorders(parent.cellBorders, style.cellBorders), margins: { ...parent.margins, ...style.margins }, cellMargins: { ...parent.cellMargins, ...style.cellMargins },
    regions,
  };
}

function readWordNumbering(zip: ZipArchive, themeFonts: Record<string, string> = {}): Map<string, Map<number, WordNumberingLevel>> {
  const result = new Map<string, Map<number, WordNumberingLevel>>();
  if (!zip.has("word/numbering.xml")) return result;
  const root = decodeXml(zip, "word/numbering.xml"), abstract = new Map<string, Map<number, WordNumberingLevel>>();
  for (const definition of elements(root, "abstractNum")) {
    const id = attribute(definition, "abstractNumId"); if (!id) continue;
    const levels = new Map<number, WordNumberingLevel>();
    for (const level of elements(definition, "lvl")) {
      const index = Number(attribute(level, "ilvl") ?? "0"), format = attribute(child(level, "numFmt"), "val") ?? "decimal", text = attribute(child(level, "lvlText"), "val") ?? "%1.";
      const paragraphProperties = child(level, "pPr"), indent = child(paragraphProperties, "ind");
      const leftRaw = attribute(indent, "left") ?? attribute(indent, "start"), hangingRaw = attribute(indent, "hanging");
      const numberTab = elements(child(paragraphProperties, "tabs") ?? level, "tab").find(tab => attribute(tab, "val") === "num");
      const tabRaw = attribute(numberTab, "pos");
      const left = Number(leftRaw), hanging = Number(hangingRaw), tabPosition = Number(tabRaw);
      levels.set(index, {
        format, text: format.toLowerCase() === "bullet" ? normalizeWordBullet(text) : text,
        ...(leftRaw !== undefined && Number.isFinite(left) ? { indentLeft: twipsToPoints(left) } : {}),
        ...(hangingRaw !== undefined && Number.isFinite(hanging) ? { hanging: twipsToPoints(hanging) } : {}),
        ...(tabRaw !== undefined && Number.isFinite(tabPosition) ? { tabPosition: twipsToPoints(tabPosition) } : {}),
        ...(child(level, "rPr") ? { labelStyle: readRunProperties(child(level, "rPr"), themeFonts) } : {}),
      });
    }
    abstract.set(id, levels);
  }
  for (const num of elements(root, "num")) {
    const id = attribute(num, "numId"), abstractId = attribute(child(num, "abstractNumId"), "val");
    if (id && abstractId && abstract.has(abstractId)) result.set(id, abstract.get(abstractId)!);
  }
  return result;
}

function normalizeWordBullet(text: string): string {
  // Word commonly stores the round bullet as U+F0B7, a private-use glyph that
  // only appears correctly when the original Symbol/Wingdings font is present.
  return text.replace(/\uF0B7/g, "\u2022");
}

function resolveWordStyle(id: string | undefined, styles: Map<string, WordStyle>, seen = new Set<string>()): WordStyle {
  if (!id || seen.has(id)) return { paragraph: {}, run: {} };
  seen.add(id);
  const style = styles.get(id); if (!style) return { paragraph: {}, run: {} };
  const parent = resolveWordStyle(style.basedOn, styles, seen);
  return { paragraph: { ...parent.paragraph, ...style.paragraph }, run: { ...parent.run, ...style.run } };
}

function readParagraphProperties(node: XmlNode | undefined): Partial<Paragraph> {
  if (!node) return {};
  const alignmentValue = attribute(child(node, "jc"), "val");
  const alignment = alignmentValue === "both" ? "justify" : alignmentValue;
  const spacing = child(node, "spacing"), indent = child(node, "ind"), shading = child(node, "shd"), paragraphBorders = child(node, "pBdr");
  const beforeRaw = attribute(spacing, "before"), afterRaw = attribute(spacing, "after"), lineRaw = attribute(spacing, "line");
  const lineRuleValue = attribute(spacing, "lineRule");
  const leftRaw = attribute(indent, "left") ?? attribute(indent, "start"), rightRaw = attribute(indent, "right") ?? attribute(indent, "end");
  const firstLineRaw = attribute(indent, "firstLine"), hangingRaw = attribute(indent, "hanging");
  const before = Number(beforeRaw), after = Number(afterRaw), line = Number(lineRaw), left = Number(leftRaw), right = Number(rightRaw), firstLine = Number(firstLineRaw), hanging = Number(hangingRaw);
  const outlineLevel = Number(attribute(child(node, "outlineLvl"), "val"));
  const fill = attribute(shading, "fill");
  const borderSpacing: NonNullable<Paragraph["borderSpacing"]> = {};
  const paragraphBorderEdges = paragraphBorders ? elements(paragraphBorders, "top").concat(elements(paragraphBorders, "left"), elements(paragraphBorders, "bottom"), elements(paragraphBorders, "right"), elements(paragraphBorders, "between")) : [];
  for (const edge of paragraphBorderEdges) {
    const raw = attribute(edge, "space"), value = Number(raw);
    if (raw !== undefined && Number.isFinite(value)) borderSpacing[edge.localName as keyof NonNullable<Paragraph["borderSpacing"]>] = value;
  }
  return {
    ...(alignment === "start" ? { alignment: "left" as const } : alignment === "end" ? { alignment: "right" as const } : alignment === "left" || alignment === "right" || alignment === "center" || alignment === "justify" ? { alignment } : {}),
    ...(beforeRaw !== undefined && Number.isFinite(before) ? { spacingBefore: twipsToPoints(before) } : {}),
    ...(afterRaw !== undefined && Number.isFinite(after) ? { spacingAfter: twipsToPoints(after) } : {}),
    ...(line > 0 ? { lineHeight: lineRuleValue === "exact" || lineRuleValue === "atLeast" ? twipsToPoints(line) : line / 240, lineHeightRule: lineRuleValue === "exact" || lineRuleValue === "atLeast" ? lineRuleValue : "auto" } : {}),
    ...(leftRaw !== undefined && Number.isFinite(left) ? { indentLeft: twipsToPoints(left) } : {}), ...(rightRaw !== undefined && Number.isFinite(right) ? { indentRight: twipsToPoints(right) } : {}),
    ...(firstLineRaw !== undefined && Number.isFinite(firstLine) ? { indentFirstLine: twipsToPoints(firstLine) } : hangingRaw !== undefined && Number.isFinite(hanging) ? { indentFirstLine: -twipsToPoints(hanging) } : {}),
    ...(child(node, "pageBreakBefore") ? { pageBreakBefore: true } : {}),
    ...(Number.isInteger(outlineLevel) && outlineLevel >= 0 && outlineLevel <= 5 ? { kind: "heading" as const } : {}),
    ...(fill && fill !== "auto" && /^[0-9a-f]{6}$/i.test(fill) ? { background: `#${fill}` } : {}),
    ...(paragraphBorders ? { borders: readParagraphBorders(paragraphBorders) } : {}),
    ...(Object.keys(borderSpacing).length ? { borderSpacing } : {}),
  };
}

function readParagraphBorders(node: XmlNode): NonNullable<Paragraph["borders"]> {
  const borders: NonNullable<Paragraph["borders"]> = {};
  for (const edge of elements(node, "top").concat(elements(node, "left"), elements(node, "bottom"), elements(node, "right"), elements(node, "between"))) {
    const name = edge.localName as keyof NonNullable<Paragraph["borders"]>;
    const value = attribute(edge, "val") ?? "single", color = attribute(edge, "color") ?? "808080", size = Number(attribute(edge, "sz"));
    borders[name] = { color: value === "nil" || value === "none" ? "transparent" : color === "auto" ? "#808080" : `#${color}`, width: value === "nil" || value === "none" ? 0 : Math.max(0.25, (Number.isFinite(size) ? size : 4) / 8), style: value === "nil" || value === "none" ? "none" : value === "dashed" || value === "dotted" || value === "double" ? value : "solid" };
  }
  return borders;
}

function readRunProperties(node: XmlNode | undefined, themeFonts: Record<string, string> = {}): Partial<TextRun> {
  if (!node) return {};
  const fonts = child(node, "rFonts"), color = attribute(child(node, "color"), "val"), highlight = attribute(child(node, "highlight"), "val"), underline = child(node, "u");
  const size = Number(attribute(child(node, "sz"), "val")), complexScriptSize = Number(attribute(child(node, "szCs"), "val"));
  const family = attribute(fonts, "ascii") || attribute(fonts, "hAnsi") || attribute(fonts, "eastAsia")
    || attribute(fonts, "cs") || themeFonts[attribute(fonts, "asciiTheme") ?? ""] || themeFonts[attribute(fonts, "hAnsiTheme") ?? ""] || themeFonts[attribute(fonts, "eastAsiaTheme") ?? ""] || themeFonts[attribute(fonts, "csTheme") ?? ""];
  const vert = attribute(child(node, "vertAlign"), "val"), spacing = Number(attribute(child(node, "spacing"), "val")), position = Number(attribute(child(node, "position"), "val"));
  const bold = onOff(child(node, "b")), italic = onOff(child(node, "i")), strike = onOff(child(node, "strike")), caps = onOff(child(node, "caps")), smallCaps = onOff(child(node, "smallCaps"));
  const underlineValue = (attribute(underline, "val") ?? "single").toLowerCase(), underlineColor = attribute(underline, "color");
  const runShading = child(node, "shd"), runFill = attribute(runShading, "fill");
  return {
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(underline ? { underline: !["none", "nil"].includes(underlineValue), ...(underlineValue === "double" || underlineValue === "dotted" || underlineValue === "dashed" || underlineValue === "wave" ? { underlineStyle: underlineValue === "wave" ? "wavy" as const : underlineValue } : { underlineStyle: "solid" as const }), ...(underlineColor && /^[0-9a-f]{6}$/i.test(underlineColor) ? { underlineColor: `#${underlineColor}` } : {}) } : {}),
    ...(strike !== undefined ? { strike } : {}),
    ...(caps !== undefined ? { caps } : {}), ...(smallCaps !== undefined ? { smallCaps } : {}),
    ...(size > 0 ? { fontSize: size / 2 } : {}), ...(complexScriptSize > 0 ? { complexScriptFontSize: complexScriptSize / 2 } : {}), ...(family ? { fontFamily: family } : {}),
    ...(Number.isFinite(spacing) && spacing !== 0 ? { letterSpacing: spacing / 20 } : {}),
    ...(Number.isFinite(position) && position !== 0 ? { position: position / 2 } : {}),
    ...(color && color !== "auto" && /^[0-9a-f]{6}$/i.test(color) ? { color: `#${color}` } : {}),
    ...(highlight && highlight !== "none" ? { highlight: wordColor(highlight) } : {}),
    ...(runFill && runFill !== "auto" && /^[0-9a-f]{6}$/i.test(runFill) ? { highlight: `#${runFill}` } : {}),
    ...(vert === "superscript" || vert === "subscript" ? { verticalAlign: vert } : {}),
  };
}

function readTableBorders(node: XmlNode | undefined): NonNullable<Table["borders"]> {
  if (!node) return {};
  const borders: NonNullable<Table["borders"]> = {};
  for (const edge of elements(node, "top").concat(elements(node, "left"), elements(node, "bottom"), elements(node, "right"), elements(node, "insideH"), elements(node, "insideV"))) {
    const name = edge.localName as keyof NonNullable<Table["borders"]>, value = attribute(edge, "val") ?? "single", color = attribute(edge, "color") ?? "808080", size = Number(attribute(edge, "sz"));
    borders[name] = { color: value === "nil" || value === "none" ? "transparent" : color === "auto" ? "#808080" : `#${color}`, width: value === "nil" || value === "none" ? 0 : Math.max(0.25, (Number.isFinite(size) ? size : 4) / 8), style: value === "nil" || value === "none" ? "none" : value === "dashed" || value === "dotted" || value === "double" ? value : "solid" };
  }
  return borders;
}

function mergeTableBorders(...sets: Array<Table["borders"] | undefined>): NonNullable<Table["borders"]> {
  return sets.reduce<NonNullable<Table["borders"]>>((result, set) => Object.assign(result, set), {});
}

function readWordTableCellMargins(node: XmlNode | undefined): TableCellMargins {
  if (!node) return {};
  const margins: TableCellMargins = {};
  for (const [edge, property] of [["top", "top"], ["right", "right"], ["bottom", "bottom"], ["left", "left"]] as const) {
    const raw = attribute(child(node, property), "w"), value = Number(raw);
    if (raw !== undefined && Number.isFinite(value)) margins[edge] = twipsToPoints(value);
  }
  return margins;
}

function readTableCellMargins(node: XmlNode | undefined): TableCellMargins {
  return readWordTableCellMargins(node);
}

function tableLookEnabled(look: XmlNode | undefined, region: "firstRow" | "lastRow" | "firstColumn" | "lastColumn"): boolean {
  const property = region === "firstRow" ? "firstRow" : region === "lastRow" ? "lastRow" : region === "firstColumn" ? "firstColumn" : "lastColumn";
  const value = attribute(look, property);
  if (value !== undefined) return value !== "0" && value.toLowerCase() !== "false";
  const lookValue = Number.parseInt(attribute(look, "val") ?? "", 16);
  if (!Number.isFinite(lookValue)) return true;
  const bit = region === "firstRow" ? 0x20 : region === "lastRow" ? 0x40 : region === "firstColumn" ? 0x80 : 0x100;
  return (lookValue & bit) !== 0;
}

function tableLookBandingEnabled(look: XmlNode | undefined, direction: "horizontal" | "vertical"): boolean {
  const property = direction === "horizontal" ? "noHBand" : "noVBand";
  const value = attribute(look, property);
  if (value !== undefined) return value === "0" || value.toLowerCase() === "false";
  const lookValue = Number.parseInt(attribute(look, "val") ?? "", 16);
  if (!Number.isFinite(lookValue)) return true;
  const bit = direction === "horizontal" ? 0x200 : 0x400;
  return (lookValue & bit) === 0;
}

function parsePageLayout(body: XmlNode): WordPageLayout | undefined {
  const section = descendants(body, "sectPr").at(-1); if (!section) return undefined;
  const size = child(section, "pgSz"), margins = child(section, "pgMar");
  const width = Number(attribute(size, "w")), height = Number(attribute(size, "h"));
  if (!(width > 0 && height > 0)) return undefined;
  return { width: twipsToPoints(width), height: twipsToPoints(height), marginTop: twipsToPoints(Number(attribute(margins, "top") ?? 1440)), marginRight: twipsToPoints(Number(attribute(margins, "right") ?? 1440)), marginBottom: twipsToPoints(Number(attribute(margins, "bottom") ?? 1440)), marginLeft: twipsToPoints(Number(attribute(margins, "left") ?? 1440)), ...(attribute(margins, "header") !== undefined ? { headerDistance: twipsToPoints(Number(attribute(margins, "header"))) } : {}), ...(attribute(margins, "footer") !== undefined ? { footerDistance: twipsToPoints(Number(attribute(margins, "footer"))) } : {}) };
}

function readHeaderFooter(zip: ZipArchive, ids: string[], relationships: Map<string, string>, context: WordParseContext): WordBlock[] {
  for (const id of ids) {
    const path = relationships.get(id);
    if (path && zip.has(path)) {
      const directory = path.slice(0, path.lastIndexOf("/")), filename = path.slice(path.lastIndexOf("/") + 1);
      const localRelationships = parseRelationships(zip, `${directory}/_rels/${filename}.rels`, directory);
      const localContext = { ...context, relationships: localRelationships, images: readWordImages(zip, localRelationships) };
      return parseWordBlocks(decodeXml(zip, path), localContext);
    }
  }
  return [];
}

function readWordImages(zip: ZipArchive, relationships: Map<string, string>): Map<string, { bytes: Uint8Array; mimeType: string }> {
  const images = new Map<string, { bytes: Uint8Array; mimeType: string }>();
  const imageTypes: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp" };
  for (const [id, path] of relationships) {
    if (!/^word\/media\//i.test(path) || !zip.has(path)) continue;
    const extension = path.split(".").pop()?.toLowerCase() ?? "", mimeType = imageTypes[extension];
    if (mimeType) images.set(id, { bytes: zip.read(path), mimeType });
  }
  return images;
}

function descendants(node: XmlNode, name: string): XmlNode[] {
  const found: XmlNode[] = [];
  const visit = (part: XmlNode) => { for (const item of part.children) if (typeof item !== "string") { if (item.localName === name) found.push(item); visit(item); } };
  visit(node); return found;
}
function descendantsWithoutNestedTables(node: XmlNode, name: string): XmlNode[] {
  const found: XmlNode[] = [];
  const visit = (part: XmlNode) => { for (const item of part.children) if (typeof item !== "string") { if (item.localName === "tbl" && item !== node) continue; if (item.localName === name) found.push(item); else visit(item); } };
  visit(node); return found;
}
function directDescendantsWithoutNestedTables(node: XmlNode, name: string): XmlNode[] {
  const found: XmlNode[] = [];
  const visit = (part: XmlNode) => { for (const item of part.children) if (typeof item !== "string") { if (item.localName === "tbl") continue; if (item.localName === name) found.push(item); else visit(item); } };
  visit(node); return found;
}
function attribute(node: XmlNode | undefined, localName: string): string | undefined {
  if (!node) return undefined;
  const entry = Object.entries(node.attributes).find(([name]) => name === localName || name.endsWith(`:${localName}`)); return entry?.[1];
}
function onOff(node: XmlNode | undefined): boolean | undefined {
  if (!node) return undefined;
  const value = attribute(node, "val"); return value === undefined || !["0", "false", "off", "no"].includes(value.toLowerCase());
}
function twipsToPoints(value: number): number { return Number.isFinite(value) ? value / 20 : 0; }
function wordColor(value: string): string {
  const colors: Record<string, string> = { yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff", red: "#ff0000", darkBlue: "#000080", darkCyan: "#008080", darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000", darkYellow: "#808000", darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000", white: "#ffffff" };
  return colors[value] ?? "#ffff00";
}
function formatListNumber(value: number, format: string): string {
  if (/^(?:lowerLetter|upperLetter)$/i.test(format)) {
    let remaining = Math.max(1, value), result = "";
    while (remaining > 0) { remaining--; result = String.fromCharCode((format === "upperLetter" ? 65 : 97) + remaining % 26) + result; remaining = Math.floor(remaining / 26); }
    return result;
  }
  if (/^(?:lowerRoman|upperRoman)$/i.test(format)) {
    let remaining = Math.max(1, Math.min(value, 3999)), result = "";
    for (const [amount, glyph] of [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]] as const) while (remaining >= amount) { result += glyph; remaining -= amount; }
    return format === "lowerRoman" ? result.toLowerCase() : result;
  }
  if (format.toLowerCase() === "ordinal") return `${value}${value % 10 === 1 && value % 100 !== 11 ? "st" : value % 10 === 2 && value % 100 !== 12 ? "nd" : value % 10 === 3 && value % 100 !== 13 ? "rd" : "th"}`;
  return String(value);
}

export function parsePptx(bytes: Uint8Array, name: string): Presentation {
  const zip = ZipArchive.open(bytes), presentation = decodeXml(zip, "ppt/presentation.xml");
  const relationships = parseRelationships(zip, "ppt/_rels/presentation.xml.rels", "ppt");
  const masterRef = descendants(child(presentation, "sldMasterIdLst") ?? presentation, "sldMasterId")[0];
  const masterRelId = Object.entries(attrs(masterRef)).find(([key]) => key.endsWith(":id"))?.[1];
  const defaultMasterPath = relationships.get(masterRelId ?? "");
  const masters = new Map<string, PptMaster>();
  const loadMaster = (masterPath: string | undefined): PptMaster => {
    const key = masterPath ?? "";
    const cached = masters.get(key); if (cached) return cached;
    const masterRels = masterPath ? parseRelationships(zip, relationshipPart(masterPath), masterPath.slice(0, masterPath.lastIndexOf("/"))) : new Map<string, string>();
    const themePath = [...masterRels.values()].find(path => /^ppt\/theme\//i.test(path)) ?? "ppt/theme/theme1.xml";
    const theme = zip.has(themePath) ? readPptTheme(decodeXml(zip, themePath)) : emptyPptTheme;
    const root = masterPath && zip.has(masterPath) ? decodeXml(zip, masterPath) : undefined;
    const entries = pptShapeTreeEntries(child(child(root, "cSld"), "spTree")).map(entry => ({ ...entry, relationships: masterRels }));
    const master = { root, relationships: masterRels, theme, entries };
    masters.set(key, master);
    return master;
  };
  const slideSize = child(presentation, "sldSz");
  const slideWidth = Math.max(1, Number(attribute(slideSize, "cx") ?? "9144000"));
  const slideHeight = Math.max(1, Number(attribute(slideSize, "cy") ?? "5143500"));
  const slideIds = elements(child(presentation, "sldIdLst") ?? presentation, "sldId");
  const slides: Presentation["slides"] = [];
  for (const slideId of slideIds) {
    const relId = Object.entries(attrs(slideId)).find(([key]) => key.endsWith(":id"))?.[1], path = relationships.get(relId ?? "");
    if (!path || !zip.has(path)) continue;
    const slide = decodeXml(zip, path), paragraphs: Paragraph[] = [], elements: PptxElement[] = [];
    const slideRels = parseRelationships(zip, relationshipPart(path), path.slice(0, path.lastIndexOf("/")));
    const layoutPath = [...slideRels.values()].find(target => /^ppt\/slideLayouts\//i.test(target));
    const layout = layoutPath && zip.has(layoutPath) ? decodeXml(zip, layoutPath) : undefined;
    const layoutTree = child(child(layout, "cSld"), "spTree");
    const layoutRels = layoutPath ? parseRelationships(zip, relationshipPart(layoutPath), layoutPath.slice(0, layoutPath.lastIndexOf("/"))) : new Map<string, string>();
    const layoutEntries = pptShapeTreeEntries(layoutTree).map(entry => ({ ...entry, relationships: layoutRels }));
    const layoutShapeEntries = layoutEntries.filter(entry => entry.item.localName === "sp");
    const { root: master, relationships: masterRels, theme: baseTheme, entries: masterEntries } = loadMaster([...layoutRels.values()].find(target => /^ppt\/slideMasters\//i.test(target)) ?? defaultMasterPath);
    const masterShapeEntries = masterEntries.filter(entry => entry.item.localName === "sp");
    const masterColorMap = child(master, "clrMap");
    const themeOverridePath = [...slideRels.values()].find(target => /^ppt\/theme\/themeOverride\d+\.xml$/i.test(target));
    const theme = themeOverridePath && zip.has(themeOverridePath)
      ? { ...baseTheme, colors: { ...baseTheme.colors, ...readPptTheme(decodeXml(zip, themeOverridePath)).colors } }
      : baseTheme;
    const overrideColorMap = (part: XmlNode | undefined) => attrs(child(child(part, "clrMapOvr"), "overrideClrMapping"));
    const colorMap = { ...defaultPptColorMap, ...Object.fromEntries(Object.entries(attrs(masterColorMap)).filter(([key]) => !key.startsWith("xmlns"))), ...overrideColorMap(layout), ...overrideColorMap(slide), ...attrs(child(slide, "clrMap")) };
    const palette = (node: XmlNode | undefined) => readPptColor(node, theme, colorMap);
    const baseColor = resolvePptSchemeColor("tx1", theme, colorMap) ?? "#000000";
    const textStyles = child(master, "txStyles");
    const masterTextStyle = (style: string) => child(textStyles, style), masterParagraphStyle = (style: string) => child(masterTextStyle(style), "lvl1pPr");
    const masterDefault = (style: string, font: string) => {
      const def = child(masterParagraphStyle(style), "defRPr");
      return { fontFamily: font, ...readPptRunStyle(def, font, palette, theme) };
    };
    const titleDefaults = masterDefault("titleStyle", theme.majorFont);
    const bodyDefaults = masterDefault("bodyStyle", theme.minorFont);
    const otherDefaults = masterDefault("otherStyle", theme.minorFont);
    const slideTree = child(child(slide, "cSld"), "spTree") ?? child(slide, "spTree");
    const shapeDefaults: Partial<TextRun> = { fontFamily: theme.minorFont, color: baseColor };

    // The background comes from the first of slide, layout and master that defines one.
    const backgroundSource = [{ part: slide, rels: slideRels }, { part: layout, rels: layoutRels }, { part: master, rels: masterRels }].find(source => child(child(source.part, "cSld"), "bg"));
    const bg = child(child(backgroundSource?.part, "cSld"), "bg"), bgPr = child(bg, "bgPr"), bgRef = child(bg, "bgRef");
    const backgroundRefIndex = Number(attribute(bgRef, "idx") ?? 0);
    const backgroundStyle = bgRef ? (backgroundRefIndex >= 1001 ? pptStyleEntry(theme.backgroundFillStyles, backgroundRefIndex - 1000) : pptStyleEntry(theme.fillStyles, backgroundRefIndex)) : undefined;
    const background = bgPr ? readPptFill(bgPr, theme, colorMap) : bgRef ? readPptStyleFill(bgRef, theme, colorMap) : undefined;
    const backgroundFill = child(bgPr, "blipFill") ?? (backgroundStyle?.localName === "blipFill" ? backgroundStyle : undefined), backgroundBlip = child(backgroundFill, "blip");
    const backgroundImagePath = backgroundSource?.rels.get(attribute(backgroundBlip, "embed") ?? "");
    if (backgroundImagePath && zip.has(backgroundImagePath)) {
      const mimeType = pptImageMime(backgroundImagePath), backgroundDuotone = child(backgroundBlip, "duotone");
      const duotoneTheme = bgRef ? pptStyleTheme(bgRef, theme, colorMap) : theme;
      const duotoneColors = backgroundDuotone?.children.filter((part): part is XmlNode => typeof part !== "string") ?? [];
      const darkColor = readPptColor(duotoneColors[0], duotoneTheme, colorMap)?.color, lightColor = readPptColor(duotoneColors[1], duotoneTheme, colorMap)?.color;
      const cropNode = child(backgroundFill, "srcRect"), cropAttrs = attrs(cropNode);
      if (mimeType) elements.push({
        kind: "image", left: 0, top: 0, width: slideWidth, height: slideHeight, isBackground: true,
        image: { bytes: zip.read(backgroundImagePath), mimeType },
        ...(cropNode ? { crop: { left: Number(cropAttrs["l"] ?? 0) / 100_000, top: Number(cropAttrs["t"] ?? 0) / 100_000, right: Number(cropAttrs["r"] ?? 0) / 100_000, bottom: Number(cropAttrs["b"] ?? 0) / 100_000 } } : {}),
        ...(darkColor && lightColor ? { duotone: { darkColor, lightColor } } : {}),
      });
    }

    const isPlaceholder = (entry: PptTreeEntry) => Boolean(child(child(child(entry.item, "nvSpPr"), "nvPr"), "ph"));
    const showMasterShapes = (part: XmlNode | undefined) => !["0", "false"].includes(attribute(part, "showMasterSp") ?? "1");
    const masterItems = showMasterShapes(slide) && showMasterShapes(layout) ? masterEntries.filter(entry => entry.item.localName !== "sp" || !isPlaceholder(entry)) : [];
    const layoutItems = showMasterShapes(slide) ? layoutEntries.filter(entry => entry.item.localName !== "sp" || !isPlaceholder(entry)) : [];
    const slideItems = pptShapeTreeEntries(slideTree).map(entry => ({ ...entry, relationships: slideRels })).flatMap(entry => {
      if (entry.item.localName !== "graphicFrame" || descendants(entry.item, "tbl").length) return [entry];
      const frameTransform = readPptTransform(child(entry.item, "xfrm"));
      if (!frameTransform) return [entry];
      const frameBox = transformPptTransform(frameTransform, entry.transform);
      if (!frameBox) return [entry];
      const diagramRelIds = descendants(entry.item, "relIds")[0];
      const dataRelId = Object.entries(attrs(diagramRelIds)).find(([key]) => key.endsWith(":dm"))?.[1];
      const dataPath = slideRels.get(dataRelId ?? ""), drawingPath = dataPath?.replace(/\/data(\d+)\.xml$/i, "/drawing$1.xml");
      if (!drawingPath || drawingPath === dataPath || !zip.has(drawingPath)) return [entry];
      const diagram = decodeXml(zip, drawingPath), drawingTree = child(diagram, "spTree");
      const scaleX = frameBox.width / Math.max(1, frameTransform.width), scaleY = frameBox.height / Math.max(1, frameTransform.height);
      const diagramTransform: PptGroupTransform = { scaleX, scaleY, translateX: frameBox.left, translateY: frameBox.top, rotation: frameBox.rotation ?? 0 };
      const diagramRels = parseRelationships(zip, relationshipPart(drawingPath), drawingPath.slice(0, drawingPath.lastIndexOf("/")));
      return pptShapeTreeEntries(drawingTree, diagramTransform).map(diagramEntry => ({ ...diagramEntry, relationships: diagramRels, diagram: true }));
    });
    const renderItems = [...masterItems, ...layoutItems, ...slideItems];
    for (const entry of renderItems) {
      const item = entry.item, fromSlide = !(masterItems as PptTreeEntry[]).includes(entry) && !(layoutItems as PptTreeEntry[]).includes(entry);
      const itemRels = entry.relationships ?? slideRels;
      if (item.localName === "sp" || item.localName === "cxnSp") {
        const placeholder = fromSlide ? child(child(child(item, "nvSpPr"), "nvPr"), "ph") : undefined;
        const placeholderType = attribute(placeholder, "type"), placeholderKeyType = placeholderType ?? "obj", placeholderIndex = attribute(placeholder, "idx") ?? "0";
        const placeholderOf = (candidate: PptTreeEntry) => child(child(child(candidate.item, "nvSpPr"), "nvPr"), "ph");
        const layoutShapeEntry = placeholder ? layoutShapeEntries.find(candidate => {
          const ref = placeholderOf(candidate);
          return ref && (attribute(ref, "type") ?? "obj") === placeholderKeyType && (attribute(ref, "idx") ?? "0") === placeholderIndex;
        }) ?? layoutShapeEntries.find(candidate => {
          const ref = placeholderOf(candidate);
          return ref && (placeholderIndex !== "0" ? (attribute(ref, "idx") ?? "0") === placeholderIndex : pptPlaceholderFamily(attribute(ref, "type")) === pptPlaceholderFamily(placeholderType));
        }) : undefined;
        const masterFamily = pptPlaceholderFamily(attribute(placeholderOf(layoutShapeEntry ?? entry), "type") ?? placeholderType);
        const masterShapeEntry = placeholder ? masterShapeEntries.find(candidate => pptPlaceholderFamily(attribute(placeholderOf(candidate), "type")) === masterFamily && Boolean(placeholderOf(candidate))) : undefined;
        const layoutShape = layoutShapeEntry?.item, masterShape = masterShapeEntry?.item;
        const shapeProps = child(item, "spPr");
        const layoutProps = child(layoutShape, "spPr"), masterProps = child(masterShape, "spPr");
        const ownBox = transformPptTransform(readPptTransform(child(shapeProps, "xfrm")), entry.transform);
        const layoutBox = layoutShapeEntry ? transformPptTransform(readPptTransform(child(layoutProps, "xfrm")), layoutShapeEntry.transform) : undefined;
        const masterBox = masterShapeEntry ? transformPptTransform(readPptTransform(child(masterProps, "xfrm")), masterShapeEntry.transform) : undefined;
        const box = ownBox ?? layoutBox ?? masterBox;
        const geometry = child(shapeProps, "prstGeom") ?? child(layoutProps, "prstGeom");
        const geometryName = attribute(geometry, "prst");
        const customGeometry = child(shapeProps, "custGeom") ?? child(layoutProps, "custGeom");
        const isLine = geometryName === "line" || geometryName === "straightConnector1" || (item.localName === "cxnSp" && !geometryName && !customGeometry);
        const pathList = child(customGeometry, "pathLst");
        const customPaths = pathList ? pathList.children.filter((part): part is XmlNode => typeof part !== "string" && part.localName === "path").flatMap(path => {
          const pathWidth = Number(attribute(path, "w") ?? box?.width ?? 0), pathHeight = Number(attribute(path, "h") ?? box?.height ?? 0);
          const d = drawingPathData(path);
          return d && (pathWidth > 0 || pathHeight > 0) ? [{ d, width: pathWidth, height: pathHeight, fill: attribute(path, "fill") !== "none", stroke: attribute(path, "stroke") !== "false" && attribute(path, "stroke") !== "0" }] : [];
        }) : [];
        const style = child(item, "style");
        const styleReference = (name: string) => child(style, name);
        const styleIndex = (name: string) => Number(attribute(styleReference(name), "idx") ?? 0);
        const fillElementNames = ["solidFill", "gradFill", "noFill", "blipFill", "pattFill", "grpFill"];
        const definesFill = (props: XmlNode | undefined) => Boolean(props) && fillElementNames.some(fillName => child(props, fillName));
        const useBackgroundFill = attribute(item, "useBgFill") === "1" || attribute(item, "useBgFill") === "true";
        const fill = useBackgroundFill ? background
          : definesFill(shapeProps) ? readPptFill(shapeProps, theme, colorMap)
            : definesFill(layoutProps) ? readPptFill(layoutProps, theme, colorMap)
              : definesFill(masterProps) ? readPptFill(masterProps, theme, colorMap)
                : readPptStyleFill(styleReference("fillRef"), theme, colorMap);
        const line = resolvePptLine([
          { node: pptStyleEntry(theme.lineStyles, styleIndex("lnRef")), theme: pptStyleTheme(styleReference("lnRef"), theme, colorMap) },
          { node: child(masterProps, "ln"), theme }, { node: child(layoutProps, "ln"), theme }, { node: child(shapeProps, "ln"), theme },
        ], colorMap);
        const shadow = child(shapeProps, "effectLst") ? readPptShadow(shapeProps, theme, colorMap)
          : child(layoutProps, "effectLst") ? readPptShadow(layoutProps, theme, colorMap)
            : readPptShadow(pptStyleEntry(theme.effectStyles, styleIndex("effectRef")), pptStyleTheme(styleReference("effectRef"), theme, colorMap), colorMap);
        const flips = { ...(box?.flipH ? { flipH: true } : {}), ...(box?.flipV ? { flipV: true } : {}) };
        const shapeBlip = child(child(shapeProps, "blipFill"), "blip"), shapeImagePath = itemRels.get(attribute(shapeBlip, "embed") ?? ""), shapeImageMime = shapeImagePath ? pptImageMime(shapeImagePath) : undefined;
        const adjustments = readPptAdjustments(geometry);
        if (box && shapeImagePath && shapeImageMime && zip.has(shapeImagePath)) {
          elements.push({ kind: "image", ...box, image: { bytes: zip.read(shapeImagePath), mimeType: shapeImageMime }, ...(geometryName && geometryName !== "rect" ? { geometry: geometryName } : {}), ...(adjustments ? { adjustments } : {}) });
        }
        if (box && (geometry || customPaths.length || isLine)) {
          if (fill || line || isLine || customPaths.length || shadow || useBackgroundFill) {
            if (isLine && !customPaths.length) elements.push({ kind: "line", ...box, ...(line ? { line } : {}), ...(shadow ? { shadow } : {}) });
            else {
              const shape: PptxShapeElement = {
                kind: "shape", ...box, ...flips, ...(geometryName ? { geometry: geometryName } : {}), ...(adjustments ? { adjustments } : {}), ...(customPaths.length ? { customPaths } : {}),
                ...(fill ? { fill } : {}), ...(line ? { line } : {}), ...(shadow ? { shadow } : {}), ...(useBackgroundFill ? { useBackgroundFill } : {}),
              };
              elements.push(shape);
            }
          }
        }
        const body = child(item, "txBody");
        if (body) {
          const layoutBody = child(layoutShape, "txBody"), masterBody = child(masterShape, "txBody");
          const isTitlePlaceholder = /^(?:title|ctrTitle)$/i.test(placeholderType ?? "");
          const inherited = isTitlePlaceholder ? titleDefaults : placeholder ? bodyDefaults : otherDefaults;
          const defaultFamily = inherited.fontFamily ?? (isTitlePlaceholder ? theme.majorFont : theme.minorFont);
          const masterParagraph = masterTextStyle(isTitlePlaceholder ? "titleStyle" : placeholder ? "bodyStyle" : "otherStyle");
          // The shape style's font reference beats the master's default text style for non-placeholder text.
          const fontReference = styleReference("fontRef"), fontReferenceColor = palette(fontReference?.children.find((part): part is XmlNode => typeof part !== "string" && pptColorElementNames.includes(part.localName)));
          const fontReferenceFamily = attribute(fontReference, "idx") === "major" ? theme.majorFont : attribute(fontReference, "idx") === "minor" ? theme.minorFont : undefined;
          const fontReferenceRun: Partial<TextRun> = { ...(fontReferenceFamily && !placeholder ? { fontFamily: fontReferenceFamily } : {}), ...(fontReferenceColor ? { color: fontReferenceColor.opacity < 1 ? `rgba(${parsePptRgb(fontReferenceColor.color).join(",")},${fontReferenceColor.opacity})` : fontReferenceColor.color } : {}) };
          // SmartArt drawings carry a separate text rectangle; PowerPoint keeps upside-down diagram labels upright.
          const textTransform = transformPptTransform(readPptTransform(child(item, "txXfrm")), entry.transform);
          const textBox = textTransform ? pptUprightText(textTransform) : box;
          const text = parsePptTextBody(body, textBox ?? { left: 0, top: 0, width: 0, height: 0 }, { ...shapeDefaults, ...inherited, fontFamily: defaultFamily, ...fontReferenceRun }, palette, theme, colorMap, layoutBody, masterParagraph, masterBody);
          if (text.paragraphs.some(paragraph => paragraph.runs.some(run => run.text.length > 0))) {
            if (fromSlide) paragraphs.push(...text.paragraphs);
            if (textBox) elements.push(text);
          }
        }
      } else if (item.localName === "pic") {
        const box = transformPptTransform(readPptTransform(child(child(item, "spPr"), "xfrm")), entry.transform);
        const embed = attribute(child(child(item, "blipFill"), "blip"), "embed"), imagePath = itemRels.get(embed ?? "");
        if (!box || !imagePath || !zip.has(imagePath)) continue;
        const mimeType = pptImageMime(imagePath);
        if (!mimeType) continue;
        const cropNode = child(child(item, "blipFill"), "srcRect"), cropAttrs = attrs(cropNode);
        const crop = cropNode ? { left: Number(cropAttrs["l"] ?? 0) / 100_000, top: Number(cropAttrs["t"] ?? 0) / 100_000, right: Number(cropAttrs["r"] ?? 0) / 100_000, bottom: Number(cropAttrs["b"] ?? 0) / 100_000 } : undefined;
        const nonVisual = descendants(item, "cNvPr")[0];
        const alt = attribute(nonVisual, "descr") ?? attribute(nonVisual, "name");
        const pictureGeometry = child(child(item, "spPr"), "prstGeom"), pictureGeometryName = attribute(pictureGeometry, "prst"), pictureAdjustments = readPptAdjustments(pictureGeometry);
        const image: PptxImageElement = {
          kind: "image", ...box, image: { bytes: zip.read(imagePath), mimeType, ...(alt ? { alt } : {}) }, ...(crop ? { crop } : {}),
          ...(pictureGeometryName && pictureGeometryName !== "rect" ? { geometry: pictureGeometryName } : {}), ...(pictureAdjustments ? { adjustments: pictureAdjustments } : {}),
        };
        elements.push(image);
        const pictureLine = resolvePptLine([{ node: child(child(item, "spPr"), "ln"), theme }], colorMap);
        if (pictureLine) elements.push({ kind: "shape", ...box, geometry: pictureGeometryName ?? "rect", ...(pictureAdjustments ? { adjustments: pictureAdjustments } : {}), line: pictureLine });
      } else if (item.localName === "graphicFrame") {
        const table = parsePptTable(item, theme, colorMap, palette, entry.transform);
        if (table) { elements.push(table); }
      }
    }
    const title = paragraphs[0]?.runs.map(run => run.text).join("").trim() || `Slide ${slides.length + 1}`;
    slides.push({ title, paragraphs, width: slideWidth, height: slideHeight, ...(background ? { background } : {}), elements });
  }
  if (!slides.length) throw new PreviewError("PPTX presentation contains no readable slides.", "invalid-file");
  return { kind: "presentation", slides, title: name };
}

function pptUprightText(box: PptTransform): PptTransform {
  const { rotation, ...rest } = box, normalized = (((rotation ?? 0) % 360) + 360) % 360;
  return Math.abs(normalized - 180) < 0.5 || normalized < 0.5 ? rest : box;
}
type PptMaster = { root: XmlNode | undefined; relationships: Map<string, string>; theme: PptTheme; entries: PptTreeEntry[] };
const emptyPptTheme: PptTheme = { colors: {}, majorFont: "Arial", minorFont: "Arial", fillStyles: [], lineStyles: [], effectStyles: [], backgroundFillStyles: [] };
/** Groups placeholder types the way slide masters define them (title, body, date, footer, slide number). */
function pptPlaceholderFamily(type: string | undefined): string {
  if (type === "title" || type === "ctrTitle") return "title";
  if (type === "dt" || type === "ftr" || type === "sldNum" || type === "hdr") return type;
  return "body";
}
function readPptAdjustments(geometry: XmlNode | undefined): Record<string, number> | undefined {
  const guides = elements(child(geometry, "avLst") ?? geometry ?? { name: "", localName: "", attributes: {}, children: [] }, "gd");
  const result: Record<string, number> = {};
  for (const guide of guides) {
    const match = /^val\s+(-?\d+(?:\.\d+)?)$/.exec(attribute(guide, "fmla") ?? ""), guideName = attribute(guide, "name");
    if (match && guideName) result[guideName] = Number(match[1]);
  }
  return Object.keys(result).length ? result : undefined;
}

export type DrawingShapeContext = {
  /** Theme part root (a:theme); colors and fonts default to an empty theme when absent. */
  theme?: XmlNode | undefined;
  /** Resolves a picture relationship id to image bytes. */
  image(embed: string): { bytes: Uint8Array; mimeType: string } | undefined;
  /** Box for a top-level shape stored without its own transform (EMU). */
  fallbackBox?: { left: number; top: number; width: number; height: number };
};

/** Converts a DrawingML shape tree outside a slide (spreadsheet drawing anchors and their groups) into positioned elements in EMU. */
export function parseDrawingShapes(tree: XmlNode, context: DrawingShapeContext): PptxElement[] {
  const theme = context.theme ? readPptTheme(context.theme) : emptyPptTheme, colorMap = defaultPptColorMap;
  const palette = (node: XmlNode | undefined) => readPptColor(node, theme, colorMap);
  const result: PptxElement[] = [];
  for (const entry of pptShapeTreeEntries(tree)) {
    const item = entry.item, props = child(item, "spPr");
    const ownBox = transformPptTransform(readPptTransform(child(props, "xfrm")), entry.transform);
    const box: PptTransform | undefined = ownBox ?? (item.localName !== "pic" && tree.children.includes(item) ? context.fallbackBox : undefined);
    if (!box) continue;
    if (item.localName === "pic") {
      const nonVisual = descendants(item, "cNvPr")[0], description = attribute(nonVisual, "descr") ?? attribute(nonVisual, "name");
      const image = context.image(attribute(child(child(item, "blipFill"), "blip"), "embed") ?? "");
      const cropNode = child(child(item, "blipFill"), "srcRect"), cropAttrs = attrs(cropNode);
      if (image) result.push({
        kind: "image", ...box, image: { ...image, ...(description ? { alt: description } : {}) },
        ...(cropNode ? { crop: { left: Number(cropAttrs["l"] ?? 0) / 100_000, top: Number(cropAttrs["t"] ?? 0) / 100_000, right: Number(cropAttrs["r"] ?? 0) / 100_000, bottom: Number(cropAttrs["b"] ?? 0) / 100_000 } } : {}),
      });
      continue;
    }
    if (item.localName !== "sp" && item.localName !== "cxnSp") continue;
    const style = child(item, "style"), reference = (name: string) => child(style, name), styleIndex = (name: string) => Number(attribute(reference(name), "idx") ?? 0);
    const geometry = child(props, "prstGeom"), geometryName = attribute(geometry, "prst"), customGeometry = child(props, "custGeom");
    const isLine = geometryName === "line" || geometryName === "straightConnector1" || (item.localName === "cxnSp" && !geometryName && !customGeometry);
    const customPaths = elements(child(customGeometry, "pathLst") ?? { name: "", localName: "", attributes: {}, children: [] }, "path").flatMap(path => {
      const pathWidth = Number(attribute(path, "w") ?? box.width), pathHeight = Number(attribute(path, "h") ?? box.height), d = drawingPathData(path);
      return d && (pathWidth > 0 || pathHeight > 0) ? [{ d, width: pathWidth, height: pathHeight, fill: attribute(path, "fill") !== "none", stroke: attribute(path, "stroke") !== "false" && attribute(path, "stroke") !== "0" }] : [];
    });
    const definesFill = ["solidFill", "gradFill", "noFill", "blipFill", "pattFill", "grpFill"].some(name => child(props, name));
    const fill = definesFill ? readPptFill(props, theme, colorMap) : readPptStyleFill(reference("fillRef"), theme, colorMap);
    const line = resolvePptLine([{ node: pptStyleEntry(theme.lineStyles, styleIndex("lnRef")), theme: pptStyleTheme(reference("lnRef"), theme, colorMap) }, { node: child(props, "ln"), theme }], colorMap);
    const shadow = child(props, "effectLst") ? readPptShadow(props, theme, colorMap) : readPptShadow(pptStyleEntry(theme.effectStyles, styleIndex("effectRef")), pptStyleTheme(reference("effectRef"), theme, colorMap), colorMap);
    const camera = readPptCamera(props), adjustments = readPptAdjustments(geometry);
    const flips = { ...(box.flipH ? { flipH: true } : {}), ...(box.flipV ? { flipV: true } : {}) };
    if (isLine && !customPaths.length) result.push({ kind: "line", ...box, ...(line ? { line } : {}), ...(shadow ? { shadow } : {}) });
    else if (fill || line || shadow) result.push({
      kind: "shape", ...box, ...flips, ...(geometryName ? { geometry: geometryName } : {}), ...(adjustments ? { adjustments } : {}), ...(customPaths.length ? { customPaths } : {}),
      ...(fill ? { fill } : {}), ...(line ? { line } : {}), ...(shadow ? { shadow } : {}), ...(camera ? { camera } : {}),
    });
    const body = child(item, "txBody");
    if (!body) continue;
    // Shape text defaults to the style's font reference (theme font and color), at Excel's 11 pt.
    const fontReference = reference("fontRef"), fontColor = palette(fontReference?.children.find((part): part is XmlNode => typeof part !== "string" && pptColorElementNames.includes(part.localName)));
    const defaults: Partial<TextRun> = { fontFamily: attribute(fontReference, "idx") === "major" ? theme.majorFont : theme.minorFont, fontSize: 11, color: fontColor?.color ?? "#000000" };
    const text = parsePptTextBody(body, presetTextBox(geometryName, box, adjustments), defaults, palette, theme, colorMap);
    if (!text.paragraphs.some(paragraph => paragraph.runs.some(run => run.text.length > 0))) continue;
    const reflection = readPptReflection(props);
    result.push({ ...text, ...(reflection ? { reflection } : {}), ...(camera ? { camera } : {}) });
  }
  return result;
}

/** Text rectangle of a preset geometry (ECMA-376 presetShapeDefinitions): round shapes keep their text inside the inscribed square, arrows inside their shaft. */
function presetTextBox(geometry: string | undefined, box: PptTransform, adjustments?: Record<string, number>): PptTransform {
  if (geometry === "ellipse" || geometry === "teardrop" || geometry === "flowChartConnector") {
    const inset = (1 - Math.SQRT1_2) / 2;
    return { ...box, left: box.left + box.width * inset, top: box.top + box.height * inset, width: box.width * Math.SQRT1_2, height: box.height * Math.SQRT1_2 };
  }
  const { width: w, height: h } = box, ss = Math.min(w, h), pin = (value: number, max: number) => Math.max(0, Math.min(max, value));
  const rect = (left: number, top: number, right: number, bottom: number) => ({ ...box, left: box.left + left, top: box.top + top, width: Math.max(0, right - left), height: Math.max(0, bottom - top) });
  if (ss <= 0) return box;
  if (geometry === "leftArrow" || geometry === "rightArrow") {
    const a1 = pin(adjustments?.["adj1"] ?? 50_000, 100_000), a2 = pin(adjustments?.["adj2"] ?? 50_000, 100_000 * w / ss);
    const dx1 = ss * a2 / 100_000, dy1 = h * a1 / 200_000, y1 = h / 2 - dy1, dx2 = y1 * dx1 / (h / 2);
    return geometry === "rightArrow" ? rect(0, y1, w - dx1 + dx2, h - y1) : rect(dx1 - dx2, y1, w, h - y1);
  }
  if (geometry === "upArrow" || geometry === "downArrow") {
    const a1 = pin(adjustments?.["adj1"] ?? 50_000, 100_000), a2 = pin(adjustments?.["adj2"] ?? 50_000, 100_000 * h / ss);
    const dy1 = ss * a2 / 100_000, dx1 = w * a1 / 200_000, x1 = w / 2 - dx1, dy2 = x1 * dy1 / (w / 2);
    return geometry === "downArrow" ? rect(x1, 0, w - x1, h - dy1 + dy2) : rect(x1, dy1 - dy2, w - x1, h);
  }
  if (geometry === "homePlate") {
    const a = pin(adjustments?.["adj"] ?? 50_000, 100_000 * w / ss);
    return rect(0, 0, (w - ss * a / 100_000 + w) / 2, h);
  }
  return box;
}

/** Perspective cameras view the scene from a fixed distance; calibrated against Excel's perspectiveRelaxed camera (about 1490 px at 100%). */
const PERSPECTIVE_DISTANCE = 14_190_000;
/** Built-in tilt (latitude, degrees) of camera presets used without an explicit rotation. */
const cameraPresetLatitude: Record<string, number> = { perspectiveRelaxed: 309.6, perspectiveRelaxedModerately: 324.8 };
function readPptCamera(props: XmlNode | undefined): PptxShapeElement["camera"] {
  const scene = child(props, "scene3d"), camera = child(scene, "camera");
  if (!camera) return undefined;
  const preset = attribute(camera, "prst") ?? "orthographicFront", rotationNode = child(camera, "rot");
  const rotation = attrs(rotationNode), angle = (key: string) => Number(rotation[key] ?? 0) / 60_000;
  const latitude = rotationNode ? angle("lat") : cameraPresetLatitude[preset] ?? 0, longitude = angle("lon"), revolution = angle("rev");
  // A scene seen straight on is drawn like a plain shape; turned scenes are projected and lit.
  if (![latitude, longitude, revolution].some(value => value % 360 !== 0)) return undefined;
  const fov = Number(attribute(camera, "fov") ?? 2_700_000) / 60_000;
  const distance = /perspective/i.test(preset) && fov > 0 && fov < 180 ? PERSPECTIVE_DISTANCE * Math.tan(Math.PI / 8) / Math.tan(fov * Math.PI / 360) : undefined;
  return { latitude, longitude, revolution, ...(distance ? { distance } : {}), ...(child(scene, "lightRig") ? { lit: true } : {}) };
}

function readPptReflection(props: XmlNode | undefined): PptxTextElement["reflection"] {
  const reflection = child(child(props, "effectLst"), "reflection");
  if (!reflection) return undefined;
  const value = (key: string, fallback: number) => Number(attribute(reflection, key) ?? fallback);
  return { startOpacity: value("stA", 100_000) / 100_000, endOpacity: value("endA", 0) / 100_000, endPosition: value("endPos", 100_000) / 100_000, distance: value("dist", 0), blur: value("blurRad", 0) };
}

const defaultPptColorMap: Record<string, string> = { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2", accent1: "accent1", accent2: "accent2", accent3: "accent3", accent4: "accent4", accent5: "accent5", accent6: "accent6", hlink: "hlink", folHlink: "folHlink" };
type PptTheme = { colors: Record<string, string>; majorFont: string; minorFont: string; fillStyles: XmlNode[]; lineStyles: XmlNode[]; effectStyles: XmlNode[]; backgroundFillStyles: XmlNode[] };
type PptTransform = { left: number; top: number; width: number; height: number; rotation?: number; flipH?: boolean; flipV?: boolean };
type PptGroupTransform = { scaleX: number; scaleY: number; translateX: number; translateY: number; rotation: number };
type PptTreeEntry = { item: XmlNode; transform: PptGroupTransform; relationships?: Map<string, string>; diagram?: boolean };

function pptShapeTreeEntries(tree: XmlNode | undefined, parentTransform: PptGroupTransform = { scaleX: 1, scaleY: 1, translateX: 0, translateY: 0, rotation: 0 }): PptTreeEntry[] {
  const entries: PptTreeEntry[] = [];
  for (const item of tree?.children ?? []) {
    if (typeof item === "string") continue;
    if (item.localName === "grpSp") {
      const xfrm = child(child(item, "grpSpPr"), "xfrm"), off = child(xfrm, "off"), ext = child(xfrm, "ext"), chOff = child(xfrm, "chOff"), chExt = child(xfrm, "chExt");
      const childWidth = Number(attribute(chExt, "cx") ?? attribute(ext, "cx") ?? 1), childHeight = Number(attribute(chExt, "cy") ?? attribute(ext, "cy") ?? 1);
      const localScaleX = Number(attribute(ext, "cx") ?? childWidth) / Math.max(1, childWidth), localScaleY = Number(attribute(ext, "cy") ?? childHeight) / Math.max(1, childHeight);
      const localX = Number(attribute(off, "x") ?? 0) - Number(attribute(chOff, "x") ?? 0) * localScaleX;
      const localY = Number(attribute(off, "y") ?? 0) - Number(attribute(chOff, "y") ?? 0) * localScaleY;
      const transform: PptGroupTransform = {
        scaleX: parentTransform.scaleX * localScaleX,
        scaleY: parentTransform.scaleY * localScaleY,
        translateX: parentTransform.translateX + localX * parentTransform.scaleX,
        translateY: parentTransform.translateY + localY * parentTransform.scaleY,
        rotation: parentTransform.rotation + Number(attribute(xfrm, "rot") ?? 0) / 60_000,
      };
      entries.push(...pptShapeTreeEntries(item, transform));
    } else if (item.localName === "AlternateContent") {
      // Prefer the plain-DrawingML fallback over extension-specific choices.
      const branch = child(item, "Fallback") ?? child(item, "Choice");
      if (branch) entries.push(...pptShapeTreeEntries(branch, parentTransform));
    } else if (["sp", "cxnSp", "pic", "graphicFrame"].includes(item.localName)) entries.push({ item, transform: parentTransform });
  }
  return entries;
}

function transformPptTransform(box: PptTransform | undefined, group: PptGroupTransform): PptTransform | undefined {
  if (!box) return undefined;
  return {
    left: group.translateX + box.left * group.scaleX,
    top: group.translateY + box.top * group.scaleY,
    width: box.width * group.scaleX,
    height: box.height * group.scaleY,
    ...((box.rotation || group.rotation) ? { rotation: (box.rotation ?? 0) + group.rotation } : {}),
    ...(box.flipH ? { flipH: true } : {}), ...(box.flipV ? { flipV: true } : {}),
  };
}

function relationshipPart(path: string): string { const slash = path.lastIndexOf("/"); return `${path.slice(0, slash)}/_rels/${path.slice(slash + 1)}.rels`; }
function pptImageMime(path: string): string | undefined { const ext = path.split(".").pop()?.toLowerCase(); return ({ png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml" } as Record<string, string>)[ext ?? ""]; }
function readPptTheme(root: XmlNode): PptTheme {
  const elementsNode = child(root, "themeElements") ?? root, scheme = child(elementsNode, "clrScheme"), colors: Record<string, string> = {};
  for (const entry of scheme?.children ?? []) if (typeof entry !== "string") {
    const color = entry.children.find((part): part is XmlNode => typeof part !== "string");
    if (color) colors[entry.localName] = attribute(color, "lastClr") ?? attribute(color, "val") ?? "000000";
  }
  const fonts = child(elementsNode, "fontScheme");
  const latin = (parent: XmlNode | undefined, name: string) => attribute(child(parent, name), "typeface");
  const format = child(elementsNode, "fmtScheme"), list = (name: string) => child(format, name)?.children.filter((part): part is XmlNode => typeof part !== "string") ?? [];
  return {
    colors, majorFont: latin(child(fonts, "majorFont"), "latin") || "Arial", minorFont: latin(child(fonts, "minorFont"), "latin") || "Arial",
    fillStyles: list("fillStyleLst"), lineStyles: list("lnStyleLst"), effectStyles: list("effectStyleLst"), backgroundFillStyles: list("bgFillStyleLst"),
  };
}
function pptAttr(node: XmlNode | undefined, name: string): string | undefined {
  if (!node) return undefined;
  return Object.entries(node.attributes).find(([key]) => key === name || key.endsWith(`:${name}`))?.[1];
}
function readPptTransform(node: XmlNode | undefined): PptTransform | undefined {
  if (!node) return undefined;
  const off = child(node, "off"), ext = child(node, "ext");
  if (!off || !ext) return undefined;
  const left = Number(attribute(off, "x") ?? 0), top = Number(attribute(off, "y") ?? 0), width = Number(attribute(ext, "cx") ?? 0), height = Number(attribute(ext, "cy") ?? 0);
  if (![left, top, width, height].every(Number.isFinite)) return undefined;
  const rotation = Number(attribute(node, "rot") ?? 0) / 60_000;
  const flipH = attribute(node, "flipH") === "1" || attribute(node, "flipH") === "true", flipV = attribute(node, "flipV") === "1" || attribute(node, "flipV") === "true";
  return { left, top, width, height, ...(rotation ? { rotation } : {}), ...(flipH ? { flipH } : {}), ...(flipV ? { flipV } : {}) };
}
function resolvePptSchemeColor(value: string, theme: PptTheme, colorMap: Record<string, string>): string | undefined {
  const mapped = colorMap[value] ?? value, color = theme.colors[mapped];
  return color ? `#${color.replace(/^#/, "")}` : undefined;
}
function readPptColor(node: XmlNode | undefined, theme: PptTheme, colorMap: Record<string, string>): { color: string; opacity: number } | undefined {
  if (!node) return undefined;
  let value = "", opacity = 1;
  if (node.localName === "schemeClr") value = resolvePptSchemeColor(attribute(node, "val") ?? "", theme, colorMap) ?? "#000000";
  else if (node.localName === "srgbClr") value = `#${(attribute(node, "val") ?? "000000").replace(/^#/, "")}`;
  else if (node.localName === "sysClr") value = `#${(attribute(node, "lastClr") ?? (attribute(node, "val") === "window" ? "FFFFFF" : "000000")).replace(/^#/, "")}`;
  else if (node.localName === "prstClr") value = pptPresetColors[attribute(node, "val") ?? ""] ?? "#000000";
  else if (node.localName === "scrgbClr") value = `#${["r", "g", "b"].map(key => linearToSrgb(Number(attribute(node, key) ?? 0) / 100_000).toString(16).padStart(2, "0")).join("")}`;
  else if (node.localName === "hslClr") value = `#${hslToRgb(Number(attribute(node, "hue") ?? 0) / 21_600_000, Number(attribute(node, "sat") ?? 0) / 100_000, Number(attribute(node, "lum") ?? 0) / 100_000).map(channel => channel.toString(16).padStart(2, "0")).join("")}`;
  else return undefined;
  let rgb = parsePptRgb(value);
  // PowerPoint applies tint/shade in linear RGB and luminance/saturation/hue changes in HSL.
  const clamp = (amount: number) => Math.max(0, Math.min(1, amount));
  const withHsl = (change: (hsl: [number, number, number]) => [number, number, number]) => { const [h, s, l] = change(rgbToHsl(rgb)); rgb = hslToRgb(((h % 1) + 1) % 1, clamp(s), clamp(l)); };
  for (const transform of node.children) if (typeof transform !== "string") {
    const raw = Number(attribute(transform, "val") ?? 0), amount = raw / 100_000;
    switch (transform.localName) {
      case "alpha": opacity *= clamp(amount); break;
      case "alphaMod": opacity = clamp(opacity * amount); break;
      case "alphaOff": opacity = clamp(opacity + amount); break;
      case "shade": rgb = rgb.map(channel => linearToSrgb(srgbToLinear(channel) * clamp(amount))) as [number, number, number]; break;
      case "tint": rgb = rgb.map(channel => linearToSrgb(1 - (1 - srgbToLinear(channel)) * clamp(amount))) as [number, number, number]; break;
      case "lumMod": withHsl(([h, s, l]) => [h, s, l * amount]); break;
      case "lumOff": withHsl(([h, s, l]) => [h, s, l + amount]); break;
      case "satMod": withHsl(([h, s, l]) => [h, s * amount, l]); break;
      case "satOff": withHsl(([h, s, l]) => [h, s + amount, l]); break;
      case "hueMod": withHsl(([h, s, l]) => [h * amount, s, l]); break;
      case "hueOff": withHsl(([h, s, l]) => [h + raw / 21_600_000, s, l]); break;
      case "gray": { const gray = Math.round(rgb[0] * 0.299 + rgb[1] * 0.587 + rgb[2] * 0.114); rgb = [gray, gray, gray]; break; }
      case "inv": rgb = rgb.map(channel => 255 - channel) as [number, number, number]; break;
      case "comp": withHsl(([h, s, l]) => [h + 0.5, s, l]); break;
    }
  }
  return { color: `#${rgb.map(channel => channel.toString(16).padStart(2, "0")).join("")}`, opacity };
}
const pptPresetColors: Record<string, string> = { black: "#000000", white: "#ffffff", red: "#ff0000", green: "#008000", blue: "#0000ff", yellow: "#ffff00", gray: "#808080", grey: "#808080", darkGray: "#a9a9a9", lightGray: "#d3d3d3", orange: "#ffa500", purple: "#800080", cyan: "#00ffff", magenta: "#ff00ff", navy: "#000080", silver: "#c0c0c0", maroon: "#800000", lime: "#00ff00", teal: "#008080", olive: "#808000" };
function srgbToLinear(channel: number): number { const value = channel / 255; return value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4; }
function linearToSrgb(value: number): number { const v = Math.max(0, Math.min(1, value)); return Math.round((v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055) * 255); }
function rgbToHsl([r, g, b]: [number, number, number]): [number, number, number] {
  const rr = r / 255, gg = g / 255, bb = b / 255, max = Math.max(rr, gg, bb), min = Math.min(rr, gg, bb), l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min, s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === rr ? (gg - bb) / d + (gg < bb ? 6 : 0) : max === gg ? (bb - rr) / d + 2 : (rr - gg) / d + 4;
  return [h / 6, s, l];
}
function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (!s) { const gray = Math.round(l * 255); return [gray, gray, gray]; }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  const channel = (t: number) => { const x = ((t % 1) + 1) % 1; return x < 1 / 6 ? p + (q - p) * 6 * x : x < 1 / 2 ? q : x < 2 / 3 ? p + (q - p) * (2 / 3 - x) * 6 : p; };
  return [channel(h + 1 / 3), channel(h), channel(h - 1 / 3)].map(value => Math.round(Math.max(0, Math.min(1, value)) * 255)) as [number, number, number];
}
function parsePptRgb(color: string): [number, number, number] {
  const hex = color.replace(/^#/, "");
  return [Number.parseInt(hex.slice(0, 2), 16) || 0, Number.parseInt(hex.slice(2, 4), 16) || 0, Number.parseInt(hex.slice(4, 6), 16) || 0];
}
function readPptFill(node: XmlNode | undefined, theme: PptTheme, colorMap: Record<string, string>): PptxFill | undefined {
  if (!node) return undefined;
  const noFill = child(node, "noFill"); if (noFill) return undefined;
  const solid = child(node, "solidFill");
  if (solid) {
    const colorNode = solid.children.find((part): part is XmlNode => typeof part !== "string");
    const color = readPptColor(colorNode, theme, colorMap);
    return color ? { color: color.color, ...(color.opacity < 1 ? { opacity: color.opacity } : {}) } : undefined;
  }
  const pattern = child(node, "pattFill");
  if (pattern) {
    // Approximate a pattern by its foreground colour; the pattern texture itself is not modelled.
    const colorNode = child(pattern, "fgClr")?.children.find((part): part is XmlNode => typeof part !== "string");
    const color = readPptColor(colorNode, theme, colorMap);
    return color ? { color: color.color, ...(color.opacity < 1 ? { opacity: color.opacity } : {}) } : undefined;
  }
  const gradient = child(node, "gradFill"), stops = child(gradient, "gsLst");
  if (gradient && stops) {
    const colors = elements(stops, "gs").map(stop => {
      const colorNode = stop.children.find((part): part is XmlNode => typeof part !== "string" && pptColorElementNames.includes(part.localName));
      const color = readPptColor(colorNode, theme, colorMap), position = Number(attribute(stop, "pos") ?? 0) / 1000;
      return color ? { css: `${color.opacity < 1 ? `color-mix(in srgb, ${color.color} ${color.opacity * 100}%, transparent)` : color.color} ${position}%`, position } : undefined;
    }).filter((stop): stop is { css: string; position: number } => Boolean(stop)).sort((a, b) => a.position - b.position).map(stop => stop.css);
    if (colors.length === 1) colors.push(colors[0]!.replace(/[\d.]+%$/, "100%"));
    const path = child(gradient, "path");
    if (path && colors.length >= 2) {
      // Path gradients radiate from the fill-to rectangle; position 0 is the centre colour.
      const focus = attrs(child(path, "fillToRect")), x = (Number(focus["l"] ?? 50_000) + 100_000 - Number(focus["r"] ?? 50_000)) / 2_000, y = (Number(focus["t"] ?? 50_000) + 100_000 - Number(focus["b"] ?? 50_000)) / 2_000;
      return { gradient: `radial-gradient(${attribute(path, "path") === "circle" ? "circle" : "ellipse"} farthest-corner at ${x}% ${y}%, ${colors.join(", ")})` };
    }
    // DrawingML angles run clockwise from the x-axis; CSS angles run clockwise from "to top".
    const angle = Number(attribute(child(gradient, "lin"), "ang") ?? 0) / 60_000;
    if (colors.length >= 2) return { gradient: `linear-gradient(${angle + 90}deg, ${colors.join(", ")})` };
  }
  return undefined;
}
function readPptLine(node: XmlNode | undefined, theme: PptTheme, colorMap: Record<string, string>): PptxLine | undefined {
  if (!node || child(node, "noFill")) return undefined;
  const fill = readPptFill(node, theme, colorMap);
  const plain = fill && "color" in fill ? fill : fill && "gradient" in fill ? { color: /#[\da-f]{6}/i.exec(fill.gradient)?.[0] ?? "#000000", opacity: undefined } : undefined;
  return plain ? { color: plain.color, ...(plain.opacity !== undefined ? { opacity: plain.opacity } : {}), width: Math.max(0, Number(attribute(node, "w") ?? 12700)), style: attribute(child(node, "prstDash"), "val") ?? "solid", ...readPptLineDetails(node) } : undefined;
}
function readPptLineDetails(node: XmlNode): Partial<PptxLine> {
  const cap = attribute(node, "cap"), result: Partial<PptxLine> = {};
  if (cap === "rnd" || cap === "sq" || cap === "flat") result.cap = cap === "rnd" ? "round" : cap === "sq" ? "square" : "flat";
  for (const [name, key] of [["headEnd", "headEnd"], ["tailEnd", "tailEnd"]] as const) {
    const end = child(node, name), type = attribute(end, "type");
    if (type && type !== "none") result[key] = { type, ...(attribute(end, "w") ? { width: attribute(end, "w") as "sm" | "med" | "lg" } : {}), ...(attribute(end, "len") ? { length: attribute(end, "len") as "sm" | "med" | "lg" } : {}) };
  }
  return result;
}
/** Resolves a shape outline from (lowest priority first) theme line style, layout and own <a:ln>; later nodes override only what they specify. */
function resolvePptLine(sources: Array<{ node: XmlNode | undefined; theme: PptTheme }>, colorMap: Record<string, string>): PptxLine | undefined {
  let result: PptxLine | undefined;
  for (const { node, theme } of sources) {
    if (!node) continue;
    if (child(node, "noFill")) { result = undefined; continue; }
    const full = readPptLine(node, theme, colorMap);
    if (full) { result = { ...result, ...full }; continue; }
    if (!result) continue;
    const width = attribute(node, "w"), dash = attribute(child(node, "prstDash"), "val");
    result = { ...result, ...(width !== undefined ? { width: Math.max(0, Number(width)) } : {}), ...(dash ? { style: dash } : {}), ...readPptLineDetails(node) };
  }
  return result;
}
const pptColorElementNames = ["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"];
/** Theme with phClr bound to the colour of a style reference (fillRef/lnRef/effectRef/bgRef). */
function pptStyleTheme(reference: XmlNode | undefined, theme: PptTheme, colorMap: Record<string, string>): PptTheme {
  const colorNode = reference?.children.find((part): part is XmlNode => typeof part !== "string" && pptColorElementNames.includes(part.localName));
  const color = readPptColor(colorNode, theme, colorMap);
  return color ? { ...theme, colors: { ...theme.colors, phClr: color.color.slice(1) } } : theme;
}
function pptStyleEntry(list: XmlNode[], index: number): XmlNode | undefined { return index >= 1 ? list[index - 1] : undefined; }
/** Wraps a bare fill element (e.g. a theme fill style) so readPptFill can read it like a shape's properties. */
function pptFillHolder(fill: XmlNode | undefined): XmlNode | undefined { return fill ? { name: "holder", localName: "holder", attributes: {}, children: [fill] } : undefined; }
function readPptStyleFill(reference: XmlNode | undefined, theme: PptTheme, colorMap: Record<string, string>): PptxFill | undefined {
  const index = Number(attribute(reference, "idx") ?? 0);
  const fill = index >= 1001 ? pptStyleEntry(theme.backgroundFillStyles, index - 1000) : pptStyleEntry(theme.fillStyles, index);
  return fill ? readPptFill(pptFillHolder(fill), pptStyleTheme(reference, theme, colorMap), colorMap) : undefined;
}
function readPptShadow(node: XmlNode | undefined, theme: PptTheme, colorMap: Record<string, string>): PptxShadow | undefined {
  const shadow = child(child(node, "effectLst"), "outerShdw");
  if (!shadow) return undefined;
  const colorNode = shadow.children.find((part): part is XmlNode => typeof part !== "string" && pptColorElementNames.includes(part.localName));
  const color = readPptColor(colorNode, theme, colorMap);
  if (!color) return undefined;
  const distance = Math.max(0, Number(attribute(shadow, "dist") ?? 0)), angle = Number(attribute(shadow, "dir") ?? 0) / 60_000 * Math.PI / 180;
  return { x: Math.cos(angle) * distance, y: Math.sin(angle) * distance, blur: Math.max(0, Number(attribute(shadow, "blurRad") ?? 0)), color: color.color, ...(color.opacity < 1 ? { opacity: color.opacity } : {}) };
}
function parsePptTextBody(body: XmlNode, box: PptTransform, defaultRun: Partial<TextRun>, palette: (node: XmlNode | undefined) => { color: string; opacity: number } | undefined, theme: PptTheme, colorMap: Record<string, string>, inheritedBody?: XmlNode, masterParagraph?: XmlNode, masterBody?: XmlNode): PptxTextElement {
  const props = child(body, "bodyPr"), inheritedProps = child(inheritedBody, "bodyPr"), masterProps = child(masterBody, "bodyPr");
  const value = (key: string) => attribute(props, key) ?? attribute(inheritedProps, key) ?? attribute(masterProps, key);
  const getMargin = (key: string, fallback: number) => Number(value(key) ?? fallback);
  const currentListStyle = child(body, "lstStyle"), inheritedListStyle = child(inheritedBody, "lstStyle");
  const listStyle = currentListStyle && currentListStyle.children.some(part => typeof part !== "string" && /^lvl[1-9]pPr$/.test(part.localName)) ? currentListStyle : inheritedListStyle ?? currentListStyle;
  const vertical = value("anchor");
  const verticalAlign = vertical === "ctr" ? "center" : vertical === "b" ? "bottom" : "top";
  const inheritedParagraph = elements(inheritedBody ?? body, "p")[0];
  const paragraphs = elements(body, "p").map(p => parsePptParagraph(p, listStyle, defaultRun, palette, theme, colorMap, inheritedParagraph, masterParagraph, Boolean(inheritedBody)));
  // normAutofit records the shrink PowerPoint applied when the text last overflowed its box.
  const autofit = child(props, "normAutofit") ?? (child(props, "noAutofit") || child(props, "spAutoFit") ? undefined : child(inheritedProps, "normAutofit"));
  const fontScale = Number(attribute(autofit, "fontScale") ?? 100_000) / 100_000, spacingReduction = Number(attribute(autofit, "lnSpcReduction") ?? 0) / 100_000;
  if (fontScale > 0 && fontScale < 1 || spacingReduction > 0) for (const paragraph of paragraphs) {
    const scale = (run: Partial<TextRun> | undefined) => { if (run?.fontSize && fontScale > 0 && fontScale < 1) run.fontSize = Math.round(run.fontSize * fontScale * 10) / 10; };
    for (const run of paragraph.runs) scale(run);
    if (paragraph.defaultRun) { paragraph.defaultRun = { ...paragraph.defaultRun }; scale(paragraph.defaultRun); }
    if (spacingReduction > 0 && paragraph.lineHeightRule !== "exact") paragraph.lineHeight = Math.max(0.1, (paragraph.lineHeight ?? 1) - spacingReduction);
  }
  const { left, top, width, height, rotation } = box, warp = attribute(child(props, "prstTxWarp"), "prst");
  return { kind: "text", left, top, width, height, ...(rotation ? { rotation } : {}), verticalAlign, wrap: value("wrap") !== "none", margins: { left: getMargin("lIns", 91440), right: getMargin("rIns", 91440), top: getMargin("tIns", 45720), bottom: getMargin("bIns", 45720) }, paragraphs, ...(warp && warp !== "textNoShape" ? { warp } : {}) };
}
function parsePptParagraph(node: XmlNode, listStyle: XmlNode | undefined, defaultRun: Partial<TextRun>, palette: (node: XmlNode | undefined) => { color: string; opacity: number } | undefined, _theme: PptTheme, _colorMap: Record<string, string>, inheritedParagraph?: XmlNode, masterParagraph?: XmlNode, hasLayoutParagraphStyle = false): Paragraph {
  const props = child(node, "pPr"), level = Math.max(0, Math.min(8, Number(attribute(props, "lvl") ?? 0)));
  const inheritedProps = child(inheritedParagraph, "pPr");
  const masterLevel = child(masterParagraph, `lvl${level + 1}pPr`) ?? child(masterParagraph, "lvl1pPr") ?? masterParagraph;
  const masterLevelRun = level > 0 ? readPptRunStyle(child(masterLevel, "defRPr"), defaultRun.fontFamily, palette, _theme) : {};
  const levelProps = child(listStyle, `lvl${level + 1}pPr`), levelRun = readPptRunStyle(child(levelProps, "defRPr"), defaultRun.fontFamily, palette, _theme);
  const inheritedRun = readPptRunStyle(child(inheritedProps, "defRPr"), defaultRun.fontFamily, palette, _theme);
  const paraDefault = { ...defaultRun, ...masterLevelRun, ...inheritedRun, ...levelRun, ...readPptRunStyle(child(props, "defRPr"), defaultRun.fontFamily, palette, _theme) };
  const runs: TextRun[] = [];
  for (const part of node.children) if (typeof part !== "string") {
    if (part.localName === "r" || part.localName === "fld") {
      const text = child(part, "t"); if (text) runs.push({ ...paraDefault, ...readPptRunStyle(child(part, "rPr"), paraDefault.fontFamily, palette, _theme), text: textContent(text) });
    } else if (part.localName === "br") runs.push({ ...paraDefault, text: "\n" });
  }
  const endStyle = child(node, "endParaRPr");
  if (!runs.length && endStyle) runs.push({ ...paraDefault, ...readPptRunStyle(endStyle, paraDefault.fontFamily, palette), text: "" });
  const masterProps = child(masterLevel, "pPr") ?? masterLevel;
  const alignment = attribute(props, "algn") ?? attribute(levelProps, "algn") ?? attribute(inheritedProps, "algn") ?? attribute(masterProps, "algn");
  // The renderer drops space before the first paragraph of a text body, as PowerPoint does, so master spacing can apply here.
  const before = pptSpacing(child(props, "spcBef")) ?? pptSpacing(child(levelProps, "spcBef")) ?? pptSpacing(child(inheritedProps, "spcBef")) ?? pptSpacing(child(masterProps, "spcBef")), after = pptSpacing(child(props, "spcAft")) ?? pptSpacing(child(levelProps, "spcAft")) ?? pptSpacing(child(inheritedProps, "spcAft")) ?? pptSpacing(child(masterProps, "spcAft"));
  const lineSpacing = child(props, "lnSpc") ?? child(levelProps, "lnSpc") ?? child(inheritedProps, "lnSpc") ?? child(masterProps, "lnSpc"), pct = Number(attribute(child(lineSpacing, "spcPct"), "val") ?? 0), pts = Number(attribute(child(lineSpacing, "spcPts"), "val") ?? 0);
  const left = Number(attribute(props, "marL") ?? attribute(levelProps, "marL") ?? attribute(inheritedProps, "marL") ?? attribute(masterProps, "marL") ?? 0) / 12700;
  const right = Number(attribute(props, "marR") ?? attribute(levelProps, "marR") ?? attribute(inheritedProps, "marR") ?? attribute(masterProps, "marR") ?? 0) / 12700;
  const indent = Number(attribute(props, "indent") ?? attribute(levelProps, "indent") ?? attribute(inheritedProps, "indent") ?? attribute(masterProps, "indent") ?? 0) / 12700;
  const bullet = child(props, "buChar") ?? child(levelProps, "buChar") ?? child(inheritedProps, "buChar") ?? child(masterProps, "buChar"), auto = child(props, "buAutoNum") ?? child(levelProps, "buAutoNum") ?? child(inheritedProps, "buAutoNum") ?? child(masterProps, "buAutoNum"), noBullet = child(props, "buNone") ?? child(levelProps, "buNone") ?? child(inheritedProps, "buNone") ?? child(masterProps, "buNone");
  const bulletChar = attribute(bullet, "char");
  const autoType = attribute(auto, "type") ?? "arabicPeriod";
  const list = !noBullet && (bulletChar || auto) ? { label: bulletChar ?? (autoType.toLowerCase().includes("roman") ? "I." : "1."), level, kind: auto ? "ordered" as const : "bullet" as const, indentLeft: left, hanging: Math.max(0, -indent) } : undefined;
  return {
    runs, defaultRun: paraDefault,
    ...(alignment && ["l", "ctr", "r", "just", "dist"].includes(alignment) ? { alignment: alignment === "ctr" ? "center" : alignment === "r" ? "right" : alignment === "just" || alignment === "dist" ? "justify" : "left" } : {}),
    ...(before !== undefined ? { spacingBefore: before } : {}), ...(after !== undefined ? { spacingAfter: after } : {}),
    ...(pct ? { lineHeight: pct / 100_000 } : pts ? { lineHeight: pts / 100, lineHeightRule: "exact" as const } : {}),
    ...(left ? { indentLeft: left } : {}), ...(right ? { indentRight: right } : {}), ...(indent ? { indentFirstLine: indent } : {}), ...(list ? { list } : {}),
  };
}
function pptSpacing(node: XmlNode | undefined): number | undefined {
  const pts = child(node, "spcPts"), pct = child(node, "spcPct");
  if (pts) return Number(attribute(pts, "val") ?? 0) / 100;
  if (pct) return Number(attribute(pct, "val") ?? 0) / 100_000 * 12;
  return undefined;
}
function readPptRunStyle(node: XmlNode | undefined, fallbackFont: string | undefined, palette: (node: XmlNode | undefined) => { color: string; opacity: number } | undefined, theme?: PptTheme): Partial<TextRun> {
  if (!node) return {};
  const a = attrs(node), result: Partial<TextRun> = {};
  if (a["sz"] !== undefined) result.fontSize = Number(a["sz"]) / 100;
  if (a["b"] !== undefined) result.bold = a["b"] === "1";
  if (a["i"] !== undefined) result.italic = a["i"] === "1";
  if (a["u"] !== undefined) result.underline = !["none", "0"].includes(a["u"]!);
  if (a["strike"] !== undefined) result.strike = !["noStrike", "none", "0"].includes(a["strike"]!);
  if (a["cap"] === "all") result.caps = true; else if (a["cap"] === "small") result.smallCaps = true;
  if (a["spc"] !== undefined) result.letterSpacing = Number(a["spc"]) / 100;
  const latin = child(node, "latin"), face = attribute(latin, "typeface");
  if (face && !face.startsWith("+")) result.fontFamily = face;
  else if (face?.startsWith("+mj") && theme) result.fontFamily = theme.majorFont;
  else if (face?.startsWith("+mn") && theme) result.fontFamily = theme.minorFont;
  else if (fallbackFont) result.fontFamily = fallbackFont;
  const solid = child(node, "solidFill"), colorNode = solid?.children.find((part): part is XmlNode => typeof part !== "string");
  const color = palette(colorNode);
  if (color) result.color = color.opacity < 1 ? `rgba(${parsePptRgb(color.color).join(",")},${color.opacity})` : color.color;
  const shadow = child(child(node, "effectLst"), "outerShdw"), shadowColor = palette(shadow?.children.find((part): part is XmlNode => typeof part !== "string" && pptColorElementNames.includes(part.localName)));
  if (shadow && shadowColor) {
    const distance = Number(attribute(shadow, "dist") ?? 0) / 12_700, angle = Number(attribute(shadow, "dir") ?? 0) / 60_000 * Math.PI / 180;
    result.shadow = { x: Math.cos(angle) * distance, y: Math.sin(angle) * distance, blur: Number(attribute(shadow, "blurRad") ?? 0) / 12_700, color: shadowColor.color, ...(shadowColor.opacity < 1 ? { opacity: shadowColor.opacity } : {}) };
  }
  return result;
}
function parsePptTable(frame: XmlNode, theme: PptTheme, colorMap: Record<string, string>, palette: (node: XmlNode | undefined) => { color: string; opacity: number } | undefined, groupTransform: PptGroupTransform = { scaleX: 1, scaleY: 1, translateX: 0, translateY: 0, rotation: 0 }): PptxTableElement | undefined {
  const box = transformPptTransform(readPptTransform(child(frame, "xfrm")), groupTransform), table = descendants(frame, "tbl")[0];
  if (!box || !table) return undefined;
  const columns = elements(child(table, "tblGrid") ?? table, "gridCol").map(col => Number(attribute(col, "w") ?? 0));
  const rows = elements(table, "tr").map(row => ({
    height: Number(attribute(row, "h") ?? 0),
    cells: elements(row, "tc").map(cell => {
      const props = child(cell, "tcPr"), fill = readPptFill(props, theme, colorMap), propsAttrs = attrs(props);
      const body = child(cell, "txBody"), textBody = body ? parsePptTextBody(body, box, { fontFamily: theme.minorFont, color: resolvePptSchemeColor("tx1", theme, colorMap) ?? "#000000" }, palette, theme, colorMap) : undefined;
      const borders: PptxTableElement["rows"][number]["cells"][number]["borders"] = {};
      for (const side of ["top", "right", "bottom", "left"] as const) {
        const line = readPptLine(child(props, `ln${side[0]!.toUpperCase()}${side.slice(1)}`), theme, colorMap);
        if (line) borders[side] = line;
      }
      const anchor = attribute(props, "anchor");
      const align = textBody?.paragraphs.find(paragraph => paragraph.alignment)?.alignment;
      return {
        paragraphs: textBody?.paragraphs ?? [], ...(fill ? { fill } : {}),
        verticalAlign: anchor === "ctr" ? "center" as const : anchor === "b" ? "bottom" as const : "top" as const,
        ...(align && align !== "justify" ? { alignment: align } : {}),
        margins: { left: Number(propsAttrs["marL"] ?? 91440), right: Number(propsAttrs["marR"] ?? 91440), top: Number(propsAttrs["marT"] ?? 45720), bottom: Number(propsAttrs["marB"] ?? 45720) },
        ...(Object.keys(borders).length ? { borders } : {}),
      };
    }),
  }));
  return { kind: "table", ...box, columns, rows };
}

function parseRelationships(zip: ZipArchive, path: string, base: string): Map<string, string> {
  const result = new Map<string, string>();
  if (!zip.has(path)) return result;
  const root = decodeXml(zip, path);
  for (const rel of elements(root, "Relationship")) {
    const id = attrs(rel)["Id"], target = attrs(rel)["Target"], mode = attrs(rel)["TargetMode"];
    if (!id || !target || mode === "External") continue;
    const joined = target.startsWith("/") ? target.slice(1) : `${base}/${target}`;
    const parts: string[] = [];
    for (const part of joined.split("/")) {
      if (part === "..") { if (!parts.length) throw new PreviewError("OOXML relationship escapes its package.", "invalid-file"); parts.pop(); }
      else if (part && part !== ".") parts.push(part);
    }
    result.set(id, parts.join("/"));
  }
  return result;
}
