import type { FileFormat } from "../detection/detect-format.js";
import { parseDocx, parsePptx } from "../parsers/ooxml.js";
import type { Paragraph, Presentation, PptxElement, PptxFill, PptxImageElement, PptxLine, PptxLineElement, PptxShapeElement, PptxTextElement, TextRun, WordPageLayout } from "../model.js";
import { fontStack } from "./fonts.js";
import { pptFillCss, renderPptxElement } from "./drawingml-renderer.js";
import type { FileRenderer, RendererContext, RendererInstance } from "./types.js";

export class DocxRenderer implements FileRenderer {
  readonly formats: readonly FileFormat[] = ["docx"];
  render({ bytes, file }: RendererContext): RendererInstance {
    const model = parseDocx(bytes, file.name), root = document.createElement("div");
    root.className = "ufp-word-document";
    const imageUrls: string[] = [];
    const pageWidth = model.page ? model.page.width / 72 * 96 : 816;
    const pageHeight = model.page ? model.page.height / 72 * 96 : 1056;
    const pageMargins = model.page ? [model.page.marginTop, model.page.marginRight, model.page.marginBottom, model.page.marginLeft].map(value => `${value}pt`).join(" ") : "72pt";
    root.style.fontFamily = model.defaultRun?.fontFamily ? fontStack(model.defaultRun.fontFamily) : "Cambria, serif";
    root.style.fontSize = `${model.defaultRun?.fontSize ?? 11}pt`;
    root.style.lineHeight = "normal";
    if (model.defaultRun?.color) root.style.color = model.defaultRun.color;
    const appendBlocks = (target: HTMLElement, blocks: typeof model.blocks, currentPageNumber = 1) => {
      for (const block of blocks) {
        if (block.type === "paragraph") target.append(renderParagraph(block.value, imageUrls, model.page, currentPageNumber));
        else if (block.type === "table") target.append(renderTable(block.value, imageUrls, model.page, currentPageNumber));
      }
    };
    let pageNumber = 0;
    function createPage(): HTMLElement {
      const firstPage = pageNumber++ === 0;
      const currentPageNumber = pageNumber;
      const pageHeaders = firstPage && model.differentFirstPageHeaderFooter ? model.firstPageHeaders : model.headers;
      const pageFooters = firstPage && model.differentFirstPageHeaderFooter ? model.firstPageFooters : model.footers;
      const current = document.createElement("article"); current.className = "ufp-word-page";
      current.style.width = `${pageWidth}px`; current.style.minHeight = `${pageHeight}px`; current.style.padding = pageMargins; current.style.position = "relative";
      current.style.isolation = "isolate";
      current.style.color = model.defaultRun?.color ?? "#111";
      root.append(current);
      if (pageHeaders?.length) {
        const header = document.createElement("div"); header.className = "ufp-word-header";
        header.style.position = "absolute"; header.style.top = `${model.page?.headerDistance ?? 36}pt`;
        header.style.left = `${model.page?.marginLeft ?? 0}pt`; header.style.right = `${model.page?.marginRight ?? 0}pt`; header.style.margin = "0";
        appendBlocks(header, pageHeaders, currentPageNumber); current.append(header);
      }
      const body = document.createElement("div"); body.className = "ufp-word-body"; current.append(body);
      if (pageFooters?.length) {
        const footer = document.createElement("div"); footer.className = "ufp-word-footer";
        footer.style.position = "absolute"; footer.style.bottom = `${model.page?.footerDistance ?? 36}pt`;
        footer.style.left = `${model.page?.marginLeft ?? 0}pt`; footer.style.right = `${model.page?.marginRight ?? 0}pt`; footer.style.margin = "0";
        appendBlocks(footer, pageFooters, currentPageNumber); current.append(footer);
      }
      return current;
    }
    let page = createPage();
    // Keep block content between the header and footer on each page.
    const bodyFor = (current: HTMLElement) => current.querySelector<HTMLElement>(".ufp-word-body")!;
    // `appendBlocks` uses this page's body, including after an explicit page break.
    const contentBlocks = model.blocks;
    let breakAfterBlock = false;
    for (const block of contentBlocks) {
      if (breakAfterBlock && bodyFor(page).childElementCount > 0) page = createPage();
      breakAfterBlock = false;
      if (block.type === "paragraph") {
        if (block.value.pageBreakBefore && bodyFor(page).childElementCount > 0) page = createPage();
        const segments = splitAtPageBreaks(block.value);
        for (const segment of segments.paragraphs) {
          if (segment.startsNewPage && bodyFor(page).childElementCount > 0) page = createPage();
        bodyFor(page).append(renderParagraph(segment.paragraph, imageUrls, model.page, pageNumber));
        }
        breakAfterBlock = segments.breakAfter;
      } else if (block.type === "table") bodyFor(page).append(renderTable(block.value, imageUrls, model.page, pageNumber));
    }
    return {
      element: root,
      setZoom(zoom) { root.style.setProperty("zoom", String(Math.max(0.1, Math.min(2, zoom)))); },
      fitToWidth() {
        const available = root.parentElement?.clientWidth ?? pageWidth;
        const scale = Math.max(0.1, Math.min(2, (available - 32) / pageWidth));
        if (scale === 1) root.style.removeProperty("zoom"); else root.style.setProperty("zoom", String(scale));
      },
      destroy() { for (const url of imageUrls) URL.revokeObjectURL(url); root.replaceChildren(); },
    };
  }
}

function renderParagraph(paragraph: Paragraph, imageUrls: string[], page?: WordPageLayout, pageNumber = 1): HTMLParagraphElement {
  const p = document.createElement(paragraph.kind === "heading" ? "h2" : "p");
  p.style.textAlign = paragraph.alignment ?? "left";
  p.style.fontSize = "inherit"; p.style.fontWeight = "inherit";
  if (paragraph.defaultRun?.fontFamily) p.style.fontFamily = fontStack(paragraph.defaultRun.fontFamily);
  if (paragraph.defaultRun?.fontSize) p.style.fontSize = `${Math.max(1, Math.min(96, paragraph.defaultRun.fontSize))}pt`;
  if (paragraph.defaultRun?.color && /^#[0-9a-f]{6}$/i.test(paragraph.defaultRun.color)) p.style.color = paragraph.defaultRun.color;
  if (paragraph.defaultRun?.bold) p.style.fontWeight = "700";
  if (paragraph.defaultRun?.italic) p.style.fontStyle = "italic";
  if (paragraph.defaultRun?.caps) p.style.textTransform = "uppercase";
  if (paragraph.defaultRun?.smallCaps) p.style.fontVariant = "small-caps";
  p.style.marginTop = `${paragraph.spacingBefore ?? 0}pt`;
  p.style.marginBottom = `${paragraph.spacingAfter ?? 0}pt`;
  if (paragraph.lineHeight !== undefined) {
    if (paragraph.lineHeightRule === "exact" || paragraph.lineHeightRule === "atLeast") {
      const points = paragraph.lineHeightRule === "atLeast" ? Math.max(paragraph.lineHeight, paragraph.defaultRun?.fontSize ?? 0) : paragraph.lineHeight;
      p.style.lineHeight = `${points}pt`;
    } else if (paragraph.lineHeightRule === "auto" && Math.abs(paragraph.lineHeight - 1) < 0.001) p.style.lineHeight = "normal";
    else p.style.lineHeight = String(paragraph.lineHeight);
  }
  if (paragraph.background && /^#[0-9a-f]{6}$/i.test(paragraph.background)) p.style.backgroundColor = paragraph.background;
  if (paragraph.borders) {
    for (const [edge, border] of Object.entries(paragraph.borders)) {
      if (!border) continue;
      const cssEdge = edge === "between" ? "bottom" : edge;
      p.style.setProperty(`border-${cssEdge}`, `${border.width}px ${border.style} ${border.color}`);
    }
  }
  if (paragraph.borderSpacing) {
    for (const [edge, spacing] of Object.entries(paragraph.borderSpacing)) {
      if (edge === "between" || spacing === undefined) continue;
      p.style.setProperty(`padding-${edge}`, `${Math.max(0, spacing)}pt`);
    }
  }
  if (paragraph.indentLeft !== undefined) p.style.marginLeft = paragraph.background && /^#[0-9a-f]{6}$/i.test(paragraph.background) ? "0" : `${paragraph.indentLeft}pt`;
  if (paragraph.indentRight !== undefined) p.style.marginRight = `${paragraph.indentRight}pt`;
  if (paragraph.indentFirstLine !== undefined) p.style.textIndent = `${paragraph.indentFirstLine}pt`;
  if (paragraph.list) {
    const list = paragraph.list;
    p.classList.add("ufp-word-list-item");
    const hasIndent = list.indentLeft !== undefined || list.hanging !== undefined || list.tabPosition !== undefined;
    const hangingLayout = list.hanging !== undefined || list.tabPosition !== undefined;
    const fallbackIndent = Math.max(18, (list.level + 1) * 18) * 0.75;
    const indentLeft = hasIndent ? paragraph.indentLeft ?? list.indentLeft ?? fallbackIndent : fallbackIndent;
    p.style.marginLeft = paragraph.background && /^#[0-9a-f]{6}$/i.test(paragraph.background) ? "0" : `${indentLeft}pt`;
    const label = document.createElement("span"); label.className = "ufp-word-list-label";
    if (hangingLayout) {
      const hanging = paragraph.indentFirstLine !== undefined ? Math.max(0, -paragraph.indentFirstLine) : list.hanging ?? 0;
      if (paragraph.indentFirstLine === undefined && hanging > 0) p.style.textIndent = `${-hanging}pt`;
      label.style.display = "inline-block";
      label.style.width = `${Math.max(0, list.tabPosition ?? list.hanging ?? 0)}pt`;
      // The paragraph's negative first-line indent also reaches inline text
      // inside this label. Reset it so the marker stays at the list position.
      label.style.textIndent = "0";
      applyListLabelStyle(label, list.labelStyle);
      label.textContent = list.label || (list.kind === "bullet" ? "•" : "");
    } else {
      applyListLabelStyle(label, list.labelStyle);
      label.textContent = `${list.label || (list.kind === "bullet" ? "•" : "")} `;
    }
    p.append(label);
  }
  for (const run of paragraph.runs) {
    const rendered = renderRun(run, imageUrls, pageNumber);
    if (run.drawing) {
      p.style.position = "relative";
      rendered.style.left = `${run.drawing.left - (paragraph.indentLeft ?? 0) * 96 / 72}px`;
      rendered.style.top = `${run.drawing.top}px`;
      if (page) {
        const pageWidth = page.width / 72 * 96, pageLeft = page.marginLeft / 72 * 96 + run.drawing.left;
        const clipRight = Math.max(0, pageLeft + run.drawing.width - pageWidth);
        if (clipRight > 0) rendered.style.clipPath = `inset(0 ${clipRight}px 0 0)`;
      }
    }
    p.append(rendered);
  }
  if (!paragraph.runs.length) p.append(document.createElement("br"));
  return p;
}
function renderRun(run: TextRun, imageUrls: string[] = [], pageNumber = 1): HTMLElement | SVGSVGElement {
  if (run.pageNumber) return renderRun({ ...run, text: String(pageNumber), pageNumber: false }, imageUrls, pageNumber);
  if (run.pageBreak) return document.createElement("br");
  if (run.drawing) return renderWordDrawing(run.drawing);
  if (run.image) {
    const image = document.createElement("img");
    const buffer = new Uint8Array(run.image.bytes.byteLength); buffer.set(run.image.bytes);
    const url = URL.createObjectURL(new Blob([buffer.buffer], { type: run.image.mimeType })); imageUrls.push(url);
    image.src = url; image.alt = run.image.alt ?? ""; image.className = "ufp-word-image";
    if (run.image.width) image.style.width = `${run.image.width}px`;
    if (run.image.height) image.style.height = `${run.image.height}px`;
    return image;
  }
  const span = document.createElement("span");
  if (run.complexScriptFontSize) appendScriptSizedText(span, run.text, run.complexScriptFontSize);
  else span.textContent = run.text;
  span.style.fontWeight = run.bold ? "700" : ""; span.style.fontStyle = run.italic ? "italic" : "";
  if (run.caps) span.style.textTransform = "uppercase";
  if (run.smallCaps) span.style.fontVariant = "small-caps";
  span.style.textDecoration = [run.underline ? "underline" : "", run.strike ? "line-through" : ""].filter(Boolean).join(" ");
  if (run.underline && run.underlineStyle) span.style.textDecorationStyle = run.underlineStyle;
  if (run.underline && run.underlineColor && /^#[0-9a-f]{6}$/i.test(run.underlineColor)) span.style.textDecorationColor = run.underlineColor;
  if (run.fontSize) span.style.fontSize = `${Math.max(1, Math.min(96, run.fontSize))}pt`;
  if (run.fontFamily) span.style.fontFamily = fontStack(run.fontFamily);
  if (run.letterSpacing) span.style.letterSpacing = `${run.letterSpacing}pt`;
  if (run.color && /^#[0-9a-f]{6}$/i.test(run.color)) span.style.color = run.color;
  if (run.highlight && /^#[0-9a-f]{6}$/i.test(run.highlight)) span.style.backgroundColor = run.highlight;
  if (run.verticalAlign) { span.style.verticalAlign = run.verticalAlign === "baseline" ? "baseline" : run.verticalAlign; span.style.fontSize ||= "smaller"; }
  else if (run.position) { span.style.position = "relative"; span.style.top = `${-run.position}pt`; }
  return span;
}

function applyListLabelStyle(label: HTMLElement, style?: Partial<TextRun>): void {
  if (!style) return;
  // Legacy Symbol-numbering glyphs are remapped to Unicode during parsing.
  // Keep the text font for those remapped characters because modern browsers
  // may select Symbol but have no glyph for the Unicode bullet code point.
  if (style.fontFamily && !/^symbol$/i.test(style.fontFamily.trim())) label.style.fontFamily = fontStack(style.fontFamily);
  if (style.fontSize) label.style.fontSize = `${Math.max(1, Math.min(96, style.fontSize))}pt`;
  if (style.bold !== undefined) label.style.fontWeight = style.bold ? "700" : "400";
  if (style.italic !== undefined) label.style.fontStyle = style.italic ? "italic" : "normal";
  if (style.color && /^#[0-9a-f]{6}$/i.test(style.color)) label.style.color = style.color;
}

const complexScriptCharacter = /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Adlam}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gurmukhi}\p{Script=Gujarati}\p{Script=Oriya}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Sinhala}\p{Script=Khmer}\p{Script=Lao}\p{Script=Myanmar}\p{Script=Thai}]/u;
const combiningMark = /^\p{M}$/u;

function appendScriptSizedText(parent: HTMLElement, text: string, fontSize: number): void {
  let segment = "", complex = false;
  const flush = () => {
    if (!segment) return;
    if (complex) {
      const span = document.createElement("span"); span.textContent = segment; span.style.fontSize = `${fontSize}pt`; parent.append(span);
    } else parent.append(document.createTextNode(segment));
    segment = "";
  };
  for (const character of text) {
    const isComplex: boolean = complexScriptCharacter.test(character) || combiningMark.test(character) && complex;
    if (segment && isComplex !== complex) flush();
    if (!segment) complex = isComplex;
    segment += character;
  }
  flush();
}

function renderWordDrawing(drawing: NonNullable<TextRun["drawing"]>): SVGSVGElement {
  const namespace = "http://www.w3.org/2000/svg", svg = document.createElementNS(namespace, "svg");
  svg.classList.add("ufp-word-drawing"); svg.setAttribute("aria-hidden", "true"); svg.setAttribute("focusable", "false");
  svg.setAttribute("viewBox", `0 0 ${drawing.width * 9525} ${drawing.height * 9525}`);
  svg.setAttribute("preserveAspectRatio", "none");
  svg.style.position = "absolute"; svg.style.width = `${drawing.width}px`; svg.style.height = `${drawing.height}px`;
  svg.style.pointerEvents = "none"; svg.style.zIndex = drawing.behindText ? "-1" : "0";
  for (const shape of drawing.paths) {
    const path = document.createElementNS(namespace, "path"); path.setAttribute("d", shape.d); path.setAttribute("transform", shape.transform);
    path.setAttribute("fill", shape.fill ?? "none");
    if (shape.stroke) path.setAttribute("stroke", shape.stroke);
    if (shape.strokeWidth) path.setAttribute("stroke-width", String(shape.strokeWidth * 9525));
    svg.append(path);
  }
  return svg;
}

function splitAtPageBreaks(paragraph: Paragraph): { paragraphs: Array<{ paragraph: Paragraph; startsNewPage: boolean }>; breakAfter: boolean } {
  const paragraphs: Array<{ paragraph: Paragraph; startsNewPage: boolean }> = [];
  let runs: TextRun[] = [], startsNewPage = false;
  for (const run of paragraph.runs) {
    if (!run.pageBreak) { runs.push(run); continue; }
    if (runs.length) paragraphs.push({ paragraph: { ...paragraph, runs, pageBreakBefore: false }, startsNewPage });
    runs = []; startsNewPage = true;
  }
  if (runs.length || !paragraphs.length) paragraphs.push({ paragraph: { ...paragraph, runs, pageBreakBefore: false }, startsNewPage });
  const trailingBreak = paragraph.runs.at(-1)?.pageBreak === true;
  return { paragraphs, breakAfter: trailingBreak };
}

function renderTable(data: import("../model.js").Table, imageUrls: string[], page?: WordPageLayout, pageNumber = 1): HTMLTableElement {
  const table = document.createElement("table");
  if (data.width) table.style.width = `${data.width}pt`;
  else if (data.widthPercent) table.style.width = `${data.widthPercent}%`;
  table.style.margin = "0";
  if (data.alignment === "center") table.style.marginInline = "auto";
  if (data.alignment === "right") table.style.marginLeft = "auto";
  const columns = Math.max(0, ...data.rows.map(row => row.cells.reduce((sum, cell) => sum + (cell.colSpan ?? 1), 0)));
  data.rows.forEach((row, rowIndex) => {
    const tr = table.insertRow();
    if (row.height) tr.style.height = `${row.height}pt`;
    row.cells.forEach((cell, cellIndex) => {
      const column = cell.column ?? row.cells.slice(0, cellIndex).reduce((sum, previous) => sum + (previous.colSpan ?? 1), 0), span = cell.colSpan ?? 1;
      const td = tr.insertCell();
      if (span > 1) td.colSpan = span;
      if (cell.rowSpan && cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
      if (cell.width) td.style.width = `${cell.width}%`;
      if (cell.background && /^#[0-9a-f]{6}$/i.test(cell.background)) td.style.backgroundColor = cell.background;
      if (cell.verticalAlign) td.style.verticalAlign = cell.verticalAlign === "center" ? "middle" : cell.verticalAlign;
      const cellMargins = { ...data.cellMargins, ...cell.margins };
      td.style.padding = `${cellMargins.top ?? 3}pt ${cellMargins.right ?? 0}pt ${cellMargins.bottom ?? 0}pt ${cellMargins.left ?? 0}pt`;
      for (const [edge, borderKey] of [["top", "top"], ["left", "left"], ["bottom", "bottom"], ["right", "right"]] as const) {
        const outer = edge === "top" ? rowIndex === 0 : edge === "bottom" ? rowIndex === data.rows.length - 1 : edge === "left" ? column === 0 : column + span >= columns;
        const inner = edge === "bottom" && rowIndex < data.rows.length - 1 ? data.borders?.insideH : edge === "right" && column + span < columns ? data.borders?.insideV : undefined;
        const cellEdge = cell.borders?.[borderKey];
        const cellInside = edge === "bottom" && rowIndex < data.rows.length - 1 ? cell.borders?.insideH : edge === "right" && column + span < columns ? cell.borders?.insideV : undefined;
        const border = cellEdge ?? cellInside ?? (outer ? data.borders?.[borderKey] : inner);
        if (border) td.style.setProperty(`border-${edge}`, `${border.width}px ${border.style} ${border.color}`);
      }
      cell.paragraphs.forEach(paragraph => td.append(renderParagraph(paragraph, imageUrls, page, pageNumber)));
    });
  });
  return table;
}

export { XlsxRenderer } from "./xlsx-renderer.js";

export class PptxRenderer implements FileRenderer {
  readonly formats: readonly FileFormat[] = ["pptx"];
  render({ bytes, file }: RendererContext): RendererInstance {
    const presentation = parsePptx(bytes, file.name), root = document.createElement("div"); root.className = "ufp-presentation";
    const nav = document.createElement("div"); nav.className = "ufp-slide-nav";
    const previous = document.createElement("button"), next = document.createElement("button"), counter = document.createElement("span");
    previous.type = next.type = "button"; previous.textContent = "Previous slide"; next.textContent = "Next slide";
    const slide = document.createElement("article"); slide.className = "ufp-slide";
    const imageUrls: string[] = [];
    let current = 0, zoom = 1;
    const show = () => {
      for (const url of imageUrls.splice(0)) URL.revokeObjectURL(url);
      slide.replaceChildren(); const data = presentation.slides[current]!;
      slide.style.aspectRatio = `${data.width ?? 16} / ${data.height ?? 9}`;
      slide.style.background = data.background ? pptFillCss(data.background) : "#ffffff";
      if (data.elements?.length) data.elements.forEach(element => slide.append(renderPptxElement(element, data, imageUrls)));
      else data.paragraphs.forEach((paragraph, index) => { const p = document.createElement(index === 0 ? "h1" : "p"); paragraph.runs.forEach(run => p.append(renderRun(run))); slide.append(p); });
      counter.textContent = `${current + 1} / ${presentation.slides.length}`;
      previous.disabled = current === 0; next.disabled = current === presentation.slides.length - 1;
    };
    previous.onclick = () => { if (current > 0) current--, show(); }; next.onclick = () => { if (current + 1 < presentation.slides.length) current++, show(); };
    nav.append(previous, counter, next); root.append(nav, slide); show();
    const keyboard = (event: KeyboardEvent) => { if (event.key === "ArrowLeft") previous.click(); if (event.key === "ArrowRight") next.click(); };
    root.tabIndex = 0; root.addEventListener("keydown", keyboard);
    return { element: root, setZoom(value) { zoom = Math.max(0.5, Math.min(2, value)); slide.style.transform = `scale(${zoom})`; }, fitToWidth() { zoom = 1; slide.style.transform = "scale(1)"; }, destroy() { root.removeEventListener("keydown", keyboard); for (const url of imageUrls) URL.revokeObjectURL(url); imageUrls.length = 0; } };
  }
}
