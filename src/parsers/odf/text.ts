import type { SpreadsheetFont, SpreadsheetRun } from "../../model.js";
import { nsAttr, nsFirst, type NsXmlNode } from "../../xml/ns-tree.js";
import { ODF } from "./namespaces.js";

/**
 * Saved display text of OpenDocument paragraphs (ODF 1.3 Part 3 §6.1; implementation
 * specification §4.5 and §6.3). Internal module.
 *
 * White space follows ODF §6.1.2: literal spaces, tabs and line ends collapse to one space and
 * are dropped at the start and end of a paragraph, while text:s, text:tab and text:line-break
 * produce significant spaces, tabs and line breaks. Literal white space collapses only with
 * adjacent literal white space, so a space on either side of text:s is kept. [Known; confirm
 * with producer fixtures. LibreOffice writes "a <text:s text:c="2"/>b" for three spaces.]
 * Paragraphs are joined with "\n". Text of foreign and unknown elements is included; comments
 * (office:annotation) are not part of the displayed text.
 */

/** Text runs sharing one span style; `style` is the innermost text:style-name, if any. */
type Segment = { text: string; style?: string };

const MAX_SPACE_COUNT = 10_000;

/** Walks the paragraphs of a cell (or comment) into styled segments. */
function paragraphSegments(paragraphs: readonly NsXmlNode[]): Segment[] {
  const segments: Segment[] = [];
  const emit = (text: string, style: string | undefined) => {
    if (!text) return;
    const last = segments[segments.length - 1];
    if (last && last.style === style) last.text += text;
    else segments.push(style === undefined ? { text } : { text, style });
  };
  paragraphs.forEach((paragraph, index) => {
    if (index) emit("\n", undefined);
    // `pending` is a collapsed space waiting for following content (dropped at the paragraph end).
    // It keeps the span style where it was written, so "plain <span>bold</span>" runs as "plain " + "bold".
    let started = false, pending = false, pendingStyle: string | undefined;
    const content = (text: string, style: string | undefined) => {
      if (pending && started) emit(" ", pendingStyle);
      pending = false; started = true;
      emit(text, style);
    };
    const walk = (node: NsXmlNode, style: string | undefined) => {
      for (const child of node.children) {
        if (typeof child === "string") {
          for (const piece of child.split(/([ \t\r\n]+)/)) {
            if (!piece) continue;
            if (/^[ \t\r\n]+$/.test(piece)) { if (!pending) pendingStyle = style; pending = true; }
            else content(piece, style);
          }
          continue;
        }
        if (child.uri === ODF.office && child.local === "annotation") continue;
        if (child.uri === ODF.text) {
          if (child.local === "s") { content(" ".repeat(spaceCount(nsAttr(child, ODF.text, "c"))), style); continue; }
          if (child.local === "tab") { content("\t", style); continue; }
          if (child.local === "line-break") { content("\n", style); continue; }
          if (child.local === "span") { walk(child, nsAttr(child, ODF.text, "style-name") ?? style); continue; }
        }
        walk(child, style);
      }
    };
    walk(paragraph, undefined);
  });
  return segments;
}

function spaceCount(value: string | undefined): number {
  if (value === undefined) return 1;
  const count = Number(value);
  return Number.isSafeInteger(count) && count >= 1 ? Math.min(count, MAX_SPACE_COUNT) : 1;
}

/** The paragraph children of a cell or comment, in document order. */
function paragraphsOf(node: NsXmlNode): NsXmlNode[] {
  return node.children.filter((child): child is NsXmlNode => typeof child !== "string" && child.uri === ODF.text && (child.local === "p" || child.local === "h"));
}

/** Saved display text of a cell; undefined when it has no text:p (an empty cell). */
export function savedCellText(cell: NsXmlNode): string | undefined {
  const paragraphs = paragraphsOf(cell);
  return paragraphs.length ? paragraphSegments(paragraphs).map(segment => segment.text).join("") : undefined;
}

/** Looks up the font of a text-family style; implemented by the style resolver. */
export interface TextFontSource { textRunFont(styleName: string): SpreadsheetFont | undefined; }

/**
 * Rich-text runs, only when the cell text uses more than one span style. The runs' text always
 * concatenates to savedCellText(cell).
 */
export function cellRuns(cell: NsXmlNode, styles: TextFontSource): SpreadsheetRun[] | undefined {
  const segments = paragraphSegments(paragraphsOf(cell));
  if (new Set(segments.map(segment => segment.style)).size < 2) return undefined;
  return segments.map(segment => {
    const font = segment.style === undefined ? undefined : styles.textRunFont(segment.style);
    return font ? { text: segment.text, font } : { text: segment.text };
  });
}

/** The first text:a link in the cell text with an http, https or mailto target. */
export function cellHyperlink(cell: NsXmlNode): string | undefined {
  const find = (node: NsXmlNode): string | undefined => {
    for (const child of node.children) {
      if (typeof child === "string" || (child.uri === ODF.office && child.local === "annotation")) continue;
      if (child.uri === ODF.text && child.local === "a") {
        const target = nsAttr(child, ODF.xlink, "href")?.trim();
        if (target && /^(https?:|mailto:)/i.test(target)) return target;
      }
      const nested = find(child);
      if (nested) return nested;
    }
    return undefined;
  };
  for (const paragraph of paragraphsOf(cell)) { const link = find(paragraph); if (link) return link; }
  return undefined;
}

/** Text of the cell's comment (office:annotation paragraphs; author and date excluded). */
export function annotationText(cell: NsXmlNode): string | undefined {
  const annotation = nsFirst(cell, ODF.office, "annotation");
  if (!annotation) return undefined;
  const paragraphs = paragraphsOf(annotation);
  return paragraphs.length ? paragraphSegments(paragraphs).map(segment => segment.text).join("") : undefined;
}
