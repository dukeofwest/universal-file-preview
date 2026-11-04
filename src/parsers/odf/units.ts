/**
 * OpenDocument lengths, colours and borders (ODF 1.3 Part 3 §18; implementation specification
 * §4.5). Internal module. Invalid or unsupported values yield undefined, never a guess.
 */

const POINTS_PER_UNIT: Readonly<Record<string, number>> = { pt: 1, pc: 12, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, px: 0.75 };
const lengthPattern = /^([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)(cm|mm|in|pt|pc|px)$/;

/** "2.258cm" | "0.178in" | "12pt" | "10mm" | "1pc" | "16px" in points; undefined for invalid input and percentages. */
export function lengthToPoints(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const match = lengthPattern.exec(value.trim());
  if (!match) return undefined;
  const points = Number(match[1]) * POINTS_PER_UNIT[match[2]!]!;
  return Number.isFinite(points) ? points : undefined;
}

/** "#RRGGBB" in lower case; "transparent" and anything else is undefined. */
export function odfColor(value: string | undefined): string | undefined {
  const text = value?.trim();
  return text && /^#[0-9a-fA-F]{6}$/.test(text) ? text.toLowerCase() : undefined;
}

export type OdfBorder = { widthPt: number; style: string; color: string };
const borderStyles = new Set(["solid", "double", "dashed", "dotted", "groove", "ridge", "inset", "outset", "hidden", "none"]);
const borderWidthKeywords: Readonly<Record<string, number>> = { thin: 0.75, medium: 2.25, thick: 3.75 };

/** "0.06pt solid #000000" (tokens in any order) as a border; undefined for "none", "hidden", a missing style or unparsable tokens. */
export function parseBorder(value: string | undefined): OdfBorder | undefined {
  if (value === undefined) return undefined;
  let widthPt: number | undefined, style: string | undefined, color: string | undefined;
  for (const token of value.trim().split(/\s+/)) {
    const lower = token.toLowerCase();
    if (borderStyles.has(lower)) style = lower;
    else if (Object.hasOwn(borderWidthKeywords, lower)) widthPt = borderWidthKeywords[lower];
    else if (lengthToPoints(token) !== undefined) widthPt = lengthToPoints(token);
    else if (odfColor(token) !== undefined) color = odfColor(token);
    else return undefined;
  }
  if (!style || style === "none" || style === "hidden") return undefined;
  return { widthPt: widthPt ?? 0.75, style, color: color ?? "#000000" };
}
