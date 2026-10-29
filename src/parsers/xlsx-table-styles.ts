/**
 * Built-in Excel table styles (TableStyleLight1..21, Medium1..28, Dark1..11), captured from Excel's own definitions.
 * Encoding per element "key:tokens", elements joined by "|". Keys: w whole table, h header row, t total row, fc/lc first/last column,
 * r1/r2 row stripes, c1/c2 column stripes. Tokens: F<color> fill, C<color> font color, B bold, I italic, S<n> stripe size,
 * <side><weight><line?><color> border with side l/t/b/r/v(inside vertical)/x(inside horizontal), weight h/n/m/k (hair/thin/medium/thick),
 * line d/-/./y/z/s (double/dashed/dotted/dashDot/dashDotDot/slantDashDot). Colors: theme index[/tint], #rrggbb, or A = the style's accent.
 */
const builtinTableStyles: Record<string, string> = {
  "Light1": "w:C1 tn1 bn1|h:C1 B bn1|t:C1 B tn1|fc:C1 B|lc:C1 B|r1:F0/-0.15|c1:F0/-0.15",
  "Light2-7": "w:CA/-0.25 tnA bnA|h:CA/-0.25 B bnA|t:CA/-0.25 B tnA|fc:CA/-0.25 B|lc:CA/-0.25 B|r1:FA/0.8|c1:FA/0.8",
  "Light8": "w:C1 ln1 tn1 bn1 rn1|h:F1 C0 B|t:C1 B tkd1|fc:C1 B|lc:C1 B|r1:tn1|r2:tn1|c1:ln1|c2:ln1",
  "Light9-14": "w:C1 lnA tnA bnA rnA|h:FA C0 B|t:C1 B tkdA|fc:C1 B|lc:C1 B|r1:tnA|r2:tnA|c1:lnA|c2:lnA",
  "Light15": "w:C1 ln1 tn1 bn1 rn1 vn1 xn1|h:C1 B bm1|t:C1 B tkd1|fc:C1 B|lc:C1 B|r1:F0/-0.15|c1:F0/-0.15",
  "Light16-21": "w:C1 lnA tnA bnA rnA vnA xnA|h:C1 B bmA|t:C1 B tkdA|fc:C1 B|lc:C1 B|r1:FA/0.8|c1:FA/0.8",
  "Medium1": "w:C1 ln1 tn1 bn1 rn1 xn1|h:F1 C0 B|t:C1 B tkd1|fc:C1 B|lc:C1 B|r1:F0/-0.15|c1:F0/-0.15",
  "Medium2-7": "w:C1 lnA/0.4 tnA/0.4 bnA/0.4 rnA/0.4 xnA/0.4|h:FA C0 B|t:C1 B tkdA|fc:C1 B|lc:C1 B|r1:FA/0.8|c1:FA/0.8",
  "Medium8": "w:F0/-0.15 C1 vn0 xn0|h:F1 C0 B bk0|t:F1 C0 B tk0|fc:F1 C0 B|lc:F1 C0 B|r1:F0/-0.35|c1:F0/-0.35",
  "Medium9-14": "w:FA/0.8 C1 vn0 xn0|h:FA C0 B bk0|t:FA C0 B tk0|fc:FA C0 B|lc:FA C0 B|r1:FA/0.6|c1:FA/0.6",
  "Medium15": "w:C1 ln1 tm1 bm1 rn1 vn1 xn1|h:F1 C0 B bm1|t:tkd1|fc:F1 C0 B|lc:F1 C0 B|r1:F0/-0.15|c1:F0/-0.15",
  "Medium16-21": "w:C1 tm1 bm1|h:FA C0 B bm1|t:tkd1|fc:FA C0 B|lc:FA C0 B|r1:F0/-0.15|c1:F0/-0.15",
  "Medium22": "w:F0/-0.15 C1 ln1 tn1 bn1 rn1 vn1 xn1|h:C1 B|t:C1 B tm1|fc:C1 B|lc:C1 B|r1:F0/-0.35|c1:F0/-0.35",
  "Medium23-28": "w:FA/0.8 C1 lnA/0.4 tnA/0.4 bnA/0.4 rnA/0.4 vnA/0.4 xnA/0.4|h:C1 B|t:C1 B tmA|fc:C1 B|lc:C1 B|r1:FA/0.6|c1:FA/0.6",
  "Dark1": "w:F1/0.45 C0|h:F1 C0 B bm0|t:F1/0.15 C0 B tm0|fc:F1/0.25 C0 B rm0|lc:F1/0.25 C0 B lm0|r1:F1/0.25|c1:F1/0.25",
  "Dark2-7": "w:FA C0|h:F1 C0 B bm0|t:FA/-0.5 C0 B tm0|fc:FA/-0.25 C0 B rm0|lc:FA/-0.25 C0 B lm0|r1:FA/-0.25|c1:FA/-0.25",
  "Dark8": "w:F0/-0.15|h:F1 C0|t:C1 B tkd1|fc:C1 B|lc:C1 B|r1:F0/-0.35|c1:F0/-0.35",
  "Dark9": "w:F4/0.8|h:F5 C0|t:C1 B tkd1|fc:C1 B|lc:C1 B|r1:F4/0.6|c1:F4/0.6",
  "Dark10": "w:F6/0.8|h:F7 C0|t:C1 B tkd1|fc:C1 B|lc:C1 B|r1:F6/0.6|c1:F6/0.6",
  "Dark11": "w:F8/0.8|h:F9 C0|t:C1 B tkd1|fc:C1 B|lc:C1 B|r1:F8/0.6|c1:F8/0.6",
};

export type TableStyleColor = { theme: number; tint: number } | { rgb: string };
export type TableStyleBorder = { style: string; color: TableStyleColor };
export type TableStyleElement = { fill?: TableStyleColor; color?: TableStyleColor; bold?: boolean; italic?: boolean; size?: number; borders?: Partial<Record<"left" | "top" | "bottom" | "right" | "insideV" | "insideH", TableStyleBorder>> };
const elementNames: Record<string, string> = { w: "wholeTable", h: "headerRow", t: "totalRow", fc: "firstColumn", lc: "lastColumn", r1: "firstRowStripe", r2: "secondRowStripe", c1: "firstColumnStripe", c2: "secondColumnStripe" };
const sides: Record<string, "left" | "top" | "bottom" | "right" | "insideV" | "insideH"> = { l: "left", t: "top", b: "bottom", r: "right", v: "insideV", x: "insideH" };

/** Returns the elements of a built-in table style keyed by ECMA element type names (wholeTable, headerRow, ...), or undefined. */
export function builtinTableStyle(name: string): Record<string, TableStyleElement> | undefined {
  const match = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name);
  if (!match) return undefined;
  const family = match[1]!, number = Number(match[2]);
  let encoded: string | undefined, accent = 4;
  for (const [key, value] of Object.entries(builtinTableStyles)) {
    const range = /^(Light|Medium|Dark)(\d+)(?:-(\d+))?$/.exec(key)!;
    const first = Number(range[2]), last = Number(range[3] ?? range[2]);
    if (range[1] === family && number >= first && number <= last) { encoded = value; accent = 4 + number - first; break; }
  }
  if (!encoded) return undefined;
  const color = (text: string): TableStyleColor => {
    if (text.startsWith("#")) return { rgb: text };
    const [theme, tint] = text.split("/");
    return { theme: theme === "A" ? accent : Number(theme), tint: Number(tint ?? 0) };
  };
  const result: Record<string, TableStyleElement> = {};
  for (const part of encoded.split("|")) {
    const [key, tokenText = ""] = part.split(":") as [string, string?];
    const element: TableStyleElement = {};
    for (const token of tokenText.split(" ").filter(Boolean)) {
      if (token[0] === "F") element.fill = color(token.slice(1));
      else if (token[0] === "C") element.color = color(token.slice(1));
      else if (token === "B") element.bold = true;
      else if (token === "I") element.italic = true;
      else if (token[0] === "S") element.size = Number(token.slice(1));
      else {
        const border = /^([ltbrvx])([hnmk])([d\-.yzs]?)(.+)$/.exec(token);
        if (!border) continue;
        const weight = border[2]!, line = border[3]!;
        const style = line === "d" ? "double" : line === "-" ? (weight === "m" ? "mediumDashed" : "dashed") : line === "." ? "dotted" : line === "y" ? (weight === "m" ? "mediumDashDot" : "dashDot") : line === "z" ? (weight === "m" ? "mediumDashDotDot" : "dashDotDot") : line === "s" ? "slantDashDot"
          : weight === "h" ? "hair" : weight === "m" ? "medium" : weight === "k" ? "thick" : "thin";
        (element.borders ??= {})[sides[border[1]!]!] = { style, color: color(border[4]!) };
      }
    }
    result[elementNames[key] ?? key] = element;
  }
  return result;
}
