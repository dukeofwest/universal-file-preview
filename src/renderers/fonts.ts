/**
 * Font names come from the file. Only letters, marks, digits, spaces and a few punctuation marks
 * that occur in real font names are accepted, so a name can never leave its quoted CSS string:
 * no quotes, backslashes, semicolons, braces, angle brackets, newlines or other control
 * characters. Any other name falls back to the generic family.
 */
const safeFontName = /^[\p{L}\p{M}\p{N}_.&()+@-][\p{L}\p{M}\p{N} _.&()+@-]{0,127}$/u;

/** CSS font stack for an Office font name, with metric-compatible local fallbacks for fonts that are often missing (cloud fonts). */
export function fontStack(fontFamily: string): string {
  const family = fontFamily.trim().replace(/ {2,}/g, " ");
  if (!safeFontName.test(family)) return "sans-serif";
  const normalized = family.toLowerCase();
  const fallback = /mono|courier|consolas|menlo|terminal/.test(normalized) ? "monospace"
    : /arial|aptos|calibri|franklin|gothic|helvetica|segoe|tahoma|trebuchet|verdana|sans/.test(normalized) ? "sans-serif"
      : /cambria|constantia|garamond|georgia|palatino|baskerville|times|serif|bookman|goudy|old style|caslon|bembo|centaur|californian|perpetua|minion|book antiqua|century|didot|bodoni|sitka|bell mt|high tower|elephant/.test(normalized) ? "serif" : "sans-serif";
  // Old-style faces are often missing locally (e.g. Office cloud fonts); Garamond-like faces keep their metrics closest.
  const oldStyle = /goudy|old style|garamond|caslon|bembo|centaur|californian|perpetua|minion|high tower/.test(normalized) ? `Garamond, "EB Garamond", "Sitka Text", Georgia, ` : "";
  const light = fallback === "sans-serif" && /\b(?:light|thin)\b/.test(normalized) ? `"Segoe UI Light", "Helvetica Neue Light", ` : "";
  // Metric-compatible substitutes keep line breaks and column fits close when the Office font is not installed ("ufp ..." faces are defined in the viewer styles).
  const metric = normalized === "aptos narrow" ? `Calibri, Carlito, ` : normalized === "calibri" ? `Carlito, ` : normalized === "cambria" ? `Caladea, ` : normalized === "arial" ? `"Liberation Sans", Arimo, ` : normalized === "times new roman" ? `"Liberation Serif", Tinos, ` : normalized === "courier new" ? `"Liberation Mono", Cousine, `
    : normalized === "franklin gothic book" ? `"ufp Franklin Gothic Book", "ufp Franklin Gothic Book Alt", `
      : normalized === "gill sans mt" ? `"ufp Gill Sans MT", "ufp Gill Sans MT Alt", ` : "";
  return `"${family}", ${metric}${oldStyle}${light}${fallback}`;
}
