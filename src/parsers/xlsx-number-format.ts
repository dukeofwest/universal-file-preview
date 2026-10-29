/** Excel number-format engine (ECMA-376 Part 1, 18.8.31) with locale-aware separators and names, as Excel uses the viewer's regional settings. */
export type FormattedValue = { text: string; color?: string; repeat?: { index: number; char: string } };
export type NumberFormatOptions = { locale?: string; date1904?: boolean };

type Token =
  | { t: "lit"; v: string } | { t: "digit"; v: "0" | "#" | "?" } | { t: "dec" } | { t: "group" } | { t: "pct" } | { t: "exp"; sign: boolean }
  | { t: "slash" } | { t: "text" } | { t: "general" } | { t: "repeat"; v: string }
  | { t: "date"; v: string } | { t: "elapsed"; v: string } | { t: "ampm"; v: string } | { t: "subsec"; v: number };
type Section = { tokens: Token[]; color?: string; condition?: { op: string; value: number }; locale?: string; isDate: boolean; hasText: boolean; isGeneral: boolean };

const builtinFormats: Record<number, string> = {
  0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 5: '"$"#,##0_);("$"#,##0)', 6: '"$"#,##0_);[Red]("$"#,##0)', 7: '"$"#,##0.00_);("$"#,##0.00)', 8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??", 14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM",
  20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)', 43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', 45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@",
};
const colorNames: Record<string, string> = { black: "#000000", blue: "#0000ff", cyan: "#00ffff", green: "#00ff00", magenta: "#ff00ff", red: "#ff0000", white: "#ffffff", yellow: "#ffff00" };
/** Excel palette for [ColorN] (and indexed colors 8..63). */
export const indexedColors = ["#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff", "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#c0c0c0", "#808080", "#9999ff", "#993366", "#ffffcc", "#ccffff", "#660066", "#ff8080", "#0066cc", "#ccccff", "#000080", "#ff00ff", "#ffff00", "#00ffff", "#800080", "#800000", "#008080", "#0000ff", "#00ccff", "#ccffff", "#ccffcc", "#ffff99", "#99ccff", "#ff99cc", "#cc99ff", "#ffcc99", "#3366ff", "#33cccc", "#99cc00", "#ffcc00", "#ff9900", "#ff6600", "#666699", "#969696", "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333"];
const lcidLocales: Record<number, string> = { 0x407: "de-DE", 0x807: "de-CH", 0xc07: "de-AT", 0x409: "en-US", 0x809: "en-GB", 0x40c: "fr-FR", 0x410: "it-IT", 0x40a: "es-ES", 0xc0a: "es-ES", 0x413: "nl-NL", 0x416: "pt-BR", 0x816: "pt-PT", 0x415: "pl-PL", 0x419: "ru-RU", 0x411: "ja-JP", 0x804: "zh-CN", 0x412: "ko-KR", 0x41d: "sv-SE", 0x406: "da-DK", 0x414: "nb-NO", 0x40b: "fi-FI", 0x405: "cs-CZ", 0x41f: "tr-TR" };

export function builtinNumberFormat(id: number, locale = "en-US"): string | undefined {
  const language = locale.slice(0, 2).toLowerCase();
  if (id === 14) return shortDatePattern(locale);
  if (id === 22) return `${shortDatePattern(locale)} ${locale === "en-US" ? "h:mm" : "hh:mm"}`;
  if (language === "de" && id >= 15 && id <= 17) return ["dd. mmm yy", "dd. mmm", "mmm yy"][id - 15];
  return builtinFormats[id];
}

const sectionCache = new Map<string, Section[]>();
function parseFormat(format: string): Section[] {
  const cached = sectionCache.get(format);
  if (cached) return cached;
  const sections: Section[] = [];
  let current: Section = { tokens: [], isDate: false, hasText: false, isGeneral: false };
  const lower = format.toLowerCase();
  for (let i = 0; i < format.length; i++) {
    const char = format[i]!, rest = lower.slice(i);
    if (char === ";") { sections.push(current); current = { tokens: [], isDate: false, hasText: false, isGeneral: false }; continue; }
    if (char === '"') { const end = format.indexOf('"', i + 1); current.tokens.push({ t: "lit", v: format.slice(i + 1, end < 0 ? format.length : end) }); i = end < 0 ? format.length : end; continue; }
    if (char === "\\") { current.tokens.push({ t: "lit", v: format[i + 1] ?? "" }); i++; continue; }
    if (char === "_") { current.tokens.push({ t: "lit", v: " " }); i++; continue; }
    if (char === "*") { current.tokens.push({ t: "repeat", v: format[i + 1] ?? " " }); i++; continue; }
    if (char === "[") {
      const end = format.indexOf("]", i), body = format.slice(i + 1, end < 0 ? format.length : end), lowerBody = body.toLowerCase();
      i = end < 0 ? format.length : end;
      if (/^(h+|m+|s+)$/.test(lowerBody)) { current.tokens.push({ t: "elapsed", v: lowerBody }); current.isDate = true; }
      else if (colorNames[lowerBody]) current.color = colorNames[lowerBody];
      else if (/^color\s*\d+$/.test(lowerBody)) { const indexed = indexedColors[Number(lowerBody.replace(/\D/g, "")) - 1]; if (indexed) current.color = indexed; }
      else if (/^(<=|>=|<>|<|>|=)/.test(body)) { const match = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:e[+-]?\d+)?)/i.exec(body); if (match) current.condition = { op: match[1]!, value: Number(match[2]) }; }
      else if (body.startsWith("$")) {
        const dash = body.indexOf("-", 1), symbol = body.slice(1, dash < 0 ? body.length : dash), lcid = dash < 0 ? NaN : Number.parseInt(body.slice(dash + 1), 16);
        if (symbol) current.tokens.push({ t: "lit", v: symbol });
        if ((lcid & 0xffff) === 0xf800) { current.tokens.push({ t: "date", v: "systemlongdate" }); current.isDate = true; }
        else if ((lcid & 0xffff) === 0xf400) { current.tokens.push({ t: "date", v: "systemtime" }); current.isDate = true; }
        else if (Number.isFinite(lcid) && lcidLocales[lcid & 0xffff]) current.locale = lcidLocales[lcid & 0xffff]!;
      }
      continue;
    }
    if (rest.startsWith("general")) { current.tokens.push({ t: "general" }); current.isGeneral = true; i += 6; continue; }
    if (rest.startsWith("am/pm")) { current.tokens.push({ t: "ampm", v: format.slice(i, i + 5) }); current.isDate = true; i += 4; continue; }
    if (rest.startsWith("a/p")) { current.tokens.push({ t: "ampm", v: format.slice(i, i + 3) }); current.isDate = true; i += 2; continue; }
    if (char === "@") { current.tokens.push({ t: "text" }); current.hasText = true; continue; }
    if (char === "0" || char === "#" || char === "?") {
      const previous = current.tokens.at(-1);
      if (char === "0" && previous?.t === "subsec") { previous.v++; continue; }
      current.tokens.push({ t: "digit", v: char }); continue;
    }
    if (char === ".") {
      const previous = current.tokens.at(-1);
      if (current.isDate && /^\.0/.test(rest) && previous && (previous.t === "date" && /^s/.test(previous.v) || previous.t === "elapsed" && previous.v.startsWith("s"))) { current.tokens.push({ t: "subsec", v: 0 }); continue; }
      current.tokens.push({ t: "dec" }); continue;
    }
    if (char === ",") { current.tokens.push({ t: "group" }); continue; }
    if (char === "%") { current.tokens.push({ t: "pct" }); continue; }
    if ((char === "E" || char === "e") && (format[i + 1] === "+" || format[i + 1] === "-") && !current.isDate) { current.tokens.push({ t: "exp", sign: format[i + 1] === "+" }); i++; continue; }
    if (char === "/") { current.tokens.push({ t: "slash" }); continue; }
    const dateMatch = /^(y+|m+|d+|h+|s+|e+|b+|g+)/i.exec(format.slice(i));
    if (dateMatch && !/^(e|b|g)/i.test(dateMatch[0]) || dateMatch && /^e+$/i.test(dateMatch[0]) && current.isDate) {
      current.tokens.push({ t: "date", v: dateMatch[0].toLowerCase() }); current.isDate = true; i += dateMatch[0].length - 1; continue;
    }
    current.tokens.push({ t: "lit", v: char });
  }
  sections.push(current);
  for (const section of sections) {
    if (!section.isDate) continue;
    // Distinguish minutes from months: "m" directly after hours or directly before seconds means minutes.
    const dateTokens = section.tokens.filter((token): token is Extract<Token, { t: "date" | "elapsed" }> => token.t === "date" || token.t === "elapsed");
    dateTokens.forEach((token, index) => {
      if (token.t !== "date" || !/^m{1,2}$/.test(token.v)) return;
      const previous = dateTokens[index - 1], next = dateTokens[index + 1];
      if (previous && /^h/.test(previous.v) || next && /^s/.test(next.v)) token.v = token.v === "m" ? "n" : "nn";
    });
    // Digit placeholders inside date sections are literal digits (e.g. "yyyy0").
    const literal: Partial<Record<Token["t"], string>> = { dec: ".", slash: "/", group: ",", pct: "%" };
    section.tokens = section.tokens.map(token => token.t === "digit" ? { t: "lit", v: token.v } : literal[token.t] !== undefined ? { t: "lit", v: literal[token.t]! } : token);
  }
  sectionCache.set(format, sections);
  return sections;
}

type LocaleInfo = { decimal: string; group: string; months: string[]; monthsShort: string[]; days: string[]; daysShort: string[]; booleans: [string, string] };
const localeCache = new Map<string, LocaleInfo>();
function localeInfo(locale: string): LocaleInfo {
  const cached = localeCache.get(locale);
  if (cached) return cached;
  let decimal = ".", group = ",";
  try {
    const parts = new Intl.NumberFormat(locale).formatToParts(12345.6);
    decimal = parts.find(part => part.type === "decimal")?.value ?? decimal; group = parts.find(part => part.type === "group")?.value ?? group;
    if (group === " " || group === " ") group = " ";
  } catch { /* unknown locale: keep invariant separators */ }
  const language = locale.slice(0, 2).toLowerCase();
  const names = (options: Intl.DateTimeFormatOptions, count: number, at: (index: number) => Date) => {
    try { const formatter = new Intl.DateTimeFormat(locale, { ...options, timeZone: "UTC" }); return Array.from({ length: count }, (_, index) => formatter.format(at(index))); }
    catch { return []; }
  };
  const months = names({ month: "long" }, 12, index => new Date(Date.UTC(2021, index, 15)));
  const days = names({ weekday: "long" }, 7, index => new Date(Date.UTC(2021, 0, 3 + index)));
  const monthsShort = language === "de" ? ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
    : language === "en" ? months.map(month => month.slice(0, 3)) : names({ month: "short" }, 12, index => new Date(Date.UTC(2021, index, 15))).map(name => name.replace(/\.$/, ""));
  const daysShort = language === "en" ? days.map(day => day.slice(0, 3)) : names({ weekday: "short" }, 7, index => new Date(Date.UTC(2021, 0, 3 + index))).map(name => name.replace(/\.$/, ""));
  const booleans: Record<string, [string, string]> = { de: ["WAHR", "FALSCH"], fr: ["VRAI", "FAUX"], es: ["VERDADERO", "FALSO"], it: ["VERO", "FALSO"], nl: ["WAAR", "ONWAAR"], pt: ["VERDADEIRO", "FALSO"] };
  const info: LocaleInfo = { decimal, group, months: months.length ? months : englishMonths, monthsShort: monthsShort.length ? monthsShort : englishMonths.map(month => month.slice(0, 3)), days: days.length ? days : englishDays, daysShort: daysShort.length ? daysShort : englishDays.map(day => day.slice(0, 3)), booleans: booleans[language] ?? ["TRUE", "FALSE"] };
  localeCache.set(locale, info);
  return info;
}
const englishMonths = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const englishDays = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

function shortDatePattern(locale: string): string {
  try {
    const parts = new Intl.DateTimeFormat(locale, { dateStyle: "short", timeZone: "UTC" }).formatToParts(new Date(Date.UTC(2021, 0, 5)));
    return parts.map(part => part.type === "day" ? (part.value.length === 2 ? "dd" : "d") : part.type === "month" ? (part.value.length === 2 ? "mm" : "m") : part.type === "year" ? "yyyy" : part.type === "literal" ? part.value.replace(/[a-z]/gi, char => `\\${char}`) : "").join("");
  } catch { return "m/d/yyyy"; }
}

export function booleanText(value: boolean, locale = "en-US"): string { return localeInfo(locale).booleans[value ? 0 : 1]; }
export function isDateFormat(format: string): boolean { return parseFormat(format).some(section => section.isDate); }

/** Excel "General": up to 11 characters, switching to scientific notation for very large/small magnitudes. */
export function formatGeneral(value: number, locale = "en-US", maxLength = 11): string {
  const decimal = localeInfo(locale).decimal;
  if (!Number.isFinite(value)) return "#NUM!";
  if (value === 0) return "0";
  const abs = Math.abs(value), sign = value < 0 ? "-" : "";
  const budget = Math.max(1, maxLength - sign.length);
  const exponent = Math.floor(Math.log10(abs));
  const fixed = (): string | undefined => {
    const intDigits = abs >= 1 ? exponent + 1 : 1;
    if (intDigits > budget) return undefined;
    const decimals = Math.max(0, Math.min(budget - intDigits - 1, 20));
    let text = abs.toFixed(decimals);
    if (text.includes(".")) text = text.replace(/0+$/, "").replace(/\.$/, "");
    if (text.replace(".", "").length > budget + 1 || text === "0" && abs > 0) return undefined;
    return text;
  };
  const scientific = (): string => {
    let mantissaExponent = exponent;
    const expText = (e: number) => `E${e < 0 ? "-" : "+"}${String(Math.abs(e)).padStart(2, "0")}`;
    const decimals = Math.max(0, budget - expText(mantissaExponent).length - 2);
    let mantissa = (abs / 10 ** mantissaExponent).toFixed(decimals);
    if (Number(mantissa) >= 10) { mantissaExponent++; mantissa = (abs / 10 ** mantissaExponent).toFixed(decimals); }
    if (mantissa.includes(".")) mantissa = mantissa.replace(/0+$/, "").replace(/\.$/, "");
    return mantissa + expText(mantissaExponent);
  };
  const text = abs >= 1e-9 && abs < 1e11 ? fixed() ?? scientific() : scientific();
  return sign + text.replace(".", decimal);
}

export function formatValue(value: number | string | boolean, format: string | undefined, options: NumberFormatOptions = {}): FormattedValue {
  const locale = options.locale ?? "en-US";
  if (typeof value === "boolean") return { text: booleanText(value, locale) };
  const sections = parseFormat(format && format.trim() ? format : "General");
  if (typeof value === "string") {
    const textSection = sections.length >= 4 ? sections[3] : sections.find(section => section.hasText);
    if (!textSection) return { text: value };
    return render(textSection, 0, value, locale, options.date1904 ?? false);
  }
  if (!Number.isFinite(value)) return { text: "#NUM!" };
  const { section, absolute } = chooseSection(sections, value);
  if (!section) return { text: formatGeneral(value, locale) };
  return render(section, absolute ? Math.abs(value) : value, undefined, locale, options.date1904 ?? false);
}

function chooseSection(sections: Section[], value: number): { section: Section | undefined; absolute: boolean } {
  const numeric = sections.slice(0, 3);
  const test = (condition: Section["condition"]) => {
    if (!condition) return false;
    const { op, value: limit } = condition;
    return op === "<" ? value < limit : op === "<=" ? value <= limit : op === ">" ? value > limit : op === ">=" ? value >= limit : op === "=" ? value === limit : value !== limit;
  };
  // A section whose condition only admits negative numbers shows them without a sign, like the default negative section.
  const negativeOnly = (condition: Section["condition"]) => !!condition && (condition.op === "<" && condition.value <= 0 || condition.op === "<=" && condition.value < 0);
  if (numeric.some(section => section.condition)) {
    const [first, second, third] = numeric;
    if (first?.condition && test(first.condition)) return { section: first, absolute: negativeOnly(first.condition) };
    if (second?.condition) return test(second.condition) ? { section: second, absolute: negativeOnly(second.condition) } : { section: third ?? first, absolute: false };
    return { section: second ?? first, absolute: false };
  }
  if (numeric.length === 1 || value > 0) return { section: numeric[0], absolute: false };
  if (value < 0) return { section: numeric[1] ?? numeric[0], absolute: numeric.length > 1 };
  return { section: numeric[2] ?? numeric[0], absolute: false };
}

function render(section: Section, value: number, text: string | undefined, locale: string, date1904: boolean): FormattedValue {
  const info = localeInfo(section.locale ?? locale), separators = localeInfo(locale);
  const out: string[] = [];
  let repeat: FormattedValue["repeat"];
  const push = (piece: string) => { out.push(piece); };
  const markRepeat = (char: string) => { if (!repeat) repeat = { index: out.join("").length, char }; };
  const finish = (): FormattedValue => {
    const result: FormattedValue = { text: out.join("") };
    if (section.color) result.color = section.color;
    if (repeat) result.repeat = repeat;
    return result;
  };
  if (text !== undefined) {
    for (const token of section.tokens) {
      if (token.t === "text" || token.t === "general") push(text);
      else if (token.t === "lit") push(token.v);
      else if (token.t === "repeat") markRepeat(token.v);
    }
    return finish();
  }
  if (section.isDate) { renderDate(section, value, info, separators, date1904, push, markRepeat); return finish(); }
  if (!section.tokens.some(token => token.t === "digit")) {
    // "General", "@" or literal-only sections: the number is shown in General form where the placeholder sits.
    const hasSlot = section.tokens.some(token => token.t === "general" || token.t === "text");
    for (const token of section.tokens) {
      if (token.t === "general" || token.t === "text") push(formatGeneral(value, locale));
      else if (token.t === "lit") push(token.v);
      else if (token.t === "pct") push("%");
      else if (token.t === "repeat") markRepeat(token.v);
    }
    if (!hasSlot && !section.tokens.some(token => token.t === "lit")) push(formatGeneral(value, locale));
    return finish();
  }
  renderNumber(section.tokens, value, separators, push, markRepeat);
  return finish();
}

function renderNumber(tokens: Token[], value: number, info: LocaleInfo, push: (piece: string) => void, markRepeat: (char: string) => void): void {
  let number = Math.abs(value);
  const digitIndexes = tokens.map((token, index) => token.t === "digit" ? index : -1).filter(index => index >= 0);
  const firstDigit = digitIndexes[0] ?? -1, lastDigit = digitIndexes.at(-1) ?? -1;
  number *= 100 ** tokens.filter(token => token.t === "pct").length;
  const expIndex = tokens.findIndex(token => token.t === "exp");
  const slashIndex = tokens.findIndex((token, index) => token.t === "slash" && index > firstDigit && index < lastDigit);
  const decIndex = slashIndex >= 0 ? -1 : tokens.findIndex((token, index) => token.t === "dec" && (expIndex < 0 || index < expIndex));
  const intEnd = decIndex >= 0 ? decIndex : expIndex >= 0 ? expIndex : tokens.length;
  // Commas right after the last integer placeholder scale by 1000 each; commas between placeholders turn on grouping.
  const lastIntDigit = Math.max(-1, ...digitIndexes.filter(index => index < intEnd));
  for (let index = lastIntDigit + 1; index < intEnd && tokens[index]!.t === "group"; index++) number /= 1000;
  const grouping = tokens.some((token, index) => token.t === "group" && index > firstDigit && index < lastIntDigit);
  if (slashIndex >= 0) {
    const pieces = renderFraction(tokens, slashIndex, number);
    if (value < 0 && /[1-9]/.test(Object.values(pieces).join(""))) push("-");
    tokens.forEach((token, index) => {
      if (pieces[index] !== undefined) push(pieces[index]!);
      else if (token.t === "lit") push(token.v);
      else if (token.t === "pct") push("%");
      else if (token.t === "repeat") markRepeat(token.v);
    });
    return;
  }
  const fracTokens = decIndex >= 0 ? tokens.slice(decIndex + 1, expIndex >= 0 ? expIndex : tokens.length) : [];
  const fracPlaceholders = fracTokens.filter((token): token is Extract<Token, { t: "digit" }> => token.t === "digit");
  let exponent = 0;
  if (expIndex >= 0 && number !== 0) {
    const intPlaceholders = digitIndexes.filter(index => index < intEnd).length;
    exponent = Math.floor(Math.log10(number));
    // "##0.0E+0" style formats keep the exponent a multiple of the integer placeholder count (engineering notation).
    if (intPlaceholders > 1 && tokens.slice(0, intEnd).some(token => token.t === "digit" && token.v === "#")) exponent = Math.floor(exponent / intPlaceholders) * intPlaceholders;
    else exponent -= Math.max(0, intPlaceholders - 1);
    number /= 10 ** exponent;
  }
  const rounded = number.toFixed(Math.min(fracPlaceholders.length, 20));
  let [intDigits = "0", fracDigits = ""] = rounded.split(".");
  if (value < 0 && /[1-9]/.test(rounded)) push("-");
  if (intDigits === "0") intDigits = "";
  // Integer placeholders are filled right-to-left; surplus digits go to the leftmost placeholder.
  const intPieces = new Map<number, string>();
  const intPlaceholderIndexes = digitIndexes.filter(index => index < intEnd);
  let remaining = intDigits, position = 0;
  for (let p = intPlaceholderIndexes.length - 1; p >= 0; p--) {
    const index = intPlaceholderIndexes[p]!, placeholder = (tokens[index] as Extract<Token, { t: "digit" }>).v;
    const take = p === 0 ? remaining : remaining.slice(-1);
    remaining = p === 0 ? "" : remaining.slice(0, -1);
    const piece = take || (placeholder === "0" ? "0" : placeholder === "?" ? " " : "");
    let grouped = "";
    for (let c = piece.length - 1; c >= 0; c--) {
      const char = piece[c]!;
      grouped = char + (grouping && position > 0 && position % 3 === 0 && char !== " " ? info.group : "") + grouped;
      position++;
    }
    intPieces.set(index, grouped);
  }
  const fracPieces = fracPlaceholders.map((_, index) => fracDigits[index] ?? "0");
  for (let p = fracPieces.length - 1; p >= 0 && fracPieces[p] === "0"; p--) {
    const kind = fracPlaceholders[p]!.v;
    if (kind === "#") fracPieces[p] = ""; else if (kind === "?") fracPieces[p] = " "; else break;
  }
  let fracCursor = 0;
  tokens.forEach((token, index) => {
    if (index === decIndex) { push(info.decimal); return; }
    if (index === expIndex) {
      const expDigits = tokens.slice(expIndex + 1).filter(next => next.t === "digit").length;
      push(`E${exponent < 0 ? "-" : (token as Extract<Token, { t: "exp" }>).sign ? "+" : ""}${String(Math.abs(exponent)).padStart(expDigits, "0")}`);
      return;
    }
    if (token.t === "digit") { if (index < intEnd) push(intPieces.get(index) ?? ""); else if (expIndex < 0 || index < expIndex) push(fracPieces[fracCursor++] ?? ""); return; }
    if (token.t === "lit") push(token.v);
    else if (token.t === "pct") push("%");
    else if (token.t === "repeat") markRepeat(token.v);
  });
}

/** Returns the text for each placeholder index of a fraction format ("# ?/?", "# ??/16", "?/?"). */
function renderFraction(tokens: Token[], slashIndex: number, number: number): Record<number, string> {
  let numeratorStart = slashIndex - 1;
  while (numeratorStart > 0 && tokens[numeratorStart - 1]!.t === "digit") numeratorStart--;
  const integerIndexes = tokens.slice(0, numeratorStart).map((token, index) => token.t === "digit" ? index : -1).filter(index => index >= 0);
  const numeratorIndexes = tokens.slice(numeratorStart, slashIndex).map((token, index) => token.t === "digit" ? numeratorStart + index : -1).filter(index => index >= 0);
  const denominatorIndexes = tokens.map((token, index) => index > slashIndex && (token.t === "digit" || token.t === "lit" && /^\d$/.test(token.v)) ? index : -1).filter(index => index >= 0);
  const fixed = denominatorIndexes.length && denominatorIndexes.every(index => tokens[index]!.t === "lit") ? Number(denominatorIndexes.map(index => (tokens[index] as { v: string }).v).join("")) : undefined;
  const whole = integerIndexes.length ? Math.floor(number) : 0, fraction = number - whole;
  let numerator = 0, denominator = 1;
  if (fixed) { denominator = fixed; numerator = Math.round(fraction * fixed); }
  else {
    const maxDenominator = 10 ** Math.max(1, denominatorIndexes.length) - 1;
    let bestError = Infinity;
    for (let candidate = 1; candidate <= maxDenominator; candidate++) {
      const guess = Math.round(fraction * candidate), error = Math.abs(fraction - guess / candidate);
      if (error < bestError - 1e-12) { bestError = error; numerator = guess; denominator = candidate; }
    }
  }
  let integer = whole;
  if (integerIndexes.length && numerator === denominator) { integer++; numerator = 0; }
  const fill = (index: number) => { const kind = (tokens[index] as { v: string }).v; return kind === "0" ? "0" : kind === "?" ? " " : ""; };
  const place = (digits: string, indexes: number[], alignRight: boolean, pieces: Record<number, string>) => {
    indexes.forEach(index => { pieces[index] = fill(index); });
    if (!indexes.length) return;
    if (alignRight) {
      let rest = digits;
      for (let p = indexes.length - 1; p >= 0; p--) { const take = p === 0 ? rest : rest.slice(-1); rest = p === 0 ? "" : rest.slice(0, -1); if (take) pieces[indexes[p]!] = take; }
    } else {
      let rest = digits;
      for (let p = 0; p < indexes.length; p++) { const take = p === indexes.length - 1 ? rest : rest.slice(0, 1); rest = p === indexes.length - 1 ? "" : rest.slice(1); if (take) pieces[indexes[p]!] = take; }
    }
  };
  const pieces: Record<number, string> = {};
  const showFraction = numerator !== 0 || !integerIndexes.length;
  place(integer === 0 && showFraction ? "" : String(integer), integerIndexes, true, pieces);
  if (showFraction) {
    place(String(numerator), numeratorIndexes, true, pieces);
    pieces[slashIndex] = "/";
    if (!fixed) place(String(denominator), denominatorIndexes, false, pieces);
  } else {
    for (const index of [...numeratorIndexes, slashIndex, ...(fixed ? [] : denominatorIndexes)]) pieces[index] = " ";
    if (fixed) for (const index of denominatorIndexes) pieces[index] = " ";
  }
  return pieces;
}

function renderDate(section: Section, serial: number, info: LocaleInfo, separators: LocaleInfo, date1904: boolean, push: (piece: string) => void, markRepeat: (char: string) => void): void {
  const subsecondDigits = Math.max(0, ...section.tokens.map(token => token.t === "subsec" ? token.v : 0));
  const scale = 10 ** subsecondDigits;
  let days = Math.floor(serial), seconds = Math.round((serial - days) * 86400 * scale) / scale;
  if (seconds >= 86400) { seconds -= 86400; days++; }
  const date = serialToDate(days, date1904);
  const twelveHour = section.tokens.some(token => token.t === "ampm");
  const hours = Math.floor(seconds / 3600), minutes = Math.floor(seconds % 3600 / 60), wholeSeconds = Math.floor(seconds % 60);
  const totalSeconds = serial * 86400;
  for (const token of section.tokens) {
    if (token.t === "lit") { push(token.v); continue; }
    if (token.t === "repeat") { markRepeat(token.v); continue; }
    if (token.t === "subsec") { push(separators.decimal + String(Math.round((seconds % 1) * scale)).padStart(token.v, "0").slice(0, token.v)); continue; }
    if (token.t === "ampm") { const pm = hours >= 12; push(token.v.length === 3 ? (pm ? token.v[2]! : token.v[0]!) : (pm ? (token.v.startsWith("a") ? "pm" : "PM") : (token.v.startsWith("a") ? "am" : "AM"))); continue; }
    if (token.t === "elapsed") {
      const width = token.v.length;
      const amount = token.v[0] === "h" ? Math.floor(totalSeconds / 3600 + 1e-9) : token.v[0] === "m" ? Math.floor(totalSeconds / 60 + 1e-9) : Math.round(totalSeconds);
      push(String(amount).padStart(width, "0")); continue;
    }
    if (token.t !== "date") continue;
    const v = token.v;
    if (v === "systemlongdate") { try { push(new Intl.DateTimeFormat(section.locale ?? undefined, { dateStyle: "full", timeZone: "UTC" }).format(date.value)); } catch { push(date.value.toISOString().slice(0, 10)); } continue; }
    if (v === "systemtime") { push(`${hours}:${String(minutes).padStart(2, "0")}:${String(wholeSeconds).padStart(2, "0")}`); continue; }
    if (v[0] === "y" || v[0] === "e") push(v.length <= 2 && v[0] === "y" ? String(date.year % 100).padStart(2, "0") : String(date.year));
    else if (v[0] === "m") push(v.length === 1 ? String(date.month + 1) : v.length === 2 ? String(date.month + 1).padStart(2, "0") : v.length === 3 ? info.monthsShort[date.month]! : v.length === 5 ? info.months[date.month]!.slice(0, 1) : info.months[date.month]!);
    else if (v[0] === "d") push(v.length === 1 ? String(date.day) : v.length === 2 ? String(date.day).padStart(2, "0") : v.length === 3 ? info.daysShort[date.weekday]! : info.days[date.weekday]!);
    else if (v[0] === "h") { const hour = twelveHour ? (hours % 12 || 12) : hours; push(v.length >= 2 ? String(hour).padStart(2, "0") : String(hour)); }
    else if (v[0] === "n") push(v.length >= 2 ? String(minutes).padStart(2, "0") : String(minutes));
    else if (v[0] === "s") push(v.length >= 2 ? String(wholeSeconds).padStart(2, "0") : String(wholeSeconds));
  }
}

/** Converts an Excel serial day number to a calendar date, honoring the 1900 leap-year bug and the 1904 date system. */
export function serialToDate(days: number, date1904 = false): { year: number; month: number; day: number; weekday: number; value: Date } {
  if (!date1904 && days === 60) return { year: 1900, month: 1, day: 29, weekday: 3, value: new Date(Date.UTC(1900, 1, 28)) };
  if (!date1904 && days === 0) return { year: 1900, month: 0, day: 0, weekday: 6, value: new Date(Date.UTC(1899, 11, 31)) };
  const base = date1904 ? Date.UTC(1904, 0, 1) : days < 60 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
  const value = new Date(base + days * 86_400_000);
  return { year: value.getUTCFullYear(), month: value.getUTCMonth(), day: value.getUTCDate(), weekday: value.getUTCDay(), value };
}

export function dateToSerial(year: number, month: number, day: number, date1904 = false): number {
  const value = Date.UTC(year, month, day);
  if (date1904) return Math.round((value - Date.UTC(1904, 0, 1)) / 86_400_000);
  const serial = Math.round((value - Date.UTC(1899, 11, 30)) / 86_400_000);
  return serial < 61 ? serial - 1 : serial;
}
