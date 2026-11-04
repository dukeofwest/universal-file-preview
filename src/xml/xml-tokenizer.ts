import { PreviewError } from "../errors.js";

/**
 * Secure pull tokenizer for XML 1.0 (implementation specification §4.1, milestone M2).
 *
 * Internal module: not exported from the package entry point (decision D2). It is the lexer
 * for the namespace-aware tree builder and stream reader used by the OpenDocument parsers.
 * The OOXML parsers keep using `parseXml` (xml-parser.ts), which is unchanged.
 *
 * Security rules match `parseXml`: DTDs and entity declarations are rejected, only the five
 * predefined entities and numeric character references are decoded, nothing external is
 * ever resolved, and every resource is bounded.
 *
 * Deliberate differences from `parseXml`, all required by XML 1.0:
 * - CDATA content is literal. `parseXml` entity-decodes it.
 * - Line ends in text and CDATA are normalised to LF (§2.11).
 * - Attribute values are normalised (§3.3.3): literal TAB, CR, LF and CRLF become one space
 *   before entity decoding, so `&#10;` still yields LF.
 * - Raw characters outside the XML Char production, `<` in attribute values, `]]>` in text,
 *   CDATA outside the root element, and attributes not separated by whitespace are rejected.
 * - Element and attribute names follow the full XML Name production, including non-ASCII.
 * - Self-closing elements produce a single `start` token with `selfClosing: true` and no `end`.
 *
 * Kept for parity with `parseXml` although XML 1.0 is stricter: hexadecimal character
 * references accept an upper-case `X` (`&#X41;`), and comments may contain `--`.
 */

export interface XmlLimits {
  /** Maximum source length in UTF-16 code units. Default 16 MiB (parity with parseXml). */
  readonly maxSourceLength?: number;
  /** Maximum number of tokens (start + end + text). Default 5,000,000. */
  readonly maxTokens?: number;
  /** Maximum number of elements. Default 250,000 (parity with parseXml). */
  readonly maxElements?: number;
  /** Maximum element nesting depth. Default 256 (parity with parseXml). */
  readonly maxDepth?: number;
  /** Maximum attributes on one element, including namespace declarations. Default 512. */
  readonly maxAttributesPerElement?: number;
  /** Maximum length of one element or attribute name. Default 1024. */
  readonly maxNameLength?: number;
  /** Checked before the first token and then every 16,384 tokens. */
  readonly signal?: AbortSignal;
}

export type XmlToken =
  | { readonly type: "start"; readonly qname: string; readonly attributes: ReadonlyArray<readonly [qname: string, value: string]>; readonly selfClosing: boolean; readonly offset: number }
  | { readonly type: "end"; readonly qname: string; readonly offset: number }
  | { readonly type: "text"; readonly value: string; readonly offset: number };

const ABORT_INTERVAL = 16_384;
/** XML 1.0 NameStartChar without ":" (regular-expression character-class body; requires the "u" flag). */
export const NCNAME_START_CHARS = "A-Z_a-z\\u00C0-\\u00D6\\u00D8-\\u00F6\\u00F8-\\u02FF\\u0370-\\u037D\\u037F-\\u1FFF\\u200C\\u200D\\u2070-\\u218F\\u2C00-\\u2FEF\\u3001-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFFD\\u{10000}-\\u{EFFFF}";
/** XML 1.0 NameChar without ":" (regular-expression character-class body; requires the "u" flag). */
export const NCNAME_CHARS = `${NCNAME_START_CHARS}\\-.0-9\\u00B7\\u0300-\\u036F\\u203F\\u2040`;
const namePattern = new RegExp(`[:${NCNAME_START_CHARS}][:${NCNAME_CHARS}]*`, "uy");
const spacePattern = /[ \t\r\n]*/y;
const declarationPattern = /<!(?:DOCTYPE|ENTITY)/iy;
const referencePattern = /&(?:(amp|lt|gt|quot|apos)|#([0-9]+)|#[xX]([0-9A-Fa-f]+));/y;
const invalidCharacter = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
const predefined: Readonly<Record<string, string>> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

export class XmlTokenizer {
  readonly #source: string;
  readonly #maxTokens: number;
  readonly #maxElements: number;
  readonly #maxDepth: number;
  readonly #maxAttributes: number;
  readonly #maxNameLength: number;
  readonly #signal: AbortSignal | undefined;
  readonly #stack: string[] = [];
  #at: number;
  #tokens = 0;
  #elements = 0;
  #rootSeen = false;
  #done = false;
  #failure: unknown;

  constructor(source: string, limits: XmlLimits = {}) {
    const maxSourceLength = limit(limits.maxSourceLength, 16 * 1024 * 1024, "maxSourceLength");
    this.#maxTokens = limit(limits.maxTokens, 5_000_000, "maxTokens");
    this.#maxElements = limit(limits.maxElements, 250_000, "maxElements");
    this.#maxDepth = limit(limits.maxDepth, 256, "maxDepth");
    this.#maxAttributes = limit(limits.maxAttributesPerElement, 512, "maxAttributesPerElement");
    this.#maxNameLength = limit(limits.maxNameLength, 1024, "maxNameLength");
    this.#signal = limits.signal;
    if (source.length > maxSourceLength) throw new PreviewError("XML document exceeds the configured size limit.", "resource-limit");
    this.#source = source;
    this.#at = source.charCodeAt(0) === 0xfeff ? 1 : 0;
    const invalid = invalidCharacter.exec(source);
    if (invalid) this.#fail(`Invalid XML character U+${invalid[0].charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}`, invalid.index);
  }

  /** Number of currently open elements. */
  get depth(): number { return this.#stack.length; }

  /**
   * Returns the next token, or `undefined` at the end of a well-formed document.
   * Throws PreviewError with code "parse-error", "resource-limit", or "aborted". Errors are
   * terminal: every later call throws the same error.
   */
  next(): XmlToken | undefined {
    if (this.#failure !== undefined) throw this.#failure;
    if (this.#done) return undefined;
    try {
      if (this.#tokens % ABORT_INTERVAL === 0 && this.#signal?.aborted) throw new PreviewError("XML parsing was cancelled.", "aborted", { cause: this.#signal.reason });
      const token = this.#read();
      if (token && ++this.#tokens > this.#maxTokens) throw new PreviewError("XML token count exceeds the configured limit.", "resource-limit");
      return token;
    } catch (error) {
      this.#failure = error;
      throw error;
    }
  }

  #read(): XmlToken | undefined {
    const source = this.#source;
    while (this.#at < source.length) {
      const at = this.#at;
      if (source.charCodeAt(at) !== 0x3c /* < */) {
        const end = source.indexOf("<", at), stop = end < 0 ? source.length : end;
        this.#at = stop;
        if (!this.#stack.length) {
          for (let index = at; index < stop; index++) if (!isSpace(source.charCodeAt(index))) this.#fail("Text outside the root element", index);
          continue;
        }
        const raw = source.slice(at, stop), terminator = raw.indexOf("]]>");
        if (terminator >= 0) this.#fail("The sequence ']]>' is not allowed in text", at + terminator);
        return { type: "text", value: this.#decode(normalizeLineEnds(raw), at), offset: at };
      }
      if (source.startsWith("<!--", at)) {
        const end = source.indexOf("-->", at + 4);
        if (end < 0) this.#fail("Unclosed XML comment", at);
        this.#at = end + 3;
        continue;
      }
      if (source.startsWith("<![CDATA[", at)) {
        if (!this.#stack.length) this.#fail("CDATA section outside the root element", at);
        const end = source.indexOf("]]>", at + 9);
        if (end < 0) this.#fail("Unclosed CDATA section", at);
        this.#at = end + 3;
        if (end === at + 9) continue;
        return { type: "text", value: normalizeLineEnds(source.slice(at + 9, end)), offset: at };
      }
      if (source.startsWith("<?", at)) {
        const end = source.indexOf("?>", at + 2);
        if (end < 0) this.#fail("Unclosed processing instruction", at);
        this.#at = end + 2;
        continue;
      }
      if (source.startsWith("<!", at)) {
        declarationPattern.lastIndex = at;
        if (declarationPattern.test(source)) this.#fail("DOCTYPE and entity declarations are not supported", at);
        this.#fail("Expected XML name", at + 1);
      }
      return source.startsWith("</", at) ? this.#endTag(at) : this.#startTag(at);
    }
    if (this.#stack.length) this.#fail(`Unclosed element ${this.#stack[this.#stack.length - 1]}`, this.#at);
    if (!this.#rootSeen) this.#fail("XML document has no root element", this.#at);
    this.#done = true;
    return undefined;
  }

  #startTag(offset: number): XmlToken {
    const source = this.#source;
    this.#at = offset + 1;
    const qname = this.#name();
    const attributes: Array<readonly [string, string]> = [], seen = new Set<string>();
    let selfClosing = false, terminated = false;
    while (this.#at < source.length) {
      const hadSpace = this.#space();
      const character = source[this.#at];
      if (character === ">") { this.#at++; terminated = true; break; }
      if (character === "/") {
        if (source[this.#at + 1] !== ">") this.#fail("Expected '>' after '/'", this.#at + 1);
        this.#at += 2; selfClosing = true; terminated = true; break;
      }
      if (this.#at >= source.length) break;
      if (!hadSpace) this.#fail("Expected whitespace before attribute", this.#at);
      const attributeName = this.#name();
      this.#space();
      if (source[this.#at] !== "=") this.#fail("Expected '=' after attribute name", this.#at);
      this.#at++;
      this.#space();
      const quote = source[this.#at];
      if (quote !== '"' && quote !== "'") this.#fail("Attribute value must be quoted", this.#at);
      const start = this.#at + 1, end = source.indexOf(quote!, start);
      if (end < 0) this.#fail("Unclosed attribute value", this.#at);
      const raw = source.slice(start, end), lessThan = raw.indexOf("<");
      if (lessThan >= 0) this.#fail("Character '<' is not allowed in attribute values", start + lessThan);
      if (seen.has(attributeName)) this.#fail(`Duplicate attribute ${attributeName}`, this.#at);
      seen.add(attributeName);
      if (attributes.length >= this.#maxAttributes) throw new PreviewError("XML element has too many attributes.", "resource-limit");
      attributes.push([attributeName, this.#decode(normalizeAttributeValue(raw), start)]);
      this.#at = end + 1;
    }
    if (!terminated) this.#fail("Unclosed start tag", offset);
    if (++this.#elements > this.#maxElements) throw new PreviewError("XML element count exceeds the configured limit.", "resource-limit");
    if (this.#stack.length + 1 > this.#maxDepth) throw new PreviewError("XML nesting exceeds the configured limit.", "resource-limit");
    if (!this.#stack.length) {
      if (this.#rootSeen) this.#fail("Multiple root elements", offset);
      this.#rootSeen = true;
    }
    if (!selfClosing) this.#stack.push(qname);
    return { type: "start", qname, attributes, selfClosing, offset };
  }

  #endTag(offset: number): XmlToken {
    this.#at = offset + 2;
    const qname = this.#name();
    this.#space();
    if (this.#source[this.#at] !== ">") this.#fail("Malformed closing tag", this.#at);
    this.#at++;
    if (this.#stack[this.#stack.length - 1] !== qname) this.#fail(`Mismatched closing tag ${qname}`, offset);
    this.#stack.pop();
    return { type: "end", qname, offset };
  }

  #name(): string {
    namePattern.lastIndex = this.#at;
    const match = namePattern.exec(this.#source);
    if (!match) this.#fail("Expected XML name", this.#at);
    const name = match![0];
    if (name.length > this.#maxNameLength) throw new PreviewError("XML name exceeds the configured length limit.", "resource-limit");
    this.#at += name.length;
    return name;
  }

  /** Skips XML whitespace (S production); returns whether any was present. */
  #space(): boolean {
    spacePattern.lastIndex = this.#at;
    spacePattern.test(this.#source);
    const moved = spacePattern.lastIndex !== this.#at;
    this.#at = spacePattern.lastIndex;
    return moved;
  }

  /** Decodes the five predefined entities and character references in `value`, which starts at `offset`. */
  #decode(value: string, offset: number): string {
    let ampersand = value.indexOf("&");
    if (ampersand < 0) return value;
    let output = "", from = 0;
    while (ampersand >= 0) {
      referencePattern.lastIndex = ampersand;
      const match = referencePattern.exec(value);
      if (!match) this.#fail("Unescaped or unsupported ampersand in XML", offset + ampersand);
      output += value.slice(from, ampersand);
      if (match![1]) output += predefined[match![1]]!;
      else {
        const point = match![2] !== undefined ? Number.parseInt(match![2], 10) : Number.parseInt(match![3]!, 16);
        if (!isXmlCharacter(point)) this.#fail("Invalid XML character reference", offset + ampersand);
        output += String.fromCodePoint(point);
      }
      from = referencePattern.lastIndex;
      ampersand = value.indexOf("&", from);
    }
    return output + value.slice(from);
  }

  #fail(message: string, at: number): never {
    throw new PreviewError(`${message} (at character ${at}).`, "parse-error");
  }
}

function isSpace(code: number): boolean { return code === 0x20 || code === 0x9 || code === 0xa || code === 0xd; }

/** XML 1.0 §2.11: CRLF and lone CR become LF. */
function normalizeLineEnds(text: string): string { return text.includes("\r") ? text.replace(/\r\n?/g, "\n") : text; }

/** XML 1.0 §3.3.3 for attributes without a DTD: each literal TAB, LF, CR, or CRLF becomes one space. */
function normalizeAttributeValue(value: string): string { return /[\t\n\r]/.test(value) ? value.replace(/\r\n|[\t\n\r]/g, " ") : value; }

function isXmlCharacter(point: number): boolean {
  return point === 0x9 || point === 0xa || point === 0xd || (point >= 0x20 && point <= 0xd7ff) || (point >= 0xe000 && point <= 0xfffd) || (point >= 0x10000 && point <= 0x10ffff);
}

function limit(value: number | undefined, fallback: number, name: string): number {
  if (value === undefined) return fallback;
  if (!Number.isSafeInteger(value) || value < 1) throw new RangeError(`${name} must be a positive safe integer.`);
  return value;
}
