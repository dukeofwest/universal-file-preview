import { PreviewError } from "../errors.js";
import { appendText, openElement, type NsXmlAttribute, type NsXmlNode } from "../xml/ns-tree.js";
import { NamespaceScope } from "../xml/xml-namespaces.js";
import { XmlTokenizer } from "../xml/xml-tokenizer.js";

/**
 * The only place where image bytes from a file become blob URLs.
 *
 * A blob URL has the origin of the page that created it. If one of type image/svg+xml is opened
 * as a document (for example "Open image in new tab"), any script in it runs in the host origin.
 * So every image URL is created here, and:
 * - only raster types and SVG are accepted;
 * - raster bytes must start with a known signature, and the blob gets the type of the actual
 *   content rather than the type declared by a file extension or content type;
 * - SVG is always passed through the sanitizer first, whatever its origin (standalone file or
 *   media inside a DOCX, PPTX or XLSX package).
 */

const SVG = "image/svg+xml";
const rasterTypes = new Set(["image/png", "image/jpeg", "image/gif", "image/bmp", "image/webp"]);

/** Raster image type from the leading signature bytes, or undefined when none matches. */
export function sniffRasterType(bytes: Uint8Array): string | undefined {
  const at = (offset: number, signature: readonly number[]) => bytes.length >= offset + signature.length && signature.every((value, index) => bytes[offset + index] === value);
  if (at(0, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return "image/png";
  if (at(0, [0xff, 0xd8, 0xff])) return "image/jpeg";
  if (at(0, [0x47, 0x49, 0x46, 0x38]) && (bytes[4] === 0x37 || bytes[4] === 0x39) && bytes[5] === 0x61) return "image/gif";
  if (at(0, [0x52, 0x49, 0x46, 0x46]) && at(8, [0x57, 0x45, 0x42, 0x50])) return "image/webp";
  if (at(0, [0x42, 0x4d]) && bytes.length >= 26) return "image/bmp";
  return undefined;
}

/**
 * The blob type to use for image bytes declared as `declaredType`, or undefined when they must not
 * be shown. A declared raster or SVG type whose bytes carry a raster signature takes the sniffed
 * raster type; image/svg+xml is returned only for declared SVG without a raster signature, and the
 * caller must sanitize it.
 */
export function imageBlobType(bytes: Uint8Array, declaredType: string): string | undefined {
  const declared = declaredType.toLowerCase().split(";")[0]!.trim();
  if (declared !== SVG && !rasterTypes.has(declared)) return undefined;
  return sniffRasterType(bytes) ?? (declared === SVG ? SVG : undefined);
}

/** Largest raster image shown: 40 megapixels (about 160 MB once decoded), the same limit as for TIFF conversion. */
export const MAX_IMAGE_PIXELS = 40_000_000;
/** Decoded pixels of all images a document has alive at once (the URLs in its list). */
export const MAX_DOCUMENT_PIXELS = 120_000_000;
const pixelsByUrl = new Map<string, number>();

/** Pixel dimensions from a raster image header, or undefined when the header cannot be read. */
export function imageDimensions(bytes: Uint8Array, type: string): { width: number; height: number } | undefined {
  const u16le = (at: number) => bytes[at]! | bytes[at + 1]! << 8, u16be = (at: number) => bytes[at]! << 8 | bytes[at + 1]!;
  const u24le = (at: number) => u16le(at) | bytes[at + 2]! << 16, u32be = (at: number) => (u16be(at) * 0x10000) + u16be(at + 2);
  const i32le = (at: number) => u16le(at) | bytes[at + 2]! << 16 | bytes[at + 3]! << 24;
  const has = (length: number) => bytes.length >= length;
  if (type === "image/png") return has(24) ? { width: u32be(16), height: u32be(20) } : undefined;
  if (type === "image/gif") return has(10) ? { width: u16le(6), height: u16le(8) } : undefined;
  if (type === "image/bmp") {
    if (!has(26)) return undefined;
    // A 12-byte (OS/2) header stores 16-bit sizes; later headers store signed 32-bit sizes (negative height = top-down).
    return i32le(14) === 12 ? { width: u16le(18), height: u16le(20) } : { width: Math.abs(i32le(18)), height: Math.abs(i32le(22)) };
  }
  if (type === "image/webp") {
    const chunk = String.fromCharCode(...bytes.subarray(12, 16));
    if (chunk === "VP8 " && has(30)) return { width: u16le(26) & 0x3fff, height: u16le(28) & 0x3fff };
    if (chunk === "VP8L" && has(25)) { const b1 = bytes[22]!, b2 = bytes[23]!, b3 = bytes[24]!; return { width: 1 + ((b1 & 0x3f) << 8 | bytes[21]!), height: 1 + ((b3 & 0x0f) << 10 | b2 << 2 | (b1 & 0xc0) >> 6) }; }
    if (chunk === "VP8X" && has(30)) return { width: 1 + u24le(24), height: 1 + u24le(27) };
    return undefined;
  }
  if (type === "image/jpeg") {
    // Walk the marker segments to the first start-of-frame (SOF0-SOF15, except DHT, JPG and DAC).
    for (let at = 2; at + 9 < bytes.length;) {
      if (bytes[at] !== 0xff) return undefined;
      const marker = bytes[at + 1]!;
      if (marker === 0xff) { at++; continue; }
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) return { width: u16be(at + 7), height: u16be(at + 5) };
      if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { at += 2; continue; }
      if (marker === 0xda || marker === 0xd9) return undefined;
      at += 2 + u16be(at + 2);
    }
  }
  return undefined;
}

/**
 * A blob URL for image bytes, or undefined for unsupported, mislabelled, invalid or oversized
 * images. `liveUrls` is the document's list of image URLs that are still shown; the new image is
 * refused if it would take their decoded pixels over MAX_DOCUMENT_PIXELS. The caller adds the URL
 * to that list and releases it with revokeImageUrls.
 */
export function createImageUrl(bytes: Uint8Array, declaredType: string, liveUrls: readonly string[] = []): string | undefined {
  const type = imageBlobType(bytes, declaredType);
  if (!type) return undefined;
  let content = bytes, pixels = 0;
  if (type === SVG) { try { content = sanitizeSvg(bytes); } catch { return undefined; } }
  else {
    // Decoding is left to the browser; refuse images whose header declares more pixels than allowed.
    const size = imageDimensions(bytes, type);
    if (size) {
      pixels = size.width * size.height;
      if (!(pixels > 0) || pixels > MAX_IMAGE_PIXELS) return undefined;
      const live = liveUrls.reduce((sum, url) => sum + (pixelsByUrl.get(url) ?? 0), 0);
      if (live + pixels > MAX_DOCUMENT_PIXELS) return undefined;
    }
  }
  const copy = new Uint8Array(content.byteLength); copy.set(content);
  const url = URL.createObjectURL(new Blob([copy.buffer], { type }));
  if (pixels) pixelsByUrl.set(url, pixels);
  return url;
}

/** Revokes image URLs made by createImageUrl and empties the list. */
export function revokeImageUrls(urls: string[]): void {
  for (const url of urls.splice(0)) { URL.revokeObjectURL(url); pixelsByUrl.delete(url); }
}

const SVG_NAMESPACE = "http://www.w3.org/2000/svg";
const svgElements = new Set(["svg", "g", "path", "rect", "circle", "ellipse", "line", "polyline", "polygon", "text", "tspan", "defs", "linearGradient", "radialGradient", "stop", "clipPath", "mask", "title", "desc"]);
/** Value rules per allowed attribute. Attributes must have no namespace; anything not listed is removed. */
const colorValue = String.raw`(?:none|currentColor|transparent|inherit|#[0-9a-f]{3,8}|[a-z]{3,20}|(?:rgba?|hsla?)\([0-9.,%\s+-]*\))`;
const localReference = String.raw`url\(#[A-Za-z0-9_.:-]{1,256}\)`;
const paint = new RegExp(String.raw`^(?:${colorValue}|${localReference}(?:\s+${colorValue})?)$`, "i");
const color = new RegExp(`^${colorValue}$`, "i");
const reference = new RegExp(`^${localReference}$`);
/** Numbers, path data, transforms and keywords: no quotes, backslashes, colons, semicolons or angle brackets. */
const plain = /^[A-Za-z0-9\s.,%()+#-]{0,100000}$/;
const svgAttributes = new Map<string, RegExp>([
  ...["viewBox", "width", "height", "x", "y", "x1", "x2", "y1", "y2", "cx", "cy", "r", "rx", "ry", "d", "points", "fill-rule", "fill-opacity", "stroke-width", "stroke-linecap", "stroke-linejoin", "stroke-dasharray", "stroke-opacity", "opacity", "transform", "font-size", "font-weight", "text-anchor", "stop-opacity", "offset", "preserveAspectRatio"].map(name => [name, plain] as const),
  ["fill", paint], ["stroke", paint], ["stop-color", color], ["clip-path", reference], ["mask", reference],
  ["font-family", /^[\p{L}\p{M}\p{N} ,'"_.-]{1,200}$/u],
  ["id", /^[A-Za-z_][A-Za-z0-9_.:-]{0,255}$/],
]);

function attributeAllowed(attribute: NsXmlAttribute): boolean {
  // Only unprefixed attributes (no namespace). Namespace declarations are not attributes here: the output
  // declares the SVG namespace once, on the root, and nothing else.
  if (attribute.uri !== "") return false;
  const rule = svgAttributes.get(attribute.local);
  return rule !== undefined && !/url\s*\(/i.test(attribute.value.replace(/url\(#/gi, "")) && rule.test(attribute.value);
}

const invalid = (cause?: unknown) => new PreviewError("Invalid SVG image.", "invalid-file", cause === undefined ? undefined : { cause });

/**
 * Parses SVG source without the DOM (no DOMParser, so Trusted Types do not apply). The tokenizer rejects
 * DOCTYPE and entity declarations, malformed XML and characters XML does not allow; it skips comments and
 * processing instructions. CDATA sections are dropped here, as the DOM sanitizer removed them.
 */
function parseSvg(source: string): NsXmlNode {
  let root: NsXmlNode | undefined;
  try {
    const tokenizer = new XmlTokenizer(source), scope = new NamespaceScope(), stack: Array<ReturnType<typeof openElement>> = [];
    for (let token = tokenizer.next(); token; token = tokenizer.next()) {
      if (token.type === "start") {
        const node = openElement(token, scope), parent = stack[stack.length - 1];
        if (parent) parent.children.push(node); else root = node;
        if (token.selfClosing) scope.exit(); else stack.push(node);
      } else if (token.type === "end") {
        stack.pop();
        scope.exit();
      } else if (!source.startsWith("<![CDATA[", token.offset)) appendText(stack[stack.length - 1]!, token.value);
    }
  } catch (error) { throw invalid(error); }
  if (!root || root.uri !== SVG_NAMESPACE || root.local !== "svg") throw invalid();
  return root;
}

/** Characters XML 1.0 does not allow: C0 controls other than tab, line feed and carriage return; U+FFFE, U+FFFF; unpaired surrogates. */
const invalidCharacter = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
function xmlCharacters(value: string): string { if (invalidCharacter.test(value)) throw invalid(); return value; }
const escapeText = (value: string) => xmlCharacters(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r/g, "&#13;");
const escapeAttribute = (value: string) => xmlCharacters(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/\t/g, "&#9;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;");

/**
 * Writes the allowed part of `node`: SVG-namespace elements in the allowlist with their allowed attributes, and
 * text. Returns the markup and how many elements and attributes were left out.
 */
function serialize(node: NsXmlNode, isRoot: boolean): { markup: string; removed: number } {
  let removed = 0, markup = `<${node.local}${isRoot ? ` xmlns="${SVG_NAMESPACE}"` : ""}`;
  for (const attribute of node.attributes) {
    if (attributeAllowed(attribute)) markup += ` ${attribute.local}="${escapeAttribute(attribute.value)}"`;
    else removed++;
  }
  let content = "";
  for (const child of node.children) {
    if (typeof child === "string") { content += escapeText(child); continue; }
    if (child.uri !== SVG_NAMESPACE || !svgElements.has(child.local)) { removed++; continue; }
    const inner = serialize(child, false);
    content += inner.markup; removed += inner.removed;
  }
  return { markup: content ? `${markup}>${content}</${node.local}>` : `${markup}/>`, removed };
}

/**
 * Reduces SVG to an element and attribute allowlist and serializes the result, then parses that output again and
 * rejects it unless it passes the same allowlist unchanged and serializes to exactly the same markup. Throws
 * PreviewError for input that is not a well-formed SVG document without a DOCTYPE, or whose sanitized form
 * does not validate.
 */
export function sanitizeSvg(bytes: Uint8Array): Uint8Array {
  let source: string;
  try { source = new TextDecoder("utf-8", { fatal: true }).decode(bytes); } catch (error) { throw invalid(error); }
  const output = serialize(parseSvg(source), true).markup;
  const check = serialize(parseSvg(output), true);
  if (check.removed !== 0 || check.markup !== output) throw invalid();
  return new TextEncoder().encode(output);
}
