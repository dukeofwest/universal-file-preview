// The SVG sanitizer of 0.2.x (DOMParser and XMLSerializer), kept unchanged as a reference for the browser
// differential tests of the DOM-free sanitizer (scripts/browser-security.mjs). Test-only; browser only.
// Generated from src/renderers/image-url.ts at commit 15fd7ce by removing the type annotations.

const SVG_NAMESPACE = "http://www.w3.org/2000/svg";
const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";
const svgElements = new Set(["svg", "g", "path", "rect", "circle", "ellipse", "line", "polyline", "polygon", "text", "tspan", "defs", "linearGradient", "radialGradient", "stop", "clipPath", "mask", "title", "desc"]);
/** Value rules per allowed attribute. Attributes must have no namespace; anything not listed is removed. */
const colorValue = String.raw`(?:none|currentColor|transparent|inherit|#[0-9a-f]{3,8}|[a-z]{3,20}|(?:rgba?|hsla?)\([0-9.,%\s+-]*\))`;
const localReference = String.raw`url\(#[A-Za-z0-9_.:-]{1,256}\)`;
const paint = new RegExp(String.raw`^(?:${colorValue}|${localReference}(?:\s+${colorValue})?)$`, "i");
const color = new RegExp(`^${colorValue}$`, "i");
const reference = new RegExp(`^${localReference}$`);
/** Numbers, path data, transforms and keywords: no quotes, backslashes, colons, semicolons or angle brackets. */
const plain = /^[A-Za-z0-9\s.,%()+#-]{0,100000}$/;
const svgAttributes = new Map([
  ...["viewBox", "width", "height", "x", "y", "x1", "x2", "y1", "y2", "cx", "cy", "r", "rx", "ry", "d", "points", "fill-rule", "fill-opacity", "stroke-width", "stroke-linecap", "stroke-linejoin", "stroke-dasharray", "stroke-opacity", "opacity", "transform", "font-size", "font-weight", "text-anchor", "stop-opacity", "offset", "preserveAspectRatio"].map(name => [name, plain]),
  ["fill", paint], ["stroke", paint], ["stop-color", color], ["clip-path", reference], ["mask", reference],
  ["font-family", /^[\p{L}\p{M}\p{N} ,'"_.-]{1,200}$/u],
  ["id", /^[A-Za-z_][A-Za-z0-9_.:-]{0,255}$/],
]);

function attributeAllowed(attribute) {
  // Namespace declarations may only bind the SVG namespace (the serializer keeps a prefix such as
  // <s:svg xmlns:s="...">); no other attribute may carry a namespace.
  if (attribute.namespaceURI === XMLNS_NAMESPACE) return attribute.value === SVG_NAMESPACE;
  if (attribute.namespaceURI !== null) return false;
  const rule = svgAttributes.get(attribute.localName);
  return rule !== undefined && !/url\s*\(/i.test(attribute.value.replace(/url\(#/gi, "")) && rule.test(attribute.value);
}

/**
 * Walks `node` and counts everything outside the allowlist: elements outside the SVG namespace or
 * not listed, attributes that are namespaced, not listed or have a value outside their rule, and
 * any child node other than text (comments, processing instructions, CDATA). With `remove`, those
 * nodes and attributes are also removed.
 */
function sweep(node, remove) {
  let violations = 0;
  for (const child of [...node.childNodes]) {
    if (child.nodeType === 3) continue; // text
    if (child.nodeType === 1 && child.namespaceURI === SVG_NAMESPACE && svgElements.has(child.localName)) { violations += sweep(child, remove); continue; }
    violations++;
    if (remove) child.remove();
  }
  for (const attribute of [...node.attributes]) {
    if (attributeAllowed(attribute)) continue;
    violations++;
    if (remove) node.removeAttributeNode(attribute);
  }
  return violations;
}

function parseSvgRoot(source) {
  const parsed = new DOMParser().parseFromString(source, "image/svg+xml"), root = parsed.documentElement;
  if (root.namespaceURI !== SVG_NAMESPACE || root.localName !== "svg" || parsed.getElementsByTagNameNS("*", "parsererror").length) throw new Error("Invalid SVG image.");
  return root;
}

/**
 * Reduces SVG to an element and attribute allowlist, then parses the result again and rejects it
 * unless it passes the same allowlist unchanged. Throws PreviewError for input that is not a
 * well-formed SVG document or whose sanitized form does not validate.
 */
export function referenceSanitizeSvg(bytes) {
  let source;
  try { source = new TextDecoder("utf-8", { fatal: true }).decode(bytes); } catch { throw new Error("Invalid SVG image."); }
  const root = parseSvgRoot(source);
  sweep(root, true);
  // Only the root element is serialized, so a DOCTYPE and prolog processing instructions are dropped.
  const output = new XMLSerializer().serializeToString(root);
  if (sweep(parseSvgRoot(output), false) !== 0) throw new Error("Invalid SVG image.");
  return new TextEncoder().encode(output);
}
