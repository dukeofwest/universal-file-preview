import { PreviewError } from "../errors.js";

export type XmlChild = XmlNode | string;
export interface XmlNode { readonly name: string; readonly localName: string; readonly attributes: Readonly<Record<string, string>>; readonly children: XmlChild[]; }
const namePattern = /^[A-Za-z_:][A-Za-z0-9_.:\-]*/;

/** Small non-validating XML 1.0 parser for document package parts. DTDs/entities are deliberately disabled. */
export function parseXml(source: string, limits: { maxNodes?: number; maxDepth?: number; maxTextLength?: number } = {}): XmlNode {
  const maxNodes = limits.maxNodes ?? 250_000, maxDepth = limits.maxDepth ?? 256, maxTextLength = limits.maxTextLength ?? 16 * 1024 * 1024;
  if (source.length > maxTextLength) throw new PreviewError("XML document exceeds the configured size limit.", "resource-limit");
  let at = source.charCodeAt(0) === 0xfeff ? 1 : 0, nodes = 0, textLength = 0;
  const stack: XmlNode[] = [];
  let root: XmlNode | undefined;
  const fail = (message: string): never => { throw new PreviewError(`${message} (at character ${at}).`, "parse-error"); };
  const skipSpace = () => { while (/\s/.test(source[at] ?? "")) at++; };
  const readName = (): string => { const match = namePattern.exec(source.slice(at)); if (!match) return fail("Expected XML name"); at += match[0].length; return match[0]; };
  const appendText = (text: string) => {
    if (!text) return;
    textLength += text.length;
    if (textLength > maxTextLength) throw new PreviewError("XML text exceeds the configured size limit.", "resource-limit");
    if (stack.length) stack[stack.length - 1]!.children.push(decodeEntities(text, fail));
    else if (text.trim()) fail("Text outside the root element");
  };

  while (at < source.length) {
    if (source[at] !== "<") { const end = source.indexOf("<", at); appendText(source.slice(at, end < 0 ? source.length : end)); at = end < 0 ? source.length : end; continue; }
    if (source.startsWith("<!--", at)) { const end = source.indexOf("-->", at + 4); if (end < 0) fail("Unclosed XML comment"); at = end + 3; continue; }
    if (source.startsWith("<![CDATA[", at)) { const end = source.indexOf("]]>", at + 9); if (end < 0) fail("Unclosed CDATA section"); appendText(source.slice(at + 9, end)); at = end + 3; continue; }
    if (source.startsWith("<?", at)) { const end = source.indexOf("?>", at + 2); if (end < 0) fail("Unclosed processing instruction"); at = end + 2; continue; }
    if (/^<!DOCTYPE|^<!ENTITY/i.test(source.slice(at))) fail("DOCTYPE and entity declarations are not supported");
    if (source.startsWith("</", at)) {
      at += 2; const name = readName(); skipSpace(); if (source[at++] !== ">") fail("Malformed closing tag");
      const open = stack.pop(); if (!open || open.name !== name) fail(`Mismatched closing tag ${name}`); continue;
    }
    at++;
    const name = readName(), attributes: Record<string, string> = {};
    let selfClosing = false, terminated = false;
    while (at < source.length) {
      skipSpace();
      if (source.startsWith("/>", at)) { at += 2; selfClosing = true; terminated = true; break; }
      if (source[at] === ">") { at++; terminated = true; break; }
      const attrName = readName(); skipSpace(); if (source[at++] !== "=") fail("Expected '=' after attribute name"); skipSpace();
      const quote = source[at++] as string; if (quote !== "'" && quote !== '"') fail("Attribute value must be quoted");
      const end = source.indexOf(quote, at); if (end < 0) fail("Unclosed attribute value");
      if (Object.hasOwn(attributes, attrName)) fail(`Duplicate attribute ${attrName}`);
      attributes[attrName] = decodeEntities(source.slice(at, end), fail); at = end + 1;
    }
    if (!terminated) fail("Unclosed start tag");
    if (++nodes > maxNodes) throw new PreviewError("XML element count exceeds the configured limit.", "resource-limit");
    if (stack.length + 1 > maxDepth) throw new PreviewError("XML nesting exceeds the configured limit.", "resource-limit");
    const node: XmlNode = { name, localName: name.slice(name.lastIndexOf(":") + 1), attributes, children: [] };
    if (stack.length) stack[stack.length - 1]!.children.push(node);
    else if (root) fail("Multiple root elements");
    else root = node;
    if (!selfClosing) stack.push(node);
  }
  if (stack.length) fail(`Unclosed element ${stack[stack.length - 1]!.name}`);
  if (!root) fail("XML document has no root element");
  return root!;
}

function decodeEntities(value: string, fail: (message: string) => never): string {
  if (/&(?!(?:amp|lt|gt|quot|apos|#(?:[0-9]+|x[0-9a-f]+));)/i.test(value)) fail("Unescaped or unsupported ampersand in XML");
  return value.replace(/&([^;]+);/g, (_whole, entity: string) => {
    if (entity === "amp") return "&"; if (entity === "lt") return "<"; if (entity === "gt") return ">";
    if (entity === "quot") return '"'; if (entity === "apos") return "'";
    if (entity.startsWith("#")) {
      const hex = entity[1]?.toLowerCase() === "x", point = Number.parseInt(entity.slice(hex ? 2 : 1), hex ? 16 : 10);
      if (!Number.isInteger(point) || point <= 0 || point > 0x10ffff || (point >= 0xd800 && point <= 0xdfff)) return fail("Invalid XML character reference");
      return String.fromCodePoint(point);
    }
    return fail(`Unsupported XML entity &${entity};`);
  });
}

export function elements(node: XmlNode, localName: string): XmlNode[] { return node.children.filter((child): child is XmlNode => typeof child !== "string" && child.localName === localName); }
export function firstElement(node: XmlNode, localName: string): XmlNode | undefined { return elements(node, localName)[0]; }
export function textContent(node: XmlNode): string { return node.children.map(child => typeof child === "string" ? child : textContent(child)).join(""); }
