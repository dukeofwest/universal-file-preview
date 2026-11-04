import { PreviewError } from "../errors.js";
import { NamespaceScope, resolveValuePrefix, type ExpandedName } from "./xml-namespaces.js";
import { XmlTokenizer, type XmlLimits, type XmlToken } from "./xml-tokenizer.js";

/**
 * Namespace-aware XML tree (implementation specification §4.1, milestone M3).
 *
 * Internal module: not exported from the package entry point (decision D2). Used for small
 * OpenDocument parts (manifest, styles, meta, settings). Large parts are streamed (M4).
 */

export interface NsXmlAttribute { readonly uri: string; readonly local: string; readonly qname: string; readonly value: string; }
export interface NsXmlNode {
  readonly uri: string;
  readonly local: string;
  readonly qname: string;
  /** Ordinary attributes in document order. Namespace declarations are not included. */
  readonly attributes: readonly NsXmlAttribute[];
  /** Child elements and text. Adjacent text and CDATA are merged into one string. */
  readonly children: ReadonlyArray<NsXmlNode | string>;
  /** In-scope prefix → URI bindings (default namespace under ""), for QName-valued attributes. Shared between elements when unchanged. */
  readonly namespaces: ReadonlyMap<string, string>;
}

type MutableNode = Omit<NsXmlNode, "children"> & { children: Array<NsXmlNode | string> };

/** Parses `source` into a namespace-resolved tree. Limits and errors are those of XmlTokenizer plus namespace well-formedness errors (parse-error). */
export function parseXmlNs(source: string, limits?: XmlLimits): NsXmlNode {
  const tokenizer = new XmlTokenizer(source, limits), scope = new NamespaceScope(), stack: MutableNode[] = [];
  let root: MutableNode | undefined;
  for (let token = tokenizer.next(); token; token = tokenizer.next()) {
    if (token.type === "start") {
      const node = openElement(token, scope), parent = stack[stack.length - 1];
      if (parent) parent.children.push(node); else root = node;
      if (token.selfClosing) scope.exit(); else stack.push(node);
    } else if (token.type === "end") {
      stack.pop();
      scope.exit();
    } else appendText(stack[stack.length - 1]!, token.value);
  }
  return root!;
}

/**
 * Resolves a start token into an element with no children yet and enters its namespace scope.
 * The caller must call `scope.exit()` when the element ends (immediately for self-closing
 * elements). Shared with the stream reader so captured subtrees match parseXmlNs exactly.
 */
export function openElement(token: Extract<XmlToken, { type: "start" }>, scope: NamespaceScope): MutableNode {
  scope.enter(token.attributes, token.offset);
  const name = scope.element(token.qname, token.offset), attributes: NsXmlAttribute[] = [], seen = new Set<string>();
  for (const [qname, value] of token.attributes) {
    if (qname === "xmlns" || qname.startsWith("xmlns:")) continue;
    const resolved = scope.attribute(qname, token.offset), key = `${resolved.uri}\u0000${resolved.local}`;
    if (seen.has(key)) throw new PreviewError(`Duplicate attribute {${resolved.uri}}${resolved.local} (at character ${token.offset}).`, "parse-error");
    seen.add(key);
    attributes.push({ uri: resolved.uri, local: resolved.local, qname, value });
  }
  return { uri: name.uri, local: name.local, qname: token.qname, attributes, children: [], namespaces: scope.snapshot() };
}

/** Appends text, merging with a preceding text child. */
export function appendText(node: MutableNode, text: string): void {
  if (!text) return;
  const last = node.children.length - 1;
  if (last >= 0 && typeof node.children[last] === "string") node.children[last] += text;
  else node.children.push(text);
}

/** The value of the attribute `{uri}local`, or undefined. */
export function nsAttr(node: NsXmlNode | undefined, uri: string, local: string): string | undefined {
  if (!node) return undefined;
  for (const attribute of node.attributes) if (attribute.local === local && attribute.uri === uri) return attribute.value;
  return undefined;
}

/** Child elements named `{uri}local`, in document order. */
export function nsChildren(node: NsXmlNode, uri: string, local: string): NsXmlNode[] {
  const output: NsXmlNode[] = [];
  for (const child of node.children) if (typeof child !== "string" && child.local === local && child.uri === uri) output.push(child);
  return output;
}

/** The first child element named `{uri}local`, or undefined. */
export function nsFirst(node: NsXmlNode | undefined, uri: string, local: string): NsXmlNode | undefined {
  if (!node) return undefined;
  for (const child of node.children) if (typeof child !== "string" && child.local === local && child.uri === uri) return child;
  return undefined;
}

/** Concatenated text of all descendants, in document order. */
export function nsText(node: NsXmlNode): string {
  let output = "";
  for (const child of node.children) output += typeof child === "string" ? child : nsText(child);
  return output;
}

/** Resolves the prefix of QName-valued attribute content (e.g. `of:=SUM([.A1])`) against the bindings in scope at `node`. */
export function nsValuePrefix(node: NsXmlNode, value: string): ExpandedName | undefined { return resolveValuePrefix(node.namespaces, value); }
