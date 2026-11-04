import { PreviewError } from "../errors.js";
import { NamespaceScope, type ExpandedName } from "./xml-namespaces.js";
import { appendText, openElement, type NsXmlAttribute, type NsXmlNode } from "./ns-tree.js";
import { XmlTokenizer, type XmlLimits } from "./xml-tokenizer.js";

/**
 * Event-based, namespace-aware XML reader with bounded subtree capture
 * (implementation specification §4.1, milestone M4).
 *
 * Internal module: not exported from the package entry point (decision D2). Large parts such
 * as an ODS content.xml are read without building a full tree: the handler sees start, end
 * and text events and can capture one bounded subtree at a time (for example a single
 * table cell) or skip subtrees it does not need.
 *
 * Every element is resolved exactly as parseXmlNs resolves it, including inside skipped and
 * captured subtrees, so streamXml rejects exactly the documents parseXmlNs rejects.
 *
 * The reader is synchronous and single-pass. Handlers must not call streamXml re-entrantly
 * for the same source; an exception thrown by a handler stops the stream and propagates.
 */

export interface XmlStartElement {
  readonly name: ExpandedName;
  readonly attributes: readonly NsXmlAttribute[];
  readonly namespaces: ReadonlyMap<string, string>;
  readonly selfClosing: boolean;
}

/**
 * - `undefined` (or no return): descendants produce events; `end` fires when the element ends.
 * - `"capture"`: the element and its descendants are delivered once to `captured` as an
 *   NsXmlNode, in place of their start/end/text events. `end` does not fire for it.
 * - `"skip"`: descendants produce no events; `end` still fires for the element.
 */
export type StartAction = void | "capture" | "skip";

/** `depth` is the element's nesting level (the root element is 1); for text it is the enclosing element's level. */
export interface XmlStreamHandler {
  start?(element: XmlStartElement, depth: number): StartAction;
  /** Also fires immediately after `start` for self-closing elements that are not captured. */
  end?(name: ExpandedName, depth: number): void;
  /** Text and CDATA outside captured or skipped subtrees. Adjacent text and CDATA may arrive as separate events. */
  text?(value: string, depth: number): void;
  captured?(node: NsXmlNode, depth: number): void;
}

export interface XmlStreamLimits extends XmlLimits {
  /** Maximum elements in one captured subtree, including its root. Default 10,000. */
  readonly maxCaptureElements?: number;
  /**
   * Elements whose content is not kept in captured subtrees (for example embedded base64
   * images): the element appears with its attributes but no children. The content is still
   * read and checked for well-formedness.
   */
  readonly emptyInCapture?: (uri: string, local: string) => boolean;
}

type MutableNode = ReturnType<typeof openElement>;
/** `emptying`: nesting depth inside an element whose content is not kept (emptyInCapture). */
type Capture = { readonly root: MutableNode; readonly stack: MutableNode[]; readonly depth: number; count: number; emptying: number };

export function streamXml(source: string, handler: XmlStreamHandler, limits: XmlStreamLimits = {}): void {
  const maxCapture = limits.maxCaptureElements ?? 10_000;
  if (!Number.isSafeInteger(maxCapture) || maxCapture < 1) throw new RangeError("maxCaptureElements must be a positive safe integer.");
  const tokenizer = new XmlTokenizer(source, limits), scope = new NamespaceScope(), emptyInCapture = limits.emptyInCapture;
  // Names of open elements that produce an `end` event (normal elements and skipped roots).
  const open: ExpandedName[] = [];
  let depth = 0, skipping = 0, capture: Capture | undefined;

  for (let token = tokenizer.next(); token; token = tokenizer.next()) {
    if (token.type === "text") {
      if (capture) { if (!capture.emptying) appendText(capture.stack[capture.stack.length - 1]!, token.value); }
      else if (!skipping) handler.text?.(token.value, depth);
      continue;
    }

    if (token.type === "end") {
      scope.exit();
      if (capture?.emptying) { capture.emptying--; depth--; continue; }
      if (capture) {
        capture.stack.pop();
        if (!capture.stack.length) { const done = capture; capture = undefined; handler.captured?.(done.root, done.depth); }
      } else if (skipping) {
        if (--skipping === 0) handler.end?.(open.pop()!, depth);
      } else handler.end?.(open.pop()!, depth);
      depth--;
      continue;
    }

    const node = openElement(token, scope);
    depth++;
    if (capture) {
      if (capture.emptying) {
        if (token.selfClosing) { scope.exit(); depth--; } else capture.emptying++;
        continue;
      }
      if (++capture.count > maxCapture) throw new PreviewError("Captured XML subtree exceeds the configured element limit.", "resource-limit");
      capture.stack[capture.stack.length - 1]!.children.push(node);
      if (token.selfClosing) { scope.exit(); depth--; }
      else if (emptyInCapture?.(node.uri, node.local)) capture.emptying = 1;
      else capture.stack.push(node);
      continue;
    }
    if (skipping) {
      if (token.selfClosing) { scope.exit(); depth--; } else skipping++;
      continue;
    }

    const name = expandedName(node);
    const action = handler.start?.({ name, attributes: node.attributes, namespaces: node.namespaces, selfClosing: token.selfClosing }, depth);
    if (action === "capture") {
      if (token.selfClosing) { scope.exit(); handler.captured?.(node, depth); depth--; }
      else capture = { root: node, stack: [node], depth, count: 1, emptying: 0 };
    } else if (token.selfClosing) {
      scope.exit();
      handler.end?.(name, depth);
      depth--;
    } else {
      open.push(name);
      if (action === "skip") skipping = 1;
    }
  }
}

function expandedName(node: NsXmlNode): ExpandedName {
  const colon = node.qname.indexOf(":");
  return { uri: node.uri, local: node.local, prefix: colon < 0 ? "" : node.qname.slice(0, colon) };
}
