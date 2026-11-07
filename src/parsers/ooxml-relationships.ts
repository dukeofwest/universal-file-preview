import type { XmlNode } from "../xml/xml-parser.js";

/**
 * Namespace-aware reading of relationship references in Office Open XML parts. Internal module.
 *
 * The XML parser used for OOXML keeps qualified names without resolving namespaces. These helpers resolve a
 * prefix as XML namespaces do: by the nearest declaration on the element or one of its ancestors up to the part's
 * root, so an attribute in an unrelated namespace is not taken for a relationship reference.
 */

/**
 * The parent of each element of a part, or null when no element below the root declares a prefix (the common
 * case: Office declares prefixes on the root, so the element and the root are the only scopes to look at). Built
 * once per part root; the map is kept only when it is needed.
 */
const parentsByRoot = new WeakMap<XmlNode, Map<XmlNode, XmlNode> | null>();
function declaresPrefix(node: XmlNode): boolean {
  for (const name in node.attributes) if (name.startsWith("xmlns:")) return true;
  return false;
}
function parentsOf(root: XmlNode): Map<XmlNode, XmlNode> | null {
  const cached = parentsByRoot.get(root);
  if (cached !== undefined) return cached;
  let nested = false;
  const stack: XmlNode[] = [root];
  while (stack.length && !nested) {
    for (const child of stack.pop()!.children) {
      if (typeof child === "string") continue;
      if (declaresPrefix(child)) { nested = true; break; }
      stack.push(child);
    }
  }
  let parents: Map<XmlNode, XmlNode> | null = null;
  if (nested) {
    parents = new Map();
    const pending: XmlNode[] = [root];
    while (pending.length) {
      const parent = pending.pop()!;
      for (const child of parent.children) if (typeof child !== "string") { parents.set(child, parent); pending.push(child); }
    }
  }
  parentsByRoot.set(root, parents);
  return parents;
}

/** The namespace bound to `prefix` where `node` stands, or undefined when no element in scope declares it. */
function namespaceOf(prefix: string, node: XmlNode, root: XmlNode | undefined): string | undefined {
  const declaration = `xmlns:${prefix}`, parents = root ? parentsOf(root) : null;
  for (let current: XmlNode | undefined = node; current;) {
    if (Object.hasOwn(current.attributes, declaration)) return current.attributes[declaration];
    if (current === root) break;
    // Up to the parent; without parent links (or for an element outside the root's tree), straight to the root.
    current = parents?.get(current) ?? root;
  }
  return undefined;
}

/** The relationships namespace (transitional and strict Office Open XML). */
const RELATIONSHIP_NAMESPACES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);

/**
 * The relationship id of an element (`r:id`), by preference: an attribute with local name "id" whose prefix is
 * bound to the relationships namespace; otherwise the conventional `r:id`, whatever `r` is bound to (files written
 * by hand or by tools often bind it loosely, and 0.2.x accepted them); otherwise an "id" with an undeclared prefix.
 * An "id" whose prefix is declared for another namespace is never used.
 */
export function relationshipId(node: XmlNode | undefined, root: XmlNode | undefined): string | undefined {
  if (!node) return undefined;
  let conventional: string | undefined, undeclared: string | undefined;
  for (const [name, value] of Object.entries(node.attributes)) {
    const colon = name.indexOf(":");
    if (colon < 0 || name.slice(colon + 1) !== "id") continue;
    const prefix = name.slice(0, colon), namespace = namespaceOf(prefix, node, root);
    if (namespace !== undefined && RELATIONSHIP_NAMESPACES.has(namespace)) return value;
    if (prefix === "r") conventional ??= value;
    else if (namespace === undefined) undeclared ??= value;
  }
  return conventional ?? undeclared;
}

/**
 * True when a Relationship element points outside the package. Any attribute with local name "TargetMode" counts,
 * whatever its prefix: treating a prefixed one as external is the safe reading (an external target is never followed).
 */
export function isExternalRelationship(relationship: XmlNode): boolean {
  for (const [name, value] of Object.entries(relationship.attributes)) {
    if ((name === "TargetMode" || name.endsWith(":TargetMode")) && value === "External") return true;
  }
  return false;
}

/** An own property of a lookup table keyed by document text (never an inherited one such as "constructor"). */
export function ownValue<T>(table: Readonly<Record<string, T>>, key: string): T | undefined {
  return Object.hasOwn(table, key) ? table[key] : undefined;
}
