import type { XmlNode } from "../xml/xml-parser.js";

/**
 * Namespace-aware reading of relationship references in Office Open XML parts. Internal module.
 *
 * The XML parser used for OOXML keeps qualified names without resolving namespaces. These helpers resolve the
 * prefix where the document declares it, on the element itself or on the part's root element (where Office
 * declares it), so an attribute in an unrelated namespace is not taken for a relationship reference.
 */

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
    const prefix = name.slice(0, colon), declaration = `xmlns:${prefix}`;
    const namespace = Object.hasOwn(node.attributes, declaration) ? node.attributes[declaration] : root && Object.hasOwn(root.attributes, declaration) ? root.attributes[declaration] : undefined;
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
