import { PreviewError } from "../errors.js";
import { NCNAME_CHARS, NCNAME_START_CHARS } from "./xml-tokenizer.js";

/**
 * Namespaces in XML 1.0 (third edition) resolution for the internal XML layer
 * (implementation specification §4.1, milestone M3).
 *
 * Internal module: not exported from the package entry point (decision D2). OpenDocument
 * lets producers choose any prefix (ODF 1.3 Part 3 §1.5), so parsers must compare
 * namespace URIs and local names, never prefixed strings.
 */

export const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
export const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";

/** `uri` is "" for names in no namespace; `prefix` is "" for unprefixed names. */
export interface ExpandedName { readonly uri: string; readonly local: string; readonly prefix: string; }

const ncName = new RegExp(`^[${NCNAME_START_CHARS}][${NCNAME_CHARS}]*$`, "u");
const initialBindings: ReadonlyMap<string, string> = new Map([["xml", XML_NAMESPACE]]);

/**
 * Namespace bindings copied while entering elements, over one document. An element that declares
 * a namespace copies every binding in scope, and nothing else bounded that work: XML nesting (256)
 * and attributes per element (512) together allow over 131,000 bindings in scope, so elements below
 * such a stack cost hundreds of millions of copies while staying far inside every element and token
 * limit. A 560 KiB spreadsheet holding 433 elements cost 17 s and 1.1 GiB that way, and a 3.1 MiB
 * SVG of 431 elements cost 8.3 s and 1.3 GiB, the sanitizer reading images through the same layer
 * (measured; both files stay far inside the element and token limits that do apply).
 *
 * Chaining removes that cost instead of bounding it, so no budget is needed and no document is
 * refused for how it declares namespaces: entering an element costs only its own declarations. A
 * lookup walks the chain and is remembered at each link, so the first lookup of a prefix costs the
 * depth and the rest are one map read. The complete set is assembled only for a caller that needs
 * more than a lookup, which the parsers never do.
 */
class Bindings implements ReadonlyMap<string, string> {
  #resolved: Map<string, string | undefined> | undefined;
  #flat: ReadonlyMap<string, string> | undefined;
  constructor(readonly declarations: ReadonlyMap<string, string>, readonly parent: Bindings | undefined) {}

  /** The URI bound to `prefix`, or undefined when it is unbound or has been undeclared. */
  get(prefix: string): string | undefined {
    const resolved = this.#resolved;
    if (resolved !== undefined) { const hit = resolved.get(prefix); if (hit !== undefined || resolved.has(prefix)) return hit; }
    const declared = this.declarations.get(prefix);
    // An empty value undeclares: xmlns="" leaves unprefixed names in no namespace (Namespaces in XML 1.0, 6.2).
    const value = declared !== undefined ? (declared === "" ? undefined : declared) : this.parent?.get(prefix);
    (this.#resolved ??= new Map()).set(prefix, value);
    return value;
  }

  /** Every binding in scope, flattened once per element that declares one. */
  #flatten(): ReadonlyMap<string, string> {
    if (this.#flat !== undefined) return this.#flat;
    const flat = new Map(this.parent === undefined ? undefined : this.parent.#flatten());
    for (const [prefix, value] of this.declarations) { if (value === "") flat.delete(prefix); else flat.set(prefix, value); }
    return (this.#flat = flat);
  }
  get size(): number { return this.#flatten().size; }
  has(prefix: string): boolean { return this.#flatten().has(prefix); }
  forEach(callback: (value: string, key: string, map: ReadonlyMap<string, string>) => void, thisArg?: unknown): void { this.#flatten().forEach(callback, thisArg); }
  entries() { return this.#flatten().entries(); }
  keys() { return this.#flatten().keys(); }
  values() { return this.#flatten().values(); }
  [Symbol.iterator]() { return this.#flatten()[Symbol.iterator](); }
}

/**
 * A stack of in-scope namespace bindings. Call `enter` with each start tag's attributes and
 * `exit` at its end. Binding maps are immutable and shared: an element without namespace
 * declarations reuses its parent's map.
 *
 * Errors are PreviewError("parse-error"). The optional `offset` arguments only add a
 * character position to the message.
 */
export class NamespaceScope {
  readonly #stack: Bindings[] = [new Bindings(initialBindings, undefined)];

  /** Applies the `xmlns` and `xmlns:p` declarations among `attributes` for a new element. */
  enter(attributes: ReadonlyArray<readonly [string, string]>, offset?: number): void {
    const current = this.#stack[this.#stack.length - 1]!;
    let declared: Map<string, string> | undefined;
    for (const [qname, value] of attributes) {
      if (qname !== "xmlns" && !qname.startsWith("xmlns:")) continue;
      const prefix = qname === "xmlns" ? "" : qname.slice(6);
      if (prefix && !ncName.test(prefix)) fail(`Invalid namespace prefix ${prefix}`, offset);
      if (prefix === "xmlns") fail("The xmlns prefix must not be declared", offset);
      if (prefix === "xml" && value !== XML_NAMESPACE) fail("The xml prefix must not be bound to another namespace", offset);
      if (prefix !== "xml" && value === XML_NAMESPACE) fail("The XML namespace must not be bound to another prefix", offset);
      if (value === XMLNS_NAMESPACE) fail("The xmlns namespace must not be declared", offset);
      if (prefix && value === "") fail(`Namespace prefix ${prefix} must not be undeclared`, offset);
      // An empty value for the default prefix is kept as a declaration: it undeclares in this scope.
      (declared ??= new Map()).set(prefix, value);
    }
    this.#stack.push(declared === undefined ? current : new Bindings(declared, current));
  }

  exit(): void {
    if (this.#stack.length <= 1) throw new RangeError("NamespaceScope.exit() called without a matching enter().");
    this.#stack.pop();
  }

  /** Resolves an element name. Unprefixed names take the default namespace. */
  element(qname: string, offset?: number): ExpandedName {
    const { prefix, local } = split(qname, offset);
    if (prefix === "xmlns") fail("The xmlns prefix must not be used for elements", offset);
    if (!prefix) return { uri: this.lookup("") ?? "", local, prefix };
    return { uri: this.#bound(prefix, offset), local, prefix };
  }

  /** Resolves a non-declaration attribute name. Unprefixed attributes are in no namespace. */
  attribute(qname: string, offset?: number): ExpandedName {
    if (qname === "xmlns" || qname.startsWith("xmlns:")) throw new RangeError("Namespace declarations are not ordinary attributes.");
    const { prefix, local } = split(qname, offset);
    if (!prefix) return { uri: "", local, prefix };
    return { uri: this.#bound(prefix, offset), local, prefix };
  }

  /**
   * Resolves the prefix of QName-valued attribute content such as `of:=SUM([.A1])`. Returns
   * undefined when there is no valid prefix or it is unbound; never throws. `local` is the
   * text after the first colon.
   */
  valuePrefix(value: string): ExpandedName | undefined { return resolveValuePrefix(this.snapshot(), value); }

  lookup(prefix: string): string | undefined { return this.#stack[this.#stack.length - 1]!.get(prefix); }

  /** Immutable in-scope bindings (prefix → URI; the default namespace has prefix ""). */
  snapshot(): ReadonlyMap<string, string> { return this.#stack[this.#stack.length - 1]!; }

  #bound(prefix: string, offset: number | undefined): string {
    const uri = this.lookup(prefix);
    if (uri === undefined) fail(`Unbound namespace prefix ${prefix}`, offset);
    return uri!;
  }
}

/** Shared by NamespaceScope.valuePrefix and nsValuePrefix. */
export function resolveValuePrefix(bindings: ReadonlyMap<string, string>, value: string): ExpandedName | undefined {
  const colon = value.indexOf(":");
  if (colon <= 0) return undefined;
  const prefix = value.slice(0, colon);
  if (!ncName.test(prefix)) return undefined;
  const uri = bindings.get(prefix);
  return uri === undefined ? undefined : { uri, local: value.slice(colon + 1), prefix };
}

function split(qname: string, offset: number | undefined): { prefix: string; local: string } {
  const colon = qname.indexOf(":");
  if (colon < 0) return { prefix: "", local: qname };
  const prefix = qname.slice(0, colon), local = qname.slice(colon + 1);
  if (!ncName.test(prefix) || !ncName.test(local)) fail(`Invalid qualified name ${qname}`, offset);
  return { prefix, local };
}

function fail(message: string, offset: number | undefined): never {
  throw new PreviewError(offset === undefined ? `${message}.` : `${message} (at character ${offset}).`, "parse-error");
}
