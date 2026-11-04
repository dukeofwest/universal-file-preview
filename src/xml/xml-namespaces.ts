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
 * A stack of in-scope namespace bindings. Call `enter` with each start tag's attributes and
 * `exit` at its end. Binding maps are immutable and shared: an element without namespace
 * declarations reuses its parent's map.
 *
 * Errors are PreviewError("parse-error"). The optional `offset` arguments only add a
 * character position to the message.
 */
export class NamespaceScope {
  readonly #stack: Array<ReadonlyMap<string, string>> = [initialBindings];

  /** Applies the `xmlns` and `xmlns:p` declarations among `attributes` for a new element. */
  enter(attributes: ReadonlyArray<readonly [string, string]>, offset?: number): void {
    const current = this.#stack[this.#stack.length - 1]!;
    let next: Map<string, string> | undefined;
    for (const [qname, value] of attributes) {
      if (qname !== "xmlns" && !qname.startsWith("xmlns:")) continue;
      const prefix = qname === "xmlns" ? "" : qname.slice(6);
      if (prefix && !ncName.test(prefix)) fail(`Invalid namespace prefix ${prefix}`, offset);
      if (prefix === "xmlns") fail("The xmlns prefix must not be declared", offset);
      if (prefix === "xml" && value !== XML_NAMESPACE) fail("The xml prefix must not be bound to another namespace", offset);
      if (prefix !== "xml" && value === XML_NAMESPACE) fail("The XML namespace must not be bound to another prefix", offset);
      if (value === XMLNS_NAMESPACE) fail("The xmlns namespace must not be declared", offset);
      if (prefix && value === "") fail(`Namespace prefix ${prefix} must not be undeclared`, offset);
      next ??= new Map(current);
      if (prefix === "" && value === "") next.delete("");
      else next.set(prefix, value);
    }
    this.#stack.push(next ?? current);
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
