/**
 * Small fixed SVG graphics (spreadsheet icons and buttons) built element by element with createElementNS and
 * setAttribute, never parsed from markup: the library has no HTML or markup sink, which keeps it compatible with
 * Trusted Types (require-trusted-types-for 'script'). Internal module.
 */
const SVG_NAMESPACE = "http://www.w3.org/2000/svg";

/** An element: tag name, attributes in order, and child elements. */
export type SvgNode = readonly [tag: string, attributes: Readonly<Record<string, string | number>>, children?: readonly SvgNode[]];

export function svgElement([tag, attributes, children = []]: SvgNode, document: Document = globalThis.document): SVGElement {
  const element = document.createElementNS(SVG_NAMESPACE, tag) as SVGElement;
  for (const [name, value] of Object.entries(attributes)) element.setAttribute(name, String(value));
  for (const child of children) element.append(svgElement(child, document));
  return element;
}
