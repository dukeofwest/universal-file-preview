import type { WordPageLayout } from "../../model.js";
import { nsAttr, nsChildren, nsFirst, type NsXmlNode } from "../../xml/ns-tree.js";
import { ODF } from "./namespaces.js";
import { lengthToPoints } from "./units.js";

/**
 * Page layout, headers and footers for OpenDocument text documents (ODF 1.3 Part 3 §16.9 to §16.11).
 *
 * Internal module: not exported from the package entry point (decision D2). A document's page
 * surfaces come from styles.xml: office:master-styles holds the master pages, each naming a
 * style:page-layout that carries the size and margins. Real pagination is out of scope, as it is for
 * DOCX: the layout gives the page its size, and the header and footer are shown once.
 */

const { style: STYLE, office: OFFICE, fo: FO } = ODF;

export interface OdtMasterPage {
  readonly header?: NsXmlNode;
  readonly footer?: NsXmlNode;
  readonly headerFirst?: NsXmlNode;
  readonly footerFirst?: NsXmlNode;
  readonly page?: WordPageLayout;
}

/**
 * The master page a document uses. `name` selects one by name, as a paragraph style's
 * style:master-page-name does; without it the first master page is used, which is what producers
 * write for a document with one page style.
 */
export function readMasterPage(stylesRoot: NsXmlNode | undefined, name?: string): OdtMasterPage | undefined {
  if (!stylesRoot) return undefined;
  const masters = nsFirst(stylesRoot, OFFICE, "master-styles");
  if (!masters) return undefined;
  const pages = [...nsChildren(masters, STYLE, "master-page")];
  const master = (name !== undefined ? pages.find(page => nsAttr(page, STYLE, "name") === name) : undefined) ?? pages[0];
  if (!master) return undefined;
  const layoutName = nsAttr(master, STYLE, "page-layout-name");
  const page = readPageLayout(stylesRoot, layoutName);
  const header = nsFirst(master, STYLE, "header");
  const footer = nsFirst(master, STYLE, "footer");
  const headerFirst = nsFirst(master, STYLE, "header-first");
  const footerFirst = nsFirst(master, STYLE, "footer-first");
  return {
    ...(header ? { header } : {}),
    ...(footer ? { footer } : {}),
    ...(headerFirst ? { headerFirst } : {}),
    ...(footerFirst ? { footerFirst } : {}),
    ...(page ? { page } : {}),
  };
}

/** A4 portrait in points, used for any dimension the document does not state. */
const DEFAULT_PAGE = { width: 595.28, height: 841.89, margin: 56.69 } as const;
/** Largest page a document may declare, so a layout cannot ask for an unbounded surface. */
const MAX_PAGE_POINTS = 20_000;

/** The size and margins of a style:page-layout, falling back to A4 for anything not stated. */
export function readPageLayout(stylesRoot: NsXmlNode, name: string | undefined): WordPageLayout | undefined {
  if (name === undefined) return undefined;
  const automatic = nsFirst(stylesRoot, OFFICE, "automatic-styles");
  const common = nsFirst(stylesRoot, OFFICE, "styles");
  const layout = [automatic, common].reduce<NsXmlNode | undefined>((found, node) => {
    if (found || !node) return found;
    return [...nsChildren(node, STYLE, "page-layout")].find(entry => nsAttr(entry, STYLE, "name") === name);
  }, undefined);
  if (!layout) return undefined;
  const properties = nsFirst(layout, STYLE, "page-layout-properties");
  if (!properties) return undefined;
  const landscape = nsAttr(properties, STYLE, "print-orientation") === "landscape";
  const width = bounded(lengthToPoints(nsAttr(properties, FO, "page-width"))) ?? (landscape ? DEFAULT_PAGE.height : DEFAULT_PAGE.width);
  const height = bounded(lengthToPoints(nsAttr(properties, FO, "page-height"))) ?? (landscape ? DEFAULT_PAGE.width : DEFAULT_PAGE.height);
  const all = bounded(lengthToPoints(nsAttr(properties, FO, "margin")));
  const margin = (side: string) => bounded(lengthToPoints(nsAttr(properties, FO, `margin-${side}`))) ?? all ?? DEFAULT_PAGE.margin;
  const headerDistance = distanceOf(nsFirst(properties, STYLE, "header-style"));
  const footerDistance = distanceOf(nsFirst(properties, STYLE, "footer-style"));
  return {
    width, height,
    marginTop: margin("top"), marginRight: margin("right"), marginBottom: margin("bottom"), marginLeft: margin("left"),
    ...(headerDistance !== undefined ? { headerDistance } : {}),
    ...(footerDistance !== undefined ? { footerDistance } : {}),
  };
}

/** The gap a header or footer style keeps from the body text. */
function distanceOf(style: NsXmlNode | undefined): number | undefined {
  if (!style) return undefined;
  const properties = nsFirst(style, STYLE, "header-footer-properties");
  if (!properties) return undefined;
  return bounded(lengthToPoints(nsAttr(properties, FO, "margin-bottom"))) ?? bounded(lengthToPoints(nsAttr(properties, FO, "margin-top")));
}

function bounded(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value) || value < 0) return undefined;
  return Math.min(value, MAX_PAGE_POINTS);
}
