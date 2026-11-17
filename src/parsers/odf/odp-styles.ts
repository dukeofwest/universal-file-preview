import { nsAttr, nsChildren, nsFirst, type NsXmlNode } from "../../xml/ns-tree.js";
import { ODF } from "./namespaces.js";

/**
 * Graphic and drawing-page style lookup for OpenDocument presentations (ODF 1.3 Part 3 §16.30 and
 * §17.6 to §17.19).
 *
 * Internal module: not exported from the package entry point (decision D2). The text style resolver
 * (odt-styles.ts) resolves paragraph and character formatting, which presentations share; what it
 * does not expose is the graphic and drawing-page property sets, which only drawings use. This index
 * keeps those: it finds a style by family and name in either the automatic or the common styles and
 * reads one attribute from it, following style:parent-style-name so that an attribute a parent sets
 * is found. Nothing is merged or copied: a lookup walks the chain and stops at the first answer.
 */

const { style: STYLE, draw: DRAW, presentation: PRESENTATION, office: OFFICE } = ODF;

/** Property element per style family, by the family's own name. */
const PROPERTY_ELEMENTS: Readonly<Record<string, string>> = {
  graphic: "graphic-properties",
  presentation: "graphic-properties",
  "drawing-page": "drawing-page-properties",
};

/** Longest parent chain followed, as in the text resolver: a cycle must not be able to hang a parse. */
const MAX_PARENTS = 64;

export interface OdpStyleIndex {
  /**
   * One attribute of a style's property element, following the parent chain. `family` is the ODF
   * style family ("graphic", "presentation" or "drawing-page").
   */
  attr(family: string, name: string | undefined, uri: string, local: string): string | undefined;
  /** A style's property element itself, for a caller that reads several attributes from it. */
  properties(family: string, name: string | undefined): NsXmlNode | undefined;
}

export function createOdpStyles(stylesRoot: NsXmlNode | undefined, automaticStyles: NsXmlNode | undefined): OdpStyleIndex {
  const byKey = new Map<string, NsXmlNode>();
  const add = (parent: NsXmlNode | undefined) => {
    if (!parent) return;
    for (const style of nsChildren(parent, STYLE, "style")) {
      const family = nsAttr(style, STYLE, "family"), name = nsAttr(style, STYLE, "name");
      if (family === undefined || name === undefined) continue;
      const key = `${family}\u0000${name}`;
      // The first definition wins, as in the text resolver: automatic styles are added first, and a
      // document that names one style twice must not depend on which copy is read.
      if (!byKey.has(key)) byKey.set(key, style);
    }
  };
  // Automatic styles first: a drawing's own style is automatic and takes precedence over a common one.
  add(automaticStyles);
  add(nsFirst(stylesRoot, OFFICE, "automatic-styles"));
  add(nsFirst(stylesRoot, OFFICE, "styles"));

  const properties = (family: string, name: string | undefined): NsXmlNode | undefined => {
    const element = PROPERTY_ELEMENTS[family];
    if (name === undefined || element === undefined) return undefined;
    let style = byKey.get(`${family}\u0000${name}`);
    for (let step = 0; style && step < MAX_PARENTS; step++) {
      const own = nsFirst(style, STYLE, element);
      if (own) return own;
      const parent = nsAttr(style, STYLE, "parent-style-name");
      style = parent !== undefined ? byKey.get(`${family}\u0000${parent}`) : undefined;
    }
    return undefined;
  };

  return {
    properties,
    attr(family, name, uri, local) {
      const element = PROPERTY_ELEMENTS[family];
      if (name === undefined || element === undefined) return undefined;
      let style = byKey.get(`${family}\u0000${name}`);
      for (let step = 0; style && step < MAX_PARENTS; step++) {
        const value = nsAttr(nsFirst(style, STYLE, element), uri, local);
        if (value !== undefined) return value;
        const parent = nsAttr(style, STYLE, "parent-style-name");
        style = parent !== undefined ? byKey.get(`${family}\u0000${parent}`) : undefined;
      }
      return undefined;
    },
  };
}

/** The families a slide's own objects use, so a caller need not spell them out. */
export const GRAPHIC_FAMILIES = ["graphic", "presentation"] as const;
export { DRAW as DRAW_NS, PRESENTATION as PRESENTATION_NS };
