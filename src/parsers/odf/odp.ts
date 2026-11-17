import { PreviewError } from "../../errors.js";
import { DiagnosticCollector } from "../../diagnostics.js";
import { checkModel, MODEL_LIMITS, ModelBudget } from "../../model-budget.js";
import type { Paragraph, PptxElement, PptxFill, PptxImageElement, PptxLine, PptxLineElement, PptxShapeElement, PptxTableElement, PptxTextElement, TextRun } from "../../model.js";
import type { DocumentSource, ExtendedPresentation } from "../../model-extensions.js";
import { PackageReader } from "../../package/package-reader.js";
import { nsAttr, nsChildren, nsFirst, parseXmlNs, type NsXmlNode } from "../../xml/ns-tree.js";
import { streamXml, type XmlStartElement } from "../../xml/xml-stream.js";
import type { ExpandedName } from "../../xml/xml-namespaces.js";
import { ODF } from "./namespaces.js";
import { lengthToPoints, odfColor } from "./units.js";
import { readOdpPackage } from "./odp-package.js";
import { createOdpStyles, type OdpStyleIndex } from "./odp-styles.js";
import { createOdtStyles, listLabel, type OdtListLevel, type OdtStyles } from "./odt-styles.js";
import { decodeBase64, resolvePackageImage, sniffImageType, type PackageImage } from "./odt-images.js";

/**
 * OpenDocument presentation parser (ODF 1.3 Part 3 §9 and §10), producing the same Presentation
 * model the PPTX parser produces so that one renderer serves both.
 *
 * Internal module: not exported from the package entry point (decision D2). The document is read in
 * one streaming pass; each drawing object is captured as a bounded subtree, so only one object is
 * held as a tree at a time. Producers may use any prefix (Part 3 §1.5), so elements are matched by
 * namespace URI and local name, never by a prefixed string.
 *
 * Milestone 4.2 reads the slides and their text: one slide per draw:page, and a text box becomes the
 * model's text element with its paragraphs, carrying the character and paragraph formatting their
 * styles hold. Images, backgrounds, shapes and tables follow in later milestones, and content this
 * milestone cannot represent is reported through the diagnostics rather than dropped silently.
 *
 * Geometry is kept in EMU, as the PPTX parser produces and the renderer expects: a slide's declared
 * size is the coordinate space its objects are placed in.
 */

const { office: OFFICE, text: TEXT, table: TABLE, draw: DRAW, presentation: PRESENTATION, style: STYLE, svg: SVG, fo: FO, xlink: XLINK } = ODF;

/** English Metric Units per point, the unit the presentation model and its renderer use. */
const EMU_PER_POINT = 12_700;
/** A slide may not declare a surface larger than this, in points, so geometry cannot run away. */
const MAX_SLIDE_POINTS = 20_000;
/** 10 x 7.5 inches, the ODF default when a presentation declares no page layout. */
const DEFAULT_SLIDE = { width: 720, height: 540 } as const;
/** Longest alternative text, or slide title, kept for an object. */
const MAX_ALT_LENGTH = 1000;
/**
 * The 0.1 inch text insets producers write, in EMU. A graphic style may state its own padding;
 * graphic properties are read in milestone 4.4, where fills and lines need them too.
 */
const STANDARD_INSETS = { top: 91_440, right: 91_440, bottom: 91_440, left: 91_440 } as const;
/** Largest image a flat document may store inline, as for text documents. */
const MAX_INLINE_IMAGE_BYTES = 32 * 1024 * 1024;

export type OdpLimits = {
  /** Decoded size of content.xml, or of a whole flat document. Default 64 MiB. */
  maxContentBytes: number;
  /** Decoded size of styles.xml. Default 16 MiB. */
  maxStylesBytes: number;
  /** All decompressed parts together. Default 256 MiB. */
  maxTotalBytes: number;
  /** Slides in the presentation. Default 1,000. */
  maxSlides: number;
  /** Objects kept on one slide. Default 500. */
  maxElementsPerSlide: number;
  /** Paragraphs kept in one text box. Default 1,000. */
  maxParagraphsPerFrame: number;
  /** Elements in one captured drawing object. Default 10,000. */
  maxCaptureElements: number;
  /** Characters in one paragraph's text. Default 1,000,000. */
  maxParagraphTextLength: number;
};
export type OdpParseOptions = { readonly signal?: AbortSignal; readonly limits?: Partial<OdpLimits> };

const DEFAULT_LIMITS: OdpLimits = {
  maxContentBytes: 64 * 1024 * 1024, maxStylesBytes: 16 * 1024 * 1024, maxTotalBytes: 256 * 1024 * 1024,
  maxSlides: 1_000, maxElementsPerSlide: 500, maxParagraphsPerFrame: 1_000,
  maxCaptureElements: 10_000, maxParagraphTextLength: 1_000_000,
};

/**
 * XML work bounds for one presentation, derived from the slide and object budget rather than from any
 * byte size, for the reason recorded in ods.ts: a part's declared and compressed sizes are both tiny
 * next to the work a document can demand. Twelve elements per permitted object covers a frame with a
 * text box of several paragraphs and spans; the base covers styles and the document scaffolding.
 */
const BASE_ELEMENTS = 250_000, ELEMENTS_PER_OBJECT = 12, TOKENS_PER_ELEMENT = 4;
const MAX_WORK_ELEMENTS = 2 ** 32;

/** Element and token bounds for the content stream, derived from the slide and object budgets. */
export function odpWorkLimits(limits: Pick<OdpLimits, "maxSlides" | "maxElementsPerSlide">): { readonly maxElements: number; readonly maxTokens: number } {
  const objects = limits.maxSlides * limits.maxElementsPerSlide;
  const maxElements = Math.min(BASE_ELEMENTS + objects * ELEMENTS_PER_OBJECT, MAX_WORK_ELEMENTS);
  return { maxElements, maxTokens: maxElements * TOKENS_PER_ELEMENT };
}

export function resolveOdpLimits(overrides: Partial<OdpLimits> | undefined): OdpLimits {
  const limits = { ...DEFAULT_LIMITS };
  for (const [name, value] of Object.entries(overrides ?? {})) {
    if (!Object.hasOwn(DEFAULT_LIMITS, name)) throw new RangeError(`Unknown ODP limit ${name}.`);
    if (value === undefined) continue;
    if (!Number.isSafeInteger(value) || value < 1) throw new RangeError(`${name} must be a positive safe integer.`);
    (limits as unknown as Record<string, number>)[name] = value;
  }
  return limits;
}

/** Parses an OpenDocument presentation package (.odp, .otp). */
export function parseOdp(bytes: Uint8Array, name: string, options: OdpParseOptions = {}): ExtendedPresentation {
  const limits = resolveOdpLimits(options.limits);
  const collector = new DiagnosticCollector();
  const reader = PackageReader.openZip(bytes, {
    partBytes: { "content.xml": limits.maxContentBytes, "styles.xml": limits.maxStylesBytes },
    maxTotalBytes: limits.maxTotalBytes,
  });
  const { format, info } = readOdpPackage(reader, collector);
  const stylesRoot = reader.has("styles.xml")
    ? parsePart(() => parseXmlNs(reader.readText("styles.xml"), { maxSourceLength: limits.maxStylesBytes, ...(options.signal ? { signal: options.signal } : {}) }), "styles.xml")
    : undefined;
  if (stylesRoot && (stylesRoot.uri !== OFFICE || stylesRoot.local !== "document-styles")) {
    throw new PreviewError("styles.xml is not an OpenDocument styles stream.", "invalid-file");
  }
  const image = (href: string) => resolvePackageImage(reader, info, href);
  return streamPresentation({ source: reader.readText("content.xml"), part: "content.xml", mode: "content", format, image, ...(stylesRoot ? { stylesRoot } : {}), ...(info.generator !== undefined ? { generator: info.generator } : {}) }, name, limits, collector, options.signal);
}

export interface PresentationSource {
  readonly source: string;
  readonly part: string;
  /** "content" reads a package's content.xml; "flat" reads a whole single-file document. */
  readonly mode: "content" | "flat";
  readonly format: "odp" | "otp" | "fodp";
  readonly stylesRoot?: NsXmlNode;
  readonly generator?: string;
  /** Resolves a draw:image reference to bytes inside the package. Absent for a flat document. */
  readonly image?: (href: string) => PackageImage | undefined;
}

/** Reads a presentation in one pass (shared by packages and, from milestone 4.5, flat documents). */
export function streamPresentation(input: PresentationSource, name: string, limits: OdpLimits, collector: DiagnosticCollector, signal: AbortSignal | undefined): ExtendedPresentation {
  const reader = new SlideReader(input, limits, collector);
  const work = odpWorkLimits(limits);
  parsePart(() => streamXml(input.source, reader, {
    maxSourceLength: limits.maxContentBytes, maxElements: work.maxElements, maxTokens: work.maxTokens,
    maxCaptureElements: limits.maxCaptureElements, emptyInCapture: emptyPolicy(input.mode), ...(signal ? { signal } : {}),
  }), input.part);
  reader.finish();

  // A template is shown as saved: its placeholders are never filled in, so the reader is told.
  if (input.format === "otp" || reader.isTemplate) collector.report("odf.template");
  const title = name.replace(/\.[^.]*$/, "") || name;
  const source: DocumentSource = {
    format: input.format,
    ...(((input.generator ?? reader.generator)?.split(/[\s$]/)[0]) ? { producer: (input.generator ?? reader.generator)!.split(/[\s$]/)[0] } : {}),
    ...(reader.formatVersion !== undefined ? { formatVersion: reader.formatVersion } : {}),
  };
  const diagnostics = collector.list();
  // A presentation with no slide at all still renders: the model's renderer always shows one surface.
  const slides = reader.slides.length ? reader.slides : [{ title, paragraphs: [], ...reader.slideSize() }];
  const model: ExtendedPresentation = {
    kind: "presentation", title, slides, source,
    ...(diagnostics.length ? { diagnostics } : {}),
  };
  reader.budget.add(model);
  checkModel(model, reader.budget.limits, reader.budget.usage);
  return model;
}

/**
 * What a captured drawing object does not keep: an embedded document is never read as content, and
 * base64 image data is dropped from a package, where the image is a package entry anyway. Inline
 * data is kept for a flat document, where it is the only form an image has (milestone 4.3).
 */
const emptyPolicy = (mode: PresentationSource["mode"]) => (uri: string, local: string) =>
  uri === OFFICE && (local === "document" || (mode === "content" && local === "binary-data"));

type Slide = ExtendedPresentation["slides"][number];

class SlideReader {
  readonly slides: Slide[] = [];
  readonly budget = new ModelBudget(MODEL_LIMITS);
  styles: OdtStyles | undefined;
  /** Graphic and drawing-page properties, which the text resolver does not expose. */
  graphics: OdpStyleIndex | undefined;
  generator: string | undefined;
  formatVersion: string | undefined;
  /** True when the document declares itself a template, which is shown as saved. */
  isTemplate = false;
  /** Depth of the office:presentation element, so its children are recognised wherever the body sits. */
  #bodyDepth: number | undefined;
  /** Depth of the draw:page being read, and the objects collected for it so far. */
  #pageDepth: number | undefined;
  #page: { name?: string; title?: string; style?: string; elements: PptxElement[]; dropped: number } | undefined;
  /** Depth of a presentation:notes subtree, whose content belongs to the notes page, not the slide. */
  #notesDepth: number | undefined;
  #unsupported = new Set<string>();
  #slidesDropped = 0;
  /** A flat document carries its style subtrees in the same file; captured and read as styles.xml. */
  readonly #flat: { styles?: NsXmlNode; automatic?: NsXmlNode; master?: NsXmlNode } = {};
  /** Characters pushed into the paragraph being built, so the length limit also bounds merged runs. */
  #paragraphText = 0;

  constructor(private readonly input: PresentationSource, private readonly limits: OdpLimits, private readonly collector: DiagnosticCollector) {
    // Build the resolver from styles.xml up front so that a content.xml without automatic styles keeps
    // its common styles, list styles and defaults; #readStyles rebuilds it if a subtree does arrive.
    if (input.mode === "content" && input.stylesRoot !== undefined) {
      this.styles = createOdtStyles(input.stylesRoot, undefined, collector);
      this.graphics = createOdpStyles(input.stylesRoot, undefined);
    }
  }

  start(element: XmlStartElement, depth: number): void | "capture" | "skip" {
    const { uri, local } = element.name;
    if (depth === 1) {
      if (this.input.mode === "flat") {
        if (uri !== OFFICE || local !== "document") throw new PreviewError("The file is not a flat OpenDocument document.", "invalid-file");
      } else if (uri !== OFFICE || local !== "document-content") {
        throw new PreviewError("content.xml is not an OpenDocument content stream.", "invalid-file");
      }
      this.formatVersion = attr(element, OFFICE, "version");
      // A flat document declares its kind on the root element; a package declares it in its mimetype.
      if (this.input.mode === "flat") {
        const declared = attr(element, OFFICE, "mimetype");
        if (declared !== undefined && !FLAT_PRESENTATION_TYPES.has(declared)) throw new PreviewError("Not an OpenDocument presentation.", "unsupported-format");
        this.isTemplate = declared === PRESENTATION_TEMPLATE_TYPE;
      }
      return;
    }
    // Styles and metadata are captured whole: both are small and are needed before the body is read.
    if (uri === OFFICE && (local === "automatic-styles" || local === "font-face-decls")) return "capture";
    if (this.input.mode === "flat" && uri === OFFICE && (local === "styles" || local === "meta" || local === "master-styles")) return "capture";
    if (uri === OFFICE && local === "presentation") { this.#bodyDepth = depth; return; }
    if (this.#bodyDepth === undefined) return;

    // Speaker notes sit inside a page but belong to its notes surface, which is not shown.
    if (uri === PRESENTATION && local === "notes") { this.#notesDepth ??= depth; this.collector.report("odf.unsupported.notes"); return "skip"; }
    if (this.#notesDepth !== undefined) return "skip";

    if (uri === DRAW && local === "page") { this.#startPage(element, depth); return; }
    if (this.#page === undefined) {
      // presentation:settings and the like sit beside the pages and carry no shown content.
      return;
    }
    // Every drawing object is captured as a bounded subtree, so only one is held as a tree at a time.
    if (uri === DRAW && DRAWING_OBJECTS.has(local)) return "capture";
    if (uri === PRESENTATION && local === "animations") return "skip";
    return;
  }

  end(name: ExpandedName, depth: number): void {
    const { uri, local } = name;
    if (uri === PRESENTATION && local === "notes" && this.#notesDepth === depth) { this.#notesDepth = undefined; return; }
    if (uri === OFFICE && local === "presentation" && this.#bodyDepth === depth) { this.#bodyDepth = undefined; return; }
    if (uri === DRAW && local === "page" && this.#pageDepth === depth) this.#finishPage();
  }

  captured(node: NsXmlNode): void {
    const { uri, local } = node;
    if (uri === OFFICE && (local === "automatic-styles" || local === "styles")) { this.#readStyles(node); return; }
    if (uri === OFFICE && local === "master-styles") { this.#flat.master = node; return; }
    if (uri === OFFICE && local === "meta") { this.generator = generatorOf(node); return; }
    if (uri === OFFICE && local === "font-face-decls") return;
    if (uri !== DRAW || !DRAWING_OBJECTS.has(local) || !this.#page) return;
    if (this.#page.elements.length >= this.limits.maxElementsPerSlide) { this.#page.dropped++; return; }
    for (const element of this.#object(node, local)) {
      if (this.#page.elements.length >= this.limits.maxElementsPerSlide) { this.#page.dropped++; continue; }
      this.budget.record(element);
      this.#page.elements.push(element);
    }
  }

  #startPage(element: XmlStartElement, depth: number): void {
    this.#pageDepth = depth;
    if (this.slides.length >= this.limits.maxSlides) { this.#slidesDropped++; this.#page = undefined; return; }
    const name = attr(element, DRAW, "name");
    const style = attr(element, DRAW, "style-name");
    this.#page = { ...(name !== undefined ? { name } : {}), ...(style !== undefined ? { style } : {}), elements: [], dropped: 0 };
  }

  #finishPage(): void {
    this.#pageDepth = undefined;
    const page = this.#page;
    this.#page = undefined;
    if (!page) return;
    if (page.dropped) this.collector.report("odf.bounds.slide-elements", { count: page.dropped });
    // The slide's title is the text of its title placeholder, which is what a reader sees as the title.
    const slide: Slide = {
      title: this.budget.text(page.title ?? page.name ?? ""),
      paragraphs: [],
      ...this.slideSize(),
      ...(this.#background(page.style) ?? {}),
      ...(page.elements.length ? { elements: page.elements } : {}),
    };
    this.slides.push(slide);
  }

  /**
   * A frame: in this milestone its text box becomes a text element. A frame holding anything else (an
   * image, an embedded object, a chart or a plugin) is reported, so a reader is told something is
   * missing rather than seeing a gap; images arrive in milestone 4.3.
   */
  #frame(frame: NsXmlNode): PptxElement | undefined {
    const box = this.#geometry(frame);
    const picture = this.#image(frame, box);
    if (picture) return picture;
    const table = nsFirst(frame, TABLE, "table");
    if (table) return this.#table(table, box);
    const textBox = nsFirst(frame, DRAW, "text-box");
    if (!textBox) return this.#defer(frameContent(frame) ?? "frame");
    const paragraphs = this.#paragraphsOf(textBox);
    if (!paragraphs.length) return undefined;
    // The title placeholder's text is the slide's title, which is what a reader sees as its name.
    if (this.#page && !this.#page.title && nsAttr(frame, PRESENTATION, "class") === "title") {
      this.#page.title = this.budget.text(paragraphs.map(paragraph => paragraph.runs.map(run => run.text).join("")).join(" ").trim().slice(0, MAX_ALT_LENGTH));
    }
    const element: PptxTextElement = { kind: "text", ...box, margins: STANDARD_INSETS, paragraphs };
    return element;
  }

  /**
   * The paragraphs of a text box, including the items of any list it holds. The label rules are the
   * text reader's: an item's first paragraph carries the label, later paragraphs of the same item are
   * indented without one, and a list header never has one.
   */
  #paragraphsOf(textBox: NsXmlNode): Paragraph[] {
    const paragraphs: Paragraph[] = [];
    /** Item number per level; index 1 is the outermost list, as in the text reader. */
    const counters: number[] = [];
    let full = false, pendingLabel = false;
    const walk = (parent: NsXmlNode, depth: number, levels: readonly OdtListLevel[] | undefined, listDepth: number): void => {
      for (const child of parent.children) {
        if (typeof child === "string" || full) continue;
        const { uri, local } = child;
        if (uri === TEXT && (local === "p" || local === "h")) {
          if (paragraphs.length >= this.limits.maxParagraphsPerFrame) { full = true; this.#defer("frame-paragraphs"); return; }
          // A list may take its numbering from the paragraph style rather than from its own name.
          const own = listDepth < 1 ? undefined : levels ?? this.styles?.listOfParagraph(nsAttr(child, TEXT, "style-name"));
          const paragraph = this.#paragraphOf(child, local === "h");
          if (listDepth > 0) {
            const level = Math.min(listDepth, 10);
            const definition = own?.[level];
            paragraph.list = {
              label: pendingLabel ? this.budget.text(own ? listLabel(own, level, counters) : "•") : "",
              level: level - 1,
              kind: definition?.kind === "ordered" ? "ordered" : "bullet",
            };
            pendingLabel = false;
          }
          paragraphs.push(paragraph);
        } else if (uri === TEXT && local === "list") {
          const own = this.styles?.list(nsAttr(child, TEXT, "style-name")) ?? levels;
          const level = Math.min(listDepth + 1, 10);
          // A new list restarts at its declared start value unless it continues the previous one.
          const continues = nsAttr(child, TEXT, "continue-numbering") === "true" || nsAttr(child, TEXT, "continue-list") !== undefined;
          if (!continues || counters[level] === undefined) counters[level] = (own?.[level]?.startValue ?? 1) - 1;
          if (depth < 16) walk(child, depth + 1, own, listDepth + 1);
        } else if (uri === TEXT && (local === "list-item" || local === "list-header")) {
          const level = Math.min(Math.max(listDepth, 1), 10);
          if (local === "list-item") {
            counters[level] = (counters[level] ?? 0) + 1;
            for (let deeper = level + 1; deeper < counters.length; deeper++) counters[deeper] = undefined as unknown as number;
          }
          pendingLabel = local === "list-item";
          if (depth < 16) walk(child, depth + 1, levels, listDepth);
        } else if (depth < 16 && (uri === TEXT || uri === OFFICE)) {
          walk(child, depth + 1, levels, listDepth);
        }
      }
    };
    walk(textBox, 0, undefined, 0);
    return paragraphs;
  }

  /** Builds one paragraph or heading from its subtree. */
  #paragraphOf(node: NsXmlNode, heading: boolean): Paragraph {
    const styleName = nsAttr(node, TEXT, "style-name");
    const runs: TextRun[] = [];
    this.#paragraphText = 0;
    this.#runs(node, this.styles?.run(styleName), runs, 0);
    const paragraph: Paragraph = { runs: runs.length ? runs : [{ text: "" }] };
    const formatting = this.styles?.paragraph(styleName);
    if (formatting) Object.assign(paragraph, formatting);
    if (styleName !== undefined) paragraph.style = this.budget.text(styleName);
    if (heading) paragraph.kind = "heading";
    for (const run of runs) this.budget.record(run);
    return paragraph;
  }

  /** Collects runs from a paragraph's children, carrying the formatting of enclosing spans. */
  #runs(node: NsXmlNode, inherited: Partial<TextRun> | undefined, runs: TextRun[], depth: number): void {
    if (depth > 32) { this.#defer("nesting"); return; }
    for (const child of node.children) {
      if (typeof child === "string") { this.#pushRun(runs, child, inherited); continue; }
      const { uri, local } = child;
      if (uri === TEXT && local === "span") {
        const own = this.styles?.run(nsAttr(child, TEXT, "style-name"));
        this.#runs(child, own ? { ...inherited, ...own } : inherited, runs, depth + 1);
      } else if (uri === TEXT && local === "s") {
        const count = Number(nsAttr(child, TEXT, "c") ?? "1");
        const spaces = Number.isFinite(count) ? Math.min(Math.max(Math.trunc(count), 0), 10_000) : 1;
        this.#pushRun(runs, " ".repeat(spaces), inherited);
      } else if (uri === TEXT && local === "tab") {
        this.#pushRun(runs, "\t", inherited);
      } else if (uri === TEXT && local === "line-break") {
        // A newline in the run text, as the DOCX parser emits for a plain w:br: the slide stylesheet
        // sets white-space: pre-wrap, so the line wraps where the break is.
        this.#pushRun(runs, "\n", inherited);
      } else if (uri === DRAW) {
        this.#defer("drawing");
      } else if (uri === TEXT || uri === OFFICE) {
        // Hyperlinks, bookmarks, fields and date placeholders: the cached text is what was saved.
        this.#runs(child, inherited, runs, depth + 1);
      }
    }
  }

  #pushRun(runs: TextRun[], text: string, inherited: Partial<TextRun> | undefined): void {
    if (!text) return;
    // Charged before the run is placed, so text appended to the previous run counts too.
    if (this.#paragraphText + text.length > this.limits.maxParagraphTextLength) {
      throw new PreviewError("A paragraph's text exceeds the configured length limit.", "resource-limit");
    }
    this.#paragraphText += text.length;
    const last = runs[runs.length - 1];
    if (last && sameFormatting(last, inherited)) { last.text = this.budget.text(last.text + text); return; }
    runs.push({ ...inherited, text: this.budget.text(text) });
  }

  /** A frame's position and size in EMU, bounded so a declared length cannot run away. */
  #geometry(frame: NsXmlNode): { left: number; top: number; width: number; height: number } {
    const emu = (value: string | undefined, fallback: number) => {
      const points = lengthToPoints(value);
      if (points === undefined || !Number.isFinite(points)) return fallback;
      return Math.round(Math.min(Math.max(points, -MAX_SLIDE_POINTS), MAX_SLIDE_POINTS) * EMU_PER_POINT);
    };
    return {
      left: emu(nsAttr(frame, SVG, "x"), 0),
      top: emu(nsAttr(frame, SVG, "y"), 0),
      width: emu(nsAttr(frame, SVG, "width"), 0),
      height: emu(nsAttr(frame, SVG, "height"), 0),
    };
  }

  /**
   * One drawing object, as the elements it becomes. A shape that carries text becomes two: the shape
   * and a text element over it, which is how the model represents a shape with a caption.
   */
  #object(node: NsXmlNode, local: string): PptxElement[] {
    if (local === "frame") return list(this.#frame(node));
    if (local === "line") return list(this.#line(node));
    if (local === "rect" || local === "ellipse" || local === "circle" || local === "custom-shape") {
      const geometry = local === "custom-shape" ? this.#presetOf(node) : local === "rect" ? "rect" : "ellipse";
      return [...list(this.#shape(node, geometry)), ...list(this.#shapeText(node))];
    }
    this.#defer(local);
    return [];
  }

  /** A basic or custom shape: its outline, fill and stroke. */
  #shape(node: NsXmlNode, geometry: string): PptxShapeElement {
    const style = styleRef(node);
    const fill = this.#graphicFill(style);
    const line = this.#graphicLine(style);
    return {
      kind: "shape", ...this.#geometry(node), geometry,
      ...(fill ? { fill } : {}), ...(line ? { line } : {}),
    };
  }

  /** The text a shape carries, as a text element over the shape's own box. */
  #shapeText(node: NsXmlNode): PptxTextElement | undefined {
    const paragraphs = this.#paragraphsOf(node);
    if (!paragraphs.length) return undefined;
    return { kind: "text", ...this.#geometry(node), margins: STANDARD_INSETS, paragraphs };
  }

  /** A straight line between the two points it declares. */
  #line(node: NsXmlNode): PptxLineElement | undefined {
    const x1 = this.#emu(nsAttr(node, SVG, "x1")), y1 = this.#emu(nsAttr(node, SVG, "y1"));
    const x2 = this.#emu(nsAttr(node, SVG, "x2")), y2 = this.#emu(nsAttr(node, SVG, "y2"));
    if (x1 === undefined || y1 === undefined || x2 === undefined || y2 === undefined) return this.#defer("line");
    const line = this.#graphicLine(styleRef(node));
    return {
      kind: "line",
      left: Math.min(x1, x2), top: Math.min(y1, y2),
      width: Math.abs(x2 - x1), height: Math.abs(y2 - y1),
      ...(line ? { line } : {}),
      // The model flips a box rather than storing the direction, as the PowerPoint parser does.
      ...(x2 < x1 ? { flipH: true } : {}), ...(y2 < y1 ? { flipV: true } : {}),
    };
  }

  /**
   * A table inside a frame. Its columns share the frame's width, since a presentation table states
   * column widths in its own styles rather than in the frame; rows share its height. Content past the
   * caps is reported, never dropped in silence.
   */
  #table(table: NsXmlNode, box: { left: number; top: number; width: number; height: number }): PptxTableElement | undefined {
    const rows: PptxTableElement["rows"] = [];
    let columns = 0, truncated = false;
    const rowNodes: NsXmlNode[] = [];
    const collectRows = (parent: NsXmlNode, depth: number) => {
      for (const child of parent.children) {
        if (typeof child === "string" || child.uri !== TABLE) continue;
        if (child.local === "table-row") rowNodes.push(child);
        // table:table-header-rows and table:table-row-group hold rows of their own.
        else if (depth < 4 && (child.local === "table-header-rows" || child.local === "table-row-group")) collectRows(child, depth + 1);
      }
    };
    collectRows(table, 0);
    for (const rowNode of rowNodes) {
      if (rows.length >= MAX_SLIDE_TABLE_ROWS) { truncated = true; break; }
      const cells: PptxTableElement["rows"][number]["cells"] = [];
      for (const cellNode of nsChildren(rowNode, TABLE, "table-cell")) {
        if (cells.length >= MAX_SLIDE_TABLE_COLUMNS) { truncated = true; break; }
        cells.push({ paragraphs: this.#paragraphsOf(cellNode), margins: STANDARD_INSETS });
      }
      if (!cells.length) continue;
      columns = Math.max(columns, cells.length);
      rows.push({ height: 0, cells });
    }
    if (truncated) this.#defer("table-size");
    if (!rows.length) return this.#defer("table");
    const columnWidth = Math.round(box.width / Math.max(columns, 1));
    const rowHeight = Math.round(box.height / rows.length);
    for (const row of rows) row.height = rowHeight;
    return { kind: "table", ...box, columns: Array.from({ length: Math.max(columns, 1) }, () => columnWidth), rows };
  }

  /** The preset outline a custom shape declares, mapped onto the geometry the renderer draws. */
  #presetOf(node: NsXmlNode): string {
    const enhanced = nsFirst(node, DRAW, "enhanced-geometry");
    const type = nsAttr(enhanced, DRAW, "type");
    if (type === undefined || type === "non-primitive") {
      // An outline given only as an equation path is drawn as a rectangle, which is stated, not hidden.
      if (type !== undefined) this.collector.report("odf.approximated.shapes");
      return "rect";
    }
    const preset = Object.hasOwn(SHAPE_PRESETS, type) ? SHAPE_PRESETS[type] : undefined;
    if (preset === undefined) {
      this.collector.report("odf.approximated.shapes");
      return "rect";
    }
    return preset;
  }

  /** A solid fill from a graphic style; any other kind of fill is reported rather than approximated. */
  #graphicFill(style: { family: string; name: string } | undefined): PptxFill | undefined {
    if (!style) return undefined;
    const fill = this.graphics?.attr(style.family, style.name, DRAW, "fill");
    if (fill === undefined || fill === "none") return undefined;
    if (fill !== "solid") { this.#defer(`fill-${fill}`); return undefined; }
    const color = odfColor(this.graphics?.attr(style.family, style.name, DRAW, "fill-color"));
    if (!color) return undefined;
    const opacity = percentage(this.graphics?.attr(style.family, style.name, DRAW, "opacity"));
    return { color, ...(opacity !== undefined && opacity < 1 ? { opacity } : {}) };
  }

  /** The stroke a graphic style declares, in EMU, or undefined when the style draws none. */
  #graphicLine(style: { family: string; name: string } | undefined): PptxLine | undefined {
    if (!style) return undefined;
    const stroke = this.graphics?.attr(style.family, style.name, DRAW, "stroke");
    if (stroke === undefined || stroke === "none") return undefined;
    const color = odfColor(this.graphics?.attr(style.family, style.name, SVG, "stroke-color")) ?? "#000000";
    const points = lengthToPoints(this.graphics?.attr(style.family, style.name, SVG, "stroke-width"));
    const width = points !== undefined && points > 0 ? Math.round(Math.min(points, MAX_SLIDE_POINTS) * EMU_PER_POINT) : HAIRLINE_EMU;
    const opacity = percentage(this.graphics?.attr(style.family, style.name, SVG, "stroke-opacity"));
    return {
      color, width,
      ...(stroke === "dash" ? { style: "dash" } : {}),
      ...(opacity !== undefined && opacity < 1 ? { opacity } : {}),
    };
  }

  /** One length in EMU, bounded; undefined when the document states none. */
  #emu(value: string | undefined): number | undefined {
    const points = lengthToPoints(value);
    if (points === undefined || !Number.isFinite(points)) return undefined;
    return Math.round(Math.min(Math.max(points, -MAX_SLIDE_POINTS), MAX_SLIDE_POINTS) * EMU_PER_POINT);
  }

  /**
   * An image frame. Only an image stored in this package, or stored inline in a flat document, is
   * shown; a frame holding anything else is reported by the caller. The media type comes from the
   * bytes themselves, so a document cannot have an image treated as a type it is not.
   */
  #image(frame: NsXmlNode, box: { left: number; top: number; width: number; height: number }): PptxImageElement | undefined {
    for (const child of frame.children) {
      if (typeof child === "string" || child.uri !== DRAW || child.local !== "image") continue;
      const href = nsAttr(child, XLINK, "href");
      const image = (href !== undefined ? this.input.image?.(href) : undefined) ?? this.#inlineImage(child);
      if (!image) {
        // The reference is there but the image cannot be read: say so rather than leave a gap.
        this.#defer("image");
        return undefined;
      }
      const alt = surfaceText(frame, SVG, "desc") ?? surfaceText(frame, SVG, "title");
      return {
        kind: "image", ...box,
        image: {
          bytes: this.budget.binary(image.bytes), mimeType: image.mimeType,
          ...(alt ? { alt: this.budget.text(alt.slice(0, MAX_ALT_LENGTH)) } : {}),
        },
      };
    }
    return undefined;
  }

  /**
   * An image a flat document stores inline as base64 (office:binary-data). The media type comes from
   * the bytes themselves, as for a package entry.
   */
  #inlineImage(image: NsXmlNode): PackageImage | undefined {
    for (const child of image.children) {
      if (typeof child === "string" || child.uri !== OFFICE || child.local !== "binary-data") continue;
      const encoded = child.children.filter((part): part is string => typeof part === "string").join("");
      const bytes = decodeBase64(encoded, MAX_INLINE_IMAGE_BYTES);
      const mimeType = bytes ? sniffImageType(bytes) : undefined;
      if (bytes && mimeType) return { bytes, mimeType };
    }
    return undefined;
  }

  /**
   * The slide background from its drawing-page style. A solid colour is shown; a gradient, bitmap or
   * hatch fill is reported instead of being approximated, so nothing on the slide is invented.
   */
  #background(styleName: string | undefined): { background: PptxFill } | undefined {
    const fill = this.graphics?.attr("drawing-page", styleName, DRAW, "fill");
    if (fill === undefined || fill === "none") return undefined;
    if (fill === "solid") {
      const color = odfColor(this.graphics?.attr("drawing-page", styleName, DRAW, "fill-color"));
      if (!color) return undefined;
      const opacity = percentage(this.graphics?.attr("drawing-page", styleName, DRAW, "opacity"));
      return { background: { color, ...(opacity !== undefined && opacity < 1 ? { opacity } : {}) } };
    }
    this.#defer(`fill-${fill}`);
    return undefined;
  }

  #readStyles(node: NsXmlNode): void {
    if (node.local === "automatic-styles") this.#flat.automatic = node;
    else this.#flat.styles = node;
    const common = this.#flat.styles ? wrapAsStyles(this.#flat.styles) : this.input.stylesRoot;
    this.styles = createOdtStyles(common, this.#flat.automatic, this.collector);
    this.graphics = createOdpStyles(common, this.#flat.automatic);
  }

  /** A flat document's captured subtrees, shaped like styles.xml so one reader serves both forms. */
  #stylesRoot(): NsXmlNode | undefined {
    if (this.input.mode === "content") return this.input.stylesRoot;
    const children = [this.#flat.styles, this.#flat.automatic, this.#flat.master].filter((node): node is NsXmlNode => node !== undefined);
    if (!children.length) return undefined;
    return { uri: OFFICE, local: "document-styles", qname: "document-styles", attributes: [], children, namespaces: children[0]!.namespaces };
  }

  /**
   * The slide surface, in EMU: the size of the first page layout the styles declare, which is what a
   * presentation with one slide size writes. Anything not stated falls back to the ODF default.
   */
  slideSize(): { width: number; height: number } {
    const layout = firstPageLayout(this.#stylesRoot());
    const side = (name: string, fallback: number) => {
      const points = lengthToPoints(nsAttr(layout, FO, name));
      return points !== undefined && points > 0 ? Math.min(points, MAX_SLIDE_POINTS) : fallback;
    };
    return {
      width: Math.round(side("page-width", DEFAULT_SLIDE.width) * EMU_PER_POINT),
      height: Math.round(side("page-height", DEFAULT_SLIDE.height) * EMU_PER_POINT),
    };
  }

  #defer(kind: string): undefined {
    // Each kind is counted once, which is what the shared message says.
    if (!this.#unsupported.has(kind)) {
      this.#unsupported.add(kind);
      this.collector.report("odf.unsupported.slide-content");
    }
    return undefined;
  }

  /** Reported once the stream has been read, so the count covers the whole document. */
  finish(): void {
    if (this.#slidesDropped) this.collector.report("odf.bounds.slides", { count: this.#slidesDropped });
  }
}

/** A hairline stroke, for a style that draws one without stating a width. */
const HAIRLINE_EMU = Math.round(0.75 * EMU_PER_POINT);
/** Caps for a table on a slide. Content past them is reported, as everywhere else. */
const MAX_SLIDE_TABLE_ROWS = 200, MAX_SLIDE_TABLE_COLUMNS = 32;

/**
 * ODF enhanced-geometry types mapped onto the preset outlines the renderer draws (ODF 1.3 Part 3
 * §19.145). A type not listed here is drawn as a rectangle and reported as approximated, so a reader
 * is told the outline is not the saved one rather than being shown a wrong shape silently.
 */
const SHAPE_PRESETS: Readonly<Record<string, string>> = {
  rectangle: "rect", "round-rectangle": "roundRect", ellipse: "ellipse", circle: "ellipse",
  "isosceles-triangle": "triangle", "right-triangle": "rtTriangle", diamond: "diamond",
  trapezoid: "trapezoid", parallelogram: "parallelogram", pentagon: "pentagon", hexagon: "hexagon",
  octagon: "octagon", cross: "plus", star5: "star5", star4: "star4", star6: "star6", star8: "star8",
  "right-arrow": "rightArrow", "left-arrow": "leftArrow", "up-arrow": "upArrow", "down-arrow": "downArrow",
  "left-right-arrow": "leftRightArrow", "up-down-arrow": "upDownArrow", ring: "donut", can: "can",
  cube: "cube", "block-arc": "blockArc", arc: "arc", pie: "pie", "circle-pie": "pie", heart: "heart",
  smiley: "smileyFace", sun: "sun", moon: "moon", cloud: "cloud", "forbidden": "noSmoking",
  "rectangular-callout": "wedgeRectCallout", "round-rectangular-callout": "wedgeRoundRectCallout",
  "round-callout": "wedgeEllipseCallout", "line-callout-1": "borderCallout1",
  "flowchart-process": "flowChartProcess", "flowchart-decision": "flowChartDecision",
  "flowchart-connector": "flowChartConnector", "flowchart-alternate-process": "flowChartAlternateProcess",
  "quad-bevel": "bevel", "paper": "foldedCorner", "vertical-scroll": "verticalScroll", "horizontal-scroll": "horizontalScroll",
};

/** The style a drawing object refers to: a graphic style, or a presentation placeholder's style. */
function styleRef(node: NsXmlNode): { family: string; name: string } | undefined {
  const graphic = nsAttr(node, DRAW, "style-name");
  if (graphic !== undefined) return { family: "graphic", name: graphic };
  const placeholder = nsAttr(node, PRESENTATION, "style-name");
  if (placeholder !== undefined) return { family: "presentation", name: placeholder };
  return undefined;
}

/** One element, or none: the object readers return undefined for anything they cannot show. */
const list = <T>(element: T | undefined): T[] => (element === undefined ? [] : [element]);

/** Drawing objects a slide may hold. Shapes, lines, frames and tables are read; the rest are reported. */
const DRAWING_OBJECTS = new Set(["frame", "custom-shape", "rect", "ellipse", "circle", "line", "polyline", "polygon", "path", "g", "connector", "measure", "caption", "page-thumbnail"]);
const PRESENTATION_TEMPLATE_TYPE = "application/vnd.oasis.opendocument.presentation-template";
const FLAT_PRESENTATION_TYPES = new Set(["application/vnd.oasis.opendocument.presentation", PRESENTATION_TEMPLATE_TYPE]);

/** What a frame holds, for the diagnostic: the first child this library recognises by name. */
function frameContent(frame: NsXmlNode): string | undefined {
  for (const child of frame.children) {
    if (typeof child === "string") continue;
    if (child.uri === DRAW && child.local === "image") return "image";
    if (child.uri === DRAW && (child.local === "object" || child.local === "object-ole")) return "object";
    if (child.uri === DRAW && child.local === "plugin") return "plugin";
    if (child.uri === TABLE && child.local === "table") return "table";
  }
  return undefined;
}

/** A percentage attribute such as draw:opacity as a fraction; undefined when it is not one. */
function percentage(value: string | undefined): number | undefined {
  const match = /^(\d+(?:\.\d+)?)%$/.exec(value?.trim() ?? "");
  if (!match) return undefined;
  return Math.min(Math.max(Number(match[1]) / 100, 0), 1);
}

/** The text of a child element, e.g. svg:desc, trimmed; undefined when it has none. */
function surfaceText(node: NsXmlNode, uri: string, local: string): string | undefined {
  const child = nsFirst(node, uri, local);
  if (!child) return undefined;
  const text = child.children.filter((part): part is string => typeof part === "string").join("").trim();
  return text || undefined;
}

/** The first style:page-layout-properties in a styles stream: a presentation's slide size. */
function firstPageLayout(stylesRoot: NsXmlNode | undefined): NsXmlNode | undefined {
  if (!stylesRoot) return undefined;
  for (const parent of [nsFirst(stylesRoot, OFFICE, "automatic-styles"), nsFirst(stylesRoot, OFFICE, "styles")]) {
    if (!parent) continue;
    for (const layout of nsChildren(parent, STYLE, "page-layout")) {
      const properties = nsFirst(layout, STYLE, "page-layout-properties");
      if (properties) return properties;
    }
  }
  return undefined;
}

/** True when `run` carries exactly the inherited formatting, so text can be appended to it. */
function sameFormatting(run: TextRun, inherited: Partial<TextRun> | undefined): boolean {
  const own = Object.keys(run).filter(key => key !== "text");
  const other = Object.keys(inherited ?? {}).filter(key => key !== "text");
  if (own.length !== other.length) return false;
  return own.every(key => (run as Record<string, unknown>)[key] === (inherited as Record<string, unknown> | undefined)?.[key]);
}

/** A flat document's office:styles subtree, wrapped so the styles reader sees a styles stream. */
function wrapAsStyles(node: NsXmlNode): NsXmlNode {
  return { uri: OFFICE, local: "document-styles", qname: "document-styles", attributes: [], children: [node], namespaces: node.namespaces };
}

function generatorOf(meta: NsXmlNode): string | undefined {
  const inner = nsFirst(meta, OFFICE, "meta") ?? meta;
  const generator = nsFirst(inner, ODF.meta, "generator");
  if (!generator) return undefined;
  const text = generator.children.filter((child): child is string => typeof child === "string").join("").trim();
  return text || undefined;
}

function attr(element: XmlStartElement, uri: string, local: string): string | undefined {
  return element.attributes.find(candidate => candidate.uri === uri && candidate.local === local)?.value;
}

/** Names the part in any error a part's own reader throws, as the text and spreadsheet parsers do. */
function parsePart<T>(read: () => T, part: string): T {
  try {
    return read();
  } catch (error) {
    if (error instanceof PreviewError) throw new PreviewError(`${part}: ${error.message}`, error.code, { cause: error });
    throw error;
  }
}
