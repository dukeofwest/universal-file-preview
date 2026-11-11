import { PreviewError } from "../../errors.js";
import { DiagnosticCollector } from "../../diagnostics.js";
import { checkModel, MODEL_LIMITS, ModelBudget } from "../../model-budget.js";
import type { Paragraph, Table, TableCell, TextRun, WordBlock } from "../../model.js";
import type { DocumentSource, ExtendedWordDocument } from "../../model-extensions.js";
import { PackageReader } from "../../package/package-reader.js";
import { nsAttr, nsChildren, parseXmlNs, type NsXmlNode } from "../../xml/ns-tree.js";
import { streamXml, type XmlStartElement } from "../../xml/xml-stream.js";
import type { ExpandedName } from "../../xml/xml-namespaces.js";
import { ODF } from "./namespaces.js";
import { lengthToPoints } from "./units.js";
import { readOdtPackage } from "./odt-package.js";
import { createOdtStyles, listLabel, type OdtListLevel, type OdtStyles } from "./odt-styles.js";
import { readMasterPage } from "./odt-page.js";
import { decodeBase64, resolvePackageImage, sniffImageType, type PackageImage } from "./odt-images.js";

/**
 * OpenDocument text parser (ODF 1.3 Part 3 §5), producing the same WordDocument model the DOCX
 * parser produces so that one renderer serves both.
 *
 * Internal module: not exported from the package entry point (decision D2). The document is read in
 * one streaming pass; each paragraph is captured as a bounded subtree, so only one paragraph is held
 * as a tree at a time. Producers may use any prefix (Part 3 §1.5), so elements are matched by
 * namespace URI and local name, never by a prefixed string.
 *
 * Milestone 3.2 reads the text body: paragraphs, headings, spans and the character and paragraph
 * formatting their styles carry. Lists, tables, images and page layout follow in later milestones,
 * and content this milestone cannot represent is reported through the diagnostics rather than
 * dropped silently.
 */

const { office: OFFICE, text: TEXT, table: TABLE, draw: DRAW, svg: SVG, xlink: XLINK } = ODF;

export type OdtLimits = {
  /** Decoded size of content.xml, or of a whole flat document. Default 64 MiB. */
  maxContentBytes: number;
  /** Decoded size of styles.xml. Default 16 MiB. */
  maxStylesBytes: number;
  /** All decompressed parts together. Default 256 MiB. */
  maxTotalBytes: number;
  /** Blocks (paragraphs, headings and later tables) in the document. Default 100,000. */
  maxBlocks: number;
  /** Elements in one captured paragraph. Default 10,000. */
  maxCaptureElements: number;
  /** Characters in one paragraph's text. Default 1,000,000. */
  maxParagraphTextLength: number;
  /** Table cells across the document. Default 100,000. */
  maxTableCells: number;
};
export type OdtParseOptions = { readonly signal?: AbortSignal; readonly limits?: Partial<OdtLimits> };

const DEFAULT_LIMITS: OdtLimits = {
  maxContentBytes: 64 * 1024 * 1024, maxStylesBytes: 16 * 1024 * 1024, maxTotalBytes: 256 * 1024 * 1024,
  maxBlocks: 100_000, maxCaptureElements: 10_000, maxParagraphTextLength: 1_000_000, maxTableCells: 100_000,
};

/**
 * XML work bounds for one text document, derived from the block budget rather than from any byte
 * size, for the reason recorded in ods.ts: a part's declared and compressed sizes are both tiny next
 * to the work a document can demand. Twenty-four elements per block covers a paragraph with several
 * spans, a heading, or a short table row; the base covers styles and the document scaffolding.
 */
const BASE_ELEMENTS = 250_000, ELEMENTS_PER_BLOCK = 24, TOKENS_PER_ELEMENT = 4;
const MAX_WORK_ELEMENTS = 2 ** 32;

/** Element and token bounds for the content stream, derived from the block budget. */
export function odtWorkLimits(limits: Pick<OdtLimits, "maxBlocks">): { readonly maxElements: number; readonly maxTokens: number } {
  const maxElements = Math.min(BASE_ELEMENTS + limits.maxBlocks * ELEMENTS_PER_BLOCK, MAX_WORK_ELEMENTS);
  return { maxElements, maxTokens: maxElements * TOKENS_PER_ELEMENT };
}

export function resolveOdtLimits(overrides: Partial<OdtLimits> | undefined): OdtLimits {
  const limits = { ...DEFAULT_LIMITS };
  for (const [name, value] of Object.entries(overrides ?? {})) {
    if (!Object.hasOwn(DEFAULT_LIMITS, name)) throw new RangeError(`Unknown ODT limit ${name}.`);
    if (value === undefined) continue;
    if (!Number.isSafeInteger(value) || value < 1) throw new RangeError(`${name} must be a positive safe integer.`);
    (limits as unknown as Record<string, number>)[name] = value;
  }
  return limits;
}

/** Parses an OpenDocument text package (.odt, .ott). */
export function parseOdt(bytes: Uint8Array, name: string, options: OdtParseOptions = {}): ExtendedWordDocument {
  const limits = resolveOdtLimits(options.limits);
  const collector = new DiagnosticCollector();
  const reader = PackageReader.openZip(bytes, {
    partBytes: { "content.xml": limits.maxContentBytes, "styles.xml": limits.maxStylesBytes },
    maxTotalBytes: limits.maxTotalBytes,
  });
  const { format, info } = readOdtPackage(reader, collector);
  const stylesRoot = reader.has("styles.xml")
    ? parsePart(() => parseXmlNs(reader.readText("styles.xml"), { maxSourceLength: limits.maxStylesBytes, ...(options.signal ? { signal: options.signal } : {}) }), "styles.xml")
    : undefined;
  if (stylesRoot && (stylesRoot.uri !== OFFICE || stylesRoot.local !== "document-styles")) {
    throw new PreviewError("styles.xml is not an OpenDocument styles stream.", "invalid-file");
  }
  const image = (href: string) => resolvePackageImage(reader, info, href);
  return streamText({ source: reader.readText("content.xml"), part: "content.xml", mode: "content", format, image, ...(stylesRoot ? { stylesRoot } : {}), ...(info.generator !== undefined ? { generator: info.generator } : {}) }, name, limits, collector, options.signal);
}

export interface TextSource {
  readonly source: string;
  readonly part: string;
  /** "content" reads a package's content.xml; "flat" reads a whole single-file document. */
  readonly mode: "content" | "flat";
  readonly format: "odt" | "ott" | "fodt";
  readonly stylesRoot?: NsXmlNode;
  readonly generator?: string;
  /** Resolves a draw:image reference to bytes inside the package. Absent for a flat document. */
  readonly image?: (href: string) => PackageImage | undefined;
}

/** Reads a text document in one pass (shared by packages and flat files). */
export function streamText(input: TextSource, name: string, limits: OdtLimits, collector: DiagnosticCollector, signal: AbortSignal | undefined): ExtendedWordDocument {
  const reader = new TextReader(input, limits, collector);
  const work = odtWorkLimits(limits);
  parsePart(() => streamXml(input.source, reader, {
    maxSourceLength: limits.maxContentBytes, maxElements: work.maxElements, maxTokens: work.maxTokens,
    maxCaptureElements: limits.maxCaptureElements, emptyInCapture: captureePolicy(input.mode), ...(signal ? { signal } : {}),
  }), input.part);

  // A template is shown as saved: its placeholders are never filled in, so the reader is told.
  if (input.format === "ott" || reader.isTemplate) collector.report("odf.template");
  const surfaces = reader.pageSurfaces();
  const title = name.replace(/\.[^.]*$/, "") || name;
  const source: DocumentSource = {
    format: input.format,
    ...(((input.generator ?? reader.generator)?.split(/[\s$]/)[0]) ? { producer: (input.generator ?? reader.generator)!.split(/[\s$]/)[0] } : {}),
    ...(reader.formatVersion !== undefined ? { formatVersion: reader.formatVersion } : {}),
  };
  const diagnostics = collector.list();
  const model: ExtendedWordDocument = {
    kind: "word", title, blocks: reader.blocks, source, ...surfaces,
    ...(reader.styles?.defaults.paragraph ? { defaultParagraph: reader.styles.defaults.paragraph } : {}),
    ...(reader.styles?.defaults.run ? { defaultRun: reader.styles.defaults.run } : {}),
    ...(diagnostics.length ? { diagnostics } : {}),
  };
  reader.budget.add(model);
  checkModel(model, reader.budget.limits, reader.budget.usage);
  return model;
}

/** Embedded binary data is not kept in captured paragraphs; images are read in milestone 3.4. */
/**
 * What a captured paragraph does not keep. An embedded document is never read as content. Base64
 * image data is dropped from a package, where the image is a package entry anyway, and kept for a
 * flat document, where inline data is the only form an image has.
 */
const captureePolicy = (mode: TextSource["mode"]) => (uri: string, local: string) =>
  uri === OFFICE && (local === "document" || (mode === "content" && local === "binary-data"));

class TextReader {
  readonly blocks: WordBlock[] = [];
  readonly budget = new ModelBudget(MODEL_LIMITS);
  styles: OdtStyles | undefined;
  generator: string | undefined;
  formatVersion: string | undefined;
  /** True when the document declares itself a template, which is shown as saved. */
  isTemplate = false;
  /** Depth of the office:text element, so its children are recognised wherever the body sits. */
  #bodyDepth: number | undefined;
  #unsupported = new Set<string>();
  /** Nesting of text:list elements, and the counter of each level (index 1 is the outermost). */
  /**
   * A flat document carries in one file what a package splits across parts, so its style, automatic
   * style and master style subtrees are captured and then read as if they were styles.xml.
   */
  readonly #flat: { styles?: NsXmlNode; automatic?: NsXmlNode; master?: NsXmlNode } = {};
  /** Cells built so far, across every table: a repeat count must not be able to exhaust memory. */
  #tableCells = 0;
  #listDepth = 0;
  #listLevels: readonly OdtListLevel[] | undefined;
  /** The enclosing lists' levels, so a nested list with its own style does not change the outer one. */
  #listStack: (readonly OdtListLevel[] | undefined)[] = [];
  #counters: number[] = [];
  /** Set by text:list-item so that only an item's first paragraph carries the label. */
  #pendingLabel = false;
  #table: TableBuilder | undefined;

  constructor(private readonly input: TextSource, private readonly limits: OdtLimits, private readonly collector: DiagnosticCollector) {
    // office:automatic-styles is optional, and #readStyles only runs when a style subtree is reached.
    // Build the resolver from styles.xml up front so that a content.xml without automatic styles keeps
    // its common styles, list styles and defaults; #readStyles rebuilds it if a subtree does arrive.
    if (input.mode === "content" && input.stylesRoot !== undefined) {
      this.styles = createOdtStyles(input.stylesRoot, undefined, collector);
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
        if (declared !== undefined && !FLAT_TEXT_TYPES.has(declared)) throw new PreviewError("Not an OpenDocument text document.", "unsupported-format");
        this.isTemplate = declared === TEXT_TEMPLATE_TYPE;
      }
      return;
    }
    // Styles and metadata are captured whole: both are small and are needed before the body is read.
    if (uri === OFFICE && (local === "automatic-styles" || local === "font-face-decls")) return "capture";
    if (this.input.mode === "flat" && uri === OFFICE && (local === "styles" || local === "meta" || local === "master-styles")) return "capture";
    if (uri === OFFICE && local === "text") { this.#bodyDepth = depth; return; }
    if (this.#bodyDepth === undefined) return;

    if (uri === TEXT && (local === "p" || local === "h")) return "capture";
    if (uri === TEXT && local === "list") { this.#enterList(element); return; }
    if (uri === TEXT && local === "list-item") { this.#startItem(true); return; }
    if (uri === TEXT && local === "list-header") { this.#startItem(false); return; }
    if (uri === TABLE && local === "table") { this.#table = new TableBuilder(attr(element, TABLE, "style-name")); return; }
    if (uri === TABLE && local === "table-column") { this.#table?.column(attr(element, TABLE, "style-name"), attr(element, TABLE, "number-columns-repeated"), this.styles); return; }
    if (uri === TABLE && local === "table-row") return this.#table ? "capture" : undefined;
    // A section keeps its paragraphs: only its own framing is not represented.
    if (uri === TEXT && local === "section") return;
    if (uri === TEXT && (local === "table-of-content" || local === "illustration-index" || local === "table-index")) return;
    if (uri === DRAW && (local === "frame" || local === "image" || local === "custom-shape")) return this.#defer("drawing");
    return;
  }

  end(name: ExpandedName, depth: number): void {
    const { uri, local } = name;
    if (uri === OFFICE && local === "text" && this.#bodyDepth === depth) { this.#bodyDepth = undefined; return; }
    if (uri === TEXT && local === "list") { this.#listDepth = Math.max(0, this.#listDepth - 1); this.#listLevels = this.#listStack.pop(); return; }
    if (uri === TABLE && local === "table" && this.#table) {
      const table = this.#table.finish(this.budget, this.styles);
      this.#table = undefined;
      if (table) this.#push({ type: "table", value: table });
    }
  }

  captured(node: NsXmlNode): void {
    const { uri, local } = node;
    if (uri === OFFICE && (local === "automatic-styles" || local === "styles")) { this.#readStyles(node); return; }
    if (uri === OFFICE && local === "master-styles") { this.#flat.master = node; return; }
    if (uri === OFFICE && local === "meta") { this.generator = generatorOf(node); return; }
    if (uri === OFFICE && local === "font-face-decls") return;
    if (uri === TABLE && local === "table-row") { this.#table?.row(node, this); return; }
    if (uri !== TEXT || (local !== "p" && local !== "h")) return;
    this.#push({ type: "paragraph", value: this.paragraphOf(node, local === "h", this.#takeLabel()) });
  }

  #push(block: WordBlock): void {
    if (this.blocks.length >= this.limits.maxBlocks) throw new PreviewError("The document has more blocks than the configured limit.", "resource-limit");
    if (block.type !== "pageBreak") this.budget.record(block.value);
    this.blocks.push(block);
  }

  #readStyles(node: NsXmlNode): void {
    // Both subtrees arrive before the body, in either order; the resolver is rebuilt from whatever
    // has been seen so far, so direct formatting and common styles are both available to it.
    if (node.local === "automatic-styles") this.#flat.automatic = node;
    else this.#flat.styles = node;
    const common = this.#flat.styles ? wrapAsStyles(this.#flat.styles) : this.input.stylesRoot;
    this.styles = createOdtStyles(common, this.#flat.automatic, this.collector);
  }

  /** A flat document's captured subtrees, shaped like styles.xml so one reader serves both forms. */
  #stylesRoot(): NsXmlNode | undefined {
    if (this.input.mode === "content") return this.input.stylesRoot;
    const children = [this.#flat.styles, this.#flat.automatic, this.#flat.master].filter((node): node is NsXmlNode => node !== undefined);
    if (!children.length) return undefined;
    const namespaces = children[0]!.namespaces;
    return { uri: OFFICE, local: "document-styles", qname: "document-styles", attributes: [], children, namespaces };
  }

  #defer(kind: string): "skip" {
    if (!this.#unsupported.has(kind)) {
      this.#unsupported.add(kind);
      this.collector.report(kind === "drawing" ? "odf.unsupported.drawings" : "odf.unsupported.text-content");
    }
    return "skip";
  }

  /**
   * The page size and the header and footer of the document's master page, built once the body has
   * been read so that the same styles apply to them. Real pagination is out of scope, as for DOCX:
   * the header and footer are shown once rather than on every page.
   */
  pageSurfaces(): Partial<Pick<ExtendedWordDocument, "page" | "headers" | "footers" | "firstPageHeaders" | "firstPageFooters" | "differentFirstPageHeaderFooter">> {
    const master = readMasterPage(this.#stylesRoot());
    if (!master) return {};
    const blocks = (node: NsXmlNode | undefined) => node ? this.#surfaceBlocks(node) : undefined;
    const headers = blocks(master.header), footers = blocks(master.footer);
    const firstPageHeaders = blocks(master.headerFirst), firstPageFooters = blocks(master.footerFirst);
    return {
      ...(master.page ? { page: master.page } : {}),
      ...(headers?.length ? { headers } : {}),
      ...(footers?.length ? { footers } : {}),
      ...(firstPageHeaders?.length ? { firstPageHeaders } : {}),
      ...(firstPageFooters?.length ? { firstPageFooters } : {}),
      ...(firstPageHeaders?.length || firstPageFooters?.length ? { differentFirstPageHeaderFooter: true } : {}),
    };
  }

  /** The paragraphs of a header or footer, which are ordinary text content. */
  #surfaceBlocks(node: NsXmlNode): WordBlock[] {
    const blocks: WordBlock[] = [];
    const walk = (parent: NsXmlNode, depth: number): void => {
      if (depth > 8 || blocks.length >= MAX_SURFACE_BLOCKS) return;
      for (const child of parent.children) {
        if (typeof child === "string" || blocks.length >= MAX_SURFACE_BLOCKS) continue;
        if (child.uri === TEXT && (child.local === "p" || child.local === "h")) {
          const paragraph = this.paragraphOf(child, child.local === "h");
          this.budget.record(paragraph);
          blocks.push({ type: "paragraph", value: paragraph });
        } else if (child.uri === TEXT || child.uri === TABLE || child.uri === OFFICE) {
          walk(child, depth + 1);
        }
      }
    };
    walk(node, 0);
    return blocks;
  }

  /**
   * Charges one table cell. Repeat counts multiply, so this is charged while a table is being built
   * rather than once it is finished: a 1 KiB document can otherwise ask for millions of cells.
   */
  chargeTableCell(): void {
    if (++this.#tableCells > this.limits.maxTableCells) throw new PreviewError("The document has more table cells than the configured limit.", "resource-limit");
  }

  /** Reports a kind of content that is kept as text rather than in its own shape. */
  reportFlattened(kind: string): void {
    if (this.#unsupported.has(kind)) return;
    this.#unsupported.add(kind);
    this.collector.report("odf.unsupported.text-content");
  }

  // -------------------------------------------------------------------------------------------
  // Lists
  // -------------------------------------------------------------------------------------------

  #enterList(element: XmlStartElement): void {
    // Pushed for every list, including one past the depth limit, so that entering and leaving stay
    // symmetrical and the enclosing list's levels are always the ones restored.
    this.#listStack.push(this.#listLevels);
    if (this.#listDepth >= 10) { this.#listDepth++; return; }
    this.#listDepth++;
    const styleName = attr(element, TEXT, "style-name");
    const levels = this.styles?.list(styleName);
    if (levels) this.#listLevels = levels;
    const level = this.#listDepth;
    // A new list restarts at its declared start value unless it continues the previous one.
    const continues = attr(element, TEXT, "continue-numbering") === "true" || attr(element, TEXT, "continue-list") !== undefined;
    if (!continues || this.#counters[level] === undefined) {
      this.#counters[level] = (this.#listLevels?.[level]?.startValue ?? 1) - 1;
    }
  }

  #startItem(numbered: boolean): void {
    const level = Math.min(this.#listDepth, 10);
    if (level < 1) return;
    if (numbered) {
      this.#counters[level] = (this.#counters[level] ?? 0) + 1;
      for (let deeper = level + 1; deeper < this.#counters.length; deeper++) this.#counters[deeper] = undefined as unknown as number;
    }
    this.#pendingLabel = numbered;
  }

  /** The list label for the next paragraph, consumed so that later paragraphs of an item have none. */
  #takeLabel(): Paragraph["list"] | undefined {
    if (this.#listDepth < 1) return undefined;
    const level = Math.min(this.#listDepth, 10);
    const levels = this.#listLevels;
    const definition = levels?.[level];
    const indent: Partial<NonNullable<Paragraph["list"]>> = {
      ...(definition?.indentLeft !== undefined ? { indentLeft: definition.indentLeft } : {}),
      ...(definition?.hanging !== undefined ? { hanging: definition.hanging } : {}),
    };
    if (!this.#pendingLabel) {
      // A continuation paragraph of the same item, or a list header: indented, but not labelled.
      return { label: "", level: level - 1, kind: definition?.kind === "ordered" ? "ordered" : "bullet", ...indent };
    }
    this.#pendingLabel = false;
    const label = levels ? listLabel(levels, level, this.#counters) : "•";
    return { label: this.budget.text(label), level: level - 1, kind: definition?.kind === "ordered" ? "ordered" : "bullet", ...indent };
  }

  // -------------------------------------------------------------------------------------------
  // Paragraphs
  // -------------------------------------------------------------------------------------------

  /** Builds one paragraph or heading from its subtree, with an optional list label. */
  paragraphOf(node: NsXmlNode, heading: boolean, list?: Paragraph["list"]): Paragraph {
    const styleName = nsAttr(node, TEXT, "style-name");
    const runs: TextRun[] = [];
    this.#runs(node, this.styles?.run(styleName), runs, 0);
    const paragraph: Paragraph = { runs: runs.length ? runs : [{ text: "" }] };
    const formatting = this.styles?.paragraph(styleName);
    if (formatting) Object.assign(paragraph, formatting);
    if (styleName !== undefined) paragraph.style = this.budget.text(styleName);
    if (list) paragraph.list = list;
    if (heading) {
      paragraph.kind = "heading";
      const declared = Number(nsAttr(node, TEXT, "outline-level"));
      const level = Number.isFinite(declared) ? Math.trunc(declared) : undefined;
      const resolved = level !== undefined && level >= 1 && level <= 10 ? level : this.styles?.outlineLevel(styleName);
      if (resolved !== undefined) paragraph.headingLevel = Math.min(Math.max(resolved, 1), 10);
    }
    for (const run of runs) this.budget.record(run);
    return paragraph;
  }

  /** Collects runs from a paragraph's children, carrying the formatting of enclosing spans. */
  #runs(node: NsXmlNode, inherited: Partial<TextRun> | undefined, runs: TextRun[], depth: number): void {
    if (depth > 32) { this.reportFlattened("nesting"); return; }
    for (const child of node.children) {
      if (typeof child === "string") { this.#push2(runs, child, inherited); continue; }
      const { uri, local } = child;
      if (uri === TEXT && local === "span") {
        const own = this.styles?.run(nsAttr(child, TEXT, "style-name"));
        this.#runs(child, own ? { ...inherited, ...own } : inherited, runs, depth + 1);
      } else if (uri === TEXT && local === "s") {
        const count = Number(nsAttr(child, TEXT, "c") ?? "1");
        const spaces = Number.isFinite(count) ? Math.min(Math.max(Math.trunc(count), 0), 10_000) : 1;
        this.#push2(runs, " ".repeat(spaces), inherited);
      } else if (uri === TEXT && local === "tab") {
        this.#push2(runs, "\t", inherited);
      } else if (uri === TEXT && local === "line-break") {
        // A newline in the run text, which is what the DOCX parser emits for a plain w:br: the page
        // stylesheet sets white-space: pre-wrap, so the line wraps where the break is. The pageBreak
        // field must not be used here; it is reserved for w:br type="page", and splitAtPageBreaks
        // turns such a run into a new page rather than a line break.
        this.#push2(runs, "\n", inherited);
      } else if (uri === TEXT && local === "note") {
        this.reportFlattened("note");
      } else if (uri === OFFICE && local === "annotation") {
        this.reportFlattened("annotation");
      } else if (uri === DRAW && local === "frame") {
        this.#frame(child, inherited, runs);
      } else if (uri === DRAW) {
        this.#defer("drawing");
      } else if (uri === TEXT || uri === OFFICE) {
        // Hyperlinks, bookmarks and fields: the cached text is what the producer saved.
        this.#runs(child, inherited, runs, depth + 1);
      }
    }
  }

  /**
   * An image frame. Only an image stored in this package is shown, at the size the frame declares.
   * A frame holding anything else (an embedded object, a chart, a text box or an unreadable image
   * format) is reported, so a reader is told something is missing instead of seeing a gap.
   */
  #frame(frame: NsXmlNode, inherited: Partial<TextRun> | undefined, runs: TextRun[]): void {
    const width = lengthToPoints(nsAttr(frame, SVG, "width"));
    const height = lengthToPoints(nsAttr(frame, SVG, "height"));
    const alt = surfaceText(frame, SVG, "desc") ?? surfaceText(frame, SVG, "title");
    for (const child of frame.children) {
      if (typeof child === "string" || child.uri !== DRAW || child.local !== "image") continue;
      const href = nsAttr(child, XLINK, "href");
      const image = (href !== undefined ? this.input.image?.(href) : undefined) ?? this.#inlineImage(child);
      if (!image) continue;
      runs.push({
        ...inherited, text: "",
        image: {
          bytes: this.budget.binary(image.bytes), mimeType: image.mimeType,
          ...(alt ? { alt: this.budget.text(alt.slice(0, MAX_ALT_LENGTH)) } : {}),
          ...(width !== undefined && width > 0 ? { width } : {}),
          ...(height !== undefined && height > 0 ? { height } : {}),
        },
      });
      return;
    }
    this.#defer("drawing");
  }

  /**
   * An image a flat document stores inline as base64 (office:binary-data). The media type comes
   * from the bytes themselves, so a document cannot have an image treated as a type it is not.
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

  #push2(runs: TextRun[], text: string, inherited: Partial<TextRun> | undefined): void {
    if (!text) return;
    const last = runs[runs.length - 1];
    if (last && !last.pageBreak && sameFormatting(last, inherited)) { last.text = this.budget.text(last.text + text); return; }
    if (runs.reduce((total, run) => total + run.text.length, 0) + text.length > this.limits.maxParagraphTextLength) {
      throw new PreviewError("A paragraph's text exceeds the configured length limit.", "resource-limit");
    }
    runs.push({ ...inherited, text: this.budget.text(text) });
  }
}

/**
 * Collects one table while its rows stream past. Rows arrive as captured subtrees, so a cell's
 * paragraphs are built from a tree; only one row is held at a time.
 */
class TableBuilder {
  readonly #rows: Array<{ cells: TableCell[] }> = [];
  readonly #columnWidths: number[] = [];

  constructor(private readonly styleName: string | undefined) {}

  column(styleName: string | undefined, repeated: string | undefined, styles: OdtStyles | undefined): void {
    const width = styles?.columnWidth(styleName);
    const count = Number(repeated ?? "1");
    const times = Number.isFinite(count) ? Math.min(Math.max(Math.trunc(count), 1), MAX_TABLE_COLUMNS) : 1;
    for (let index = 0; index < times && this.#columnWidths.length < MAX_TABLE_COLUMNS; index++) this.#columnWidths.push(width ?? 0);
  }

  row(node: NsXmlNode, reader: TextReader): void {
    if (this.#rows.length >= MAX_TABLE_ROWS) return;
    const repeated = Number(nsAttr(node, TABLE, "number-rows-repeated") ?? "1");
    const times = Number.isFinite(repeated) ? Math.min(Math.max(Math.trunc(repeated), 1), MAX_TABLE_ROWS - this.#rows.length) : 1;
    const cells = this.#cells(node, reader);
    for (let index = 0; index < times; index++) {
      if (index > 0) for (const _ of cells) reader.chargeTableCell();
      this.#rows.push({ cells: index === 0 ? cells : cells.map(cell => ({ ...cell, paragraphs: cell.paragraphs.map(paragraph => ({ ...paragraph })) })) });
    }
  }

  #cells(node: NsXmlNode, reader: TextReader): TableCell[] {
    const cells: TableCell[] = [];
    let column = 0;
    for (const child of node.children) {
      if (typeof child === "string" || child.uri !== TABLE) continue;
      if (child.local === "covered-table-cell") { column++; continue; }
      if (child.local !== "table-cell") continue;
      const repeated = Number(nsAttr(child, TABLE, "number-columns-repeated") ?? "1");
      const times = Number.isFinite(repeated) ? Math.min(Math.max(Math.trunc(repeated), 1), MAX_TABLE_COLUMNS) : 1;
      const colSpan = spanOf(nsAttr(child, TABLE, "number-columns-spanned"));
      const rowSpan = spanOf(nsAttr(child, TABLE, "number-rows-spanned"));
      const paragraphs = cellParagraphs(child, reader);
      const style = reader.styles?.cell(nsAttr(child, TABLE, "style-name"));
      for (let index = 0; index < times && cells.length < MAX_TABLE_COLUMNS; index++) {
        reader.chargeTableCell();
        const width = this.#columnWidths[column];
        cells.push({
          paragraphs: index === 0 ? paragraphs : paragraphs.map(paragraph => ({ ...paragraph })),
          column,
          ...(width ? { width } : {}),
          ...(colSpan > 1 ? { colSpan } : {}),
          ...(rowSpan > 1 ? { rowSpan } : {}),
          ...(style?.background ? { background: style.background } : {}),
          ...(style?.verticalAlign ? { verticalAlign: style.verticalAlign } : {}),
          ...(style?.borders ? { borders: style.borders } : {}),
          ...(style?.margins ? { margins: style.margins } : {}),
        });
        column += colSpan;
      }
    }
    return cells;
  }

  finish(budget: ModelBudget, styles: OdtStyles | undefined): Table | undefined {
    if (!this.#rows.length) return undefined;
    const style = styles?.table(this.styleName);
    const table: Table = {
      rows: this.#rows,
      ...(style?.width !== undefined ? { width: style.width } : {}),
      ...(style?.widthPercent !== undefined ? { widthPercent: style.widthPercent } : {}),
      ...(style?.alignment !== undefined ? { alignment: style.alignment } : {}),
    };
    for (const row of this.#rows) { budget.record(row); for (const cell of row.cells) budget.record(cell); }
    return table;
  }
}

const MAX_TABLE_ROWS = 10_000, MAX_TABLE_COLUMNS = 256;
/** Largest number of blocks a header or footer may contribute, so one cannot grow without bound. */
const TEXT_TEMPLATE_TYPE = "application/vnd.oasis.opendocument.text-template";
const FLAT_TEXT_TYPES = new Set(["application/vnd.oasis.opendocument.text", TEXT_TEMPLATE_TYPE]);
const MAX_SURFACE_BLOCKS = 100, MAX_ALT_LENGTH = 1_000;
/** Largest image a flat document may store inline, decoded. Packages are bounded by their byte budget. */
const MAX_INLINE_IMAGE_BYTES = 32 * 1024 * 1024;

/** The text of a child element such as svg:desc, used as an image description. */
function surfaceText(node: NsXmlNode, uri: string, local: string): string | undefined {
  for (const child of node.children) {
    if (typeof child === "string" || child.uri !== uri || child.local !== local) continue;
    const text = child.children.filter((part) => typeof part === "string").join("").trim();
    if (text) return text;
  }
  return undefined;
}
function spanOf(value: string | undefined): number {
  const span = Number(value ?? "1");
  return Number.isFinite(span) ? Math.min(Math.max(Math.trunc(span), 1), MAX_TABLE_COLUMNS) : 1;
}

/**
 * The paragraphs of a table cell. A cell may hold lists and further tables, which the cell model
 * cannot carry, so their paragraphs are kept in order and the loss of their shape is reported.
 */
function cellParagraphs(cell: NsXmlNode, reader: TextReader): Paragraph[] {
  const paragraphs: Paragraph[] = [];
  const walk = (node: NsXmlNode, depth: number, counters: number[], levels: readonly OdtListLevel[] | undefined, listDepth: number): void => {
    if (depth > 16 || paragraphs.length >= MAX_CELL_PARAGRAPHS) return;
    for (const child of node.children) {
      if (typeof child === "string" || paragraphs.length >= MAX_CELL_PARAGRAPHS) continue;
      const { uri, local } = child;
      if (uri === TEXT && (local === "p" || local === "h")) {
        const list = listDepth > 0 && levels
          ? { label: listLabel(levels, Math.min(listDepth, 10), counters), level: Math.min(listDepth, 10) - 1, kind: (levels[Math.min(listDepth, 10)]?.kind ?? "bullet") as "ordered" | "bullet" }
          : undefined;
        paragraphs.push(reader.paragraphOf(child, local === "h", list));
      } else if (uri === TEXT && local === "list") {
        const style = reader.styles?.list(nsAttr(child, TEXT, "style-name")) ?? levels;
        walk(child, depth + 1, counters, style, listDepth + 1);
      } else if (uri === TEXT && (local === "list-item" || local === "list-header")) {
        const level = Math.min(listDepth, 10);
        if (local === "list-item") counters[level] = (counters[level] ?? ((levels?.[level]?.startValue ?? 1) - 1)) + 1;
        walk(child, depth + 1, counters, levels, listDepth);
      } else if (uri === TABLE && local === "table") {
        reader.reportFlattened("nested-table");
        walk(child, depth + 1, counters, levels, listDepth);
      } else if (uri === TABLE || uri === TEXT || uri === OFFICE) {
        walk(child, depth + 1, counters, levels, listDepth);
      }
    }
  };
  walk(cell, 0, [], undefined, 0);
  return paragraphs.length ? paragraphs : [{ runs: [{ text: "" }] }];
}
const MAX_CELL_PARAGRAPHS = 1_000;


/** Two runs may be merged when the formatting carried over from their spans is identical. */
function sameFormatting(run: TextRun, inherited: Partial<TextRun> | undefined): boolean {
  const keys = new Set([...Object.keys(run).filter(key => key !== "text"), ...Object.keys(inherited ?? {})]);
  for (const key of keys) {
    if (key === "text") continue;
    if ((run as Record<string, unknown>)[key] !== (inherited as Record<string, unknown> | undefined)?.[key]) return false;
  }
  return true;
}

/** A flat document's office:styles subtree, shaped like a styles.xml root so one reader serves both. */
function wrapAsStyles(node: NsXmlNode): NsXmlNode {
  // The qname is only carried for diagnostics; it must not be a prefixed literal, because a producer
  // may bind any prefix and this library never compares prefixed names (odf-core guard).
  return { uri: OFFICE, local: "document-styles", qname: "document-styles", attributes: [], children: [node], namespaces: node.namespaces };
}

function generatorOf(meta: NsXmlNode): string | undefined {
  for (const child of nsChildren(meta, ODF.meta, "generator")) {
    const text = child.children.filter((part): part is string => typeof part === "string").join("").trim();
    if (text) return text;
  }
  return undefined;
}

function attr(element: XmlStartElement, uri: string, local: string): string | undefined {
  for (const attribute of element.attributes) if (attribute.uri === uri && attribute.local === local) return attribute.value;
  return undefined;
}

/** Prefixes parse errors with the part they came from, as the spreadsheet parser does. */
function parsePart<T>(read: () => T, part: string): T {
  try { return read(); }
  catch (error) {
    if (error instanceof PreviewError && error.code === "parse-error") throw new PreviewError(`${part}: ${error.message}`, "parse-error", { cause: error });
    throw error;
  }
}

