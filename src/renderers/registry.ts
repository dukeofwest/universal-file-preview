import type { FileFormat } from "../detection/detect-format.js";
import type { FileRenderer } from "./types.js";
import { PreviewError } from "../errors.js";
import { ImageRenderer } from "./image-renderer.js";
import { DocxRenderer, PptxRenderer, XlsxRenderer } from "./ooxml-renderers.js";
import { OdsRenderer } from "./ods-renderer.js";
import type { RendererContext, RendererInstance } from "./types.js";
import type { DocumentModel } from "../worker/protocol.js";

/** A built-in renderer, which can render a model parsed elsewhere (by the parser worker). Internal. */
export interface ParsedModelRenderer extends FileRenderer { renderParsed(model: DocumentModel, context: RendererContext): RendererInstance }
/**
 * True only for the library's own document renderers. Renderers an application registers always parse and
 * render through render(), on the main thread, even if they happen to have a renderParsed method.
 */
export function isBuiltInDocumentRenderer(renderer: FileRenderer): renderer is ParsedModelRenderer {
  return renderer instanceof DocxRenderer || renderer instanceof PptxRenderer || renderer instanceof XlsxRenderer || renderer instanceof OdsRenderer;
}

export class RendererRegistry {
  readonly #renderers = new Map<FileFormat, FileRenderer>();
  register(renderer: FileRenderer): () => void {
    const formats = new Set<FileFormat>();
    for (const format of renderer.formats) {
      if (formats.has(format)) throw new PreviewError(`Renderer lists ${format} more than once.`, "renderer-error");
      formats.add(format);
      if (this.#renderers.has(format)) throw new PreviewError(`A renderer is already registered for ${format}.`, "renderer-error");
    }
    for (const format of formats) {
      this.#renderers.set(format, renderer);
    }
    return () => { for (const format of renderer.formats) if (this.#renderers.get(format) === renderer) this.#renderers.delete(format); };
  }
  get(format: FileFormat): FileRenderer | undefined { return this.#renderers.get(format); }
}

export function createDefaultRendererRegistry(): RendererRegistry {
  const registry = new RendererRegistry();
  registry.register(new ImageRenderer()); registry.register(new DocxRenderer()); registry.register(new XlsxRenderer()); registry.register(new PptxRenderer());
  registry.register(new OdsRenderer());
  return registry;
}
