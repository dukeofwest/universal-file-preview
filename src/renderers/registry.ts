import type { FileFormat } from "../detection/detect-format.js";
import type { FileRenderer } from "./types.js";
import { PreviewError } from "../errors.js";
import { ImageRenderer } from "./image-renderer.js";
import { DocxRenderer, PptxRenderer, XlsxRenderer } from "./ooxml-renderers.js";

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
  return registry;
}
