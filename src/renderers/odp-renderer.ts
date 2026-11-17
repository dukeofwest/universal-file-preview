import type { FileFormat } from "../detection/detect-format.js";
import type { ExtendedPresentation } from "../model-extensions.js";
import { parseModel, viewerLocale } from "../parsers/parse-model.js";
import type { DocumentModel } from "../worker/protocol.js";
import { PptxRenderer } from "./ooxml-renderers.js";
import type { FileRenderer, RendererContext, RendererInstance } from "./types.js";
import { renderNotice } from "./xlsx-renderer.js";
import { noticeEntries } from "./workbook-layout.js";

/**
 * Renderer for OpenDocument presentations: .odp, the .otp template and the flat .fodp.
 *
 * Internal module: not exported from the package entry point (decision D2). It composes the
 * PowerPoint renderer rather than extracting anything from it, so PPTX output is untouched: the
 * parsed model is the same Presentation, and the only addition is the notice listing what a
 * presentation contains that the preview cannot show. Nothing is dropped silently.
 */
export class OdpRenderer implements FileRenderer {
  readonly formats: readonly FileFormat[] = ["odp", "otp", "fodp"];
  readonly #slides = new PptxRenderer();

  render(context: RendererContext): RendererInstance {
    const format = context.format === "fodp" ? "fodp" : context.format === "otp" ? "otp" : "odp";
    return this.renderParsed(parseModel(format, context.bytes, context.file.name, { locale: viewerLocale(), signal: context.signal }), context);
  }

  /** Renders a parsed model (from the main thread or the parser worker). Internal. */
  renderParsed(parsed: DocumentModel, context: RendererContext): RendererInstance {
    const instance = this.#slides.renderParsed(parsed, context);
    const notices = noticeEntries((parsed as ExtendedPresentation).diagnostics);
    if (!notices.length) return instance;
    const root = document.createElement("div");
    root.className = "ufp-slides-wrapper";
    root.append(renderNotice(notices), instance.element);
    return { ...instance, element: root };
  }
}
