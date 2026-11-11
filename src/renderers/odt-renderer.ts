import type { FileFormat } from "../detection/detect-format.js";
import type { ExtendedWordDocument } from "../model-extensions.js";
import { parseModel, viewerLocale } from "../parsers/parse-model.js";
import type { DocumentModel } from "../worker/protocol.js";
import { DocxRenderer } from "./ooxml-renderers.js";
import type { FileRenderer, RendererContext, RendererInstance } from "./types.js";
import { renderNotice } from "./xlsx-renderer.js";
import { noticeEntries } from "./workbook-layout.js";

/**
 * Renderer for OpenDocument text documents: .odt, the .ott template and the flat .fodt.
 *
 * Internal module: not exported from the package entry point (decision D2). It composes the Word
 * renderer rather than extracting anything from it, so DOCX output is untouched: the parsed model is
 * the same WordDocument, and the only addition is the notice listing what a document contains that
 * the preview cannot show. Nothing is dropped silently.
 */
export class OdtRenderer implements FileRenderer {
  readonly formats: readonly FileFormat[] = ["odt", "ott", "fodt"];
  readonly #word = new DocxRenderer();

  render(context: RendererContext): RendererInstance {
    const format = context.format === "fodt" ? "fodt" : context.format === "ott" ? "ott" : "odt";
    return this.renderParsed(parseModel(format, context.bytes, context.file.name, { locale: viewerLocale() }), context);
  }

  /** Renders a parsed model (from the main thread or the parser worker). Internal. */
  renderParsed(parsed: DocumentModel, context: RendererContext): RendererInstance {
    const instance = this.#word.renderParsed(parsed, context);
    const notices = noticeEntries((parsed as ExtendedWordDocument).diagnostics);
    if (!notices.length) return instance;
    const root = document.createElement("div");
    root.className = "ufp-word-wrapper";
    root.append(renderNotice(notices), instance.element);
    return { ...instance, element: root };
  }
}
