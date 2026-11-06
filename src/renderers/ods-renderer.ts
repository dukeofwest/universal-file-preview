import type { FileFormat } from "../detection/detect-format.js";
import { PreviewError } from "../errors.js";
import { parseModel, viewerLocale } from "../parsers/parse-model.js";
import type { DocumentModel } from "../worker/protocol.js";
import type { FileRenderer, RendererContext, RendererInstance } from "./types.js";
import { renderWorkbook } from "./xlsx-renderer.js";

/**
 * OpenDocument spreadsheet previews (implementation specification §4.6, milestones M12 and M15):
 * the ODS parser (or the flat .fods parser) feeds the shared workbook grid. Saved display text is
 * shown as written by the producing application; the viewer's locale only affects text this
 * library formats itself.
 */
export class OdsRenderer implements FileRenderer {
  readonly formats: readonly FileFormat[] = ["ods", "ots", "fods"];
  render(context: RendererContext): RendererInstance {
    const format = context.format === "fods" || context.format === "ots" ? context.format : "ods";
    return this.renderParsed(parseModel(format, context.bytes, context.file.name, { locale: viewerLocale(), signal: context.signal }), context);
  }
  /** Renders a parsed model (from the main thread or the parser worker). Internal. */
  renderParsed(parsed: DocumentModel, _context: RendererContext): RendererInstance {
    if (parsed.kind !== "workbook") throw new PreviewError("The document model does not match the renderer.", "renderer-error");
    return renderWorkbook(parsed, { locale: viewerLocale() });
  }
}
