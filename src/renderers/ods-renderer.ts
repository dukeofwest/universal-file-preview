import type { FileFormat } from "../detection/detect-format.js";
import { parseFods } from "../parsers/odf/flat.js";
import { parseOds } from "../parsers/odf/ods.js";
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
  render({ bytes, file, format, signal }: RendererContext): RendererInstance {
    const locale = typeof navigator !== "undefined" && navigator.language ? navigator.language : "en-US";
    const workbook = format === "fods" ? parseFods(bytes, file.name, { signal }) : parseOds(bytes, file.name, { signal });
    return renderWorkbook(workbook, { locale });
  }
}
