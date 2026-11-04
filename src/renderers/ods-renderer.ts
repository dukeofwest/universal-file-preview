import type { FileFormat } from "../detection/detect-format.js";
import { parseOds } from "../parsers/odf/ods.js";
import type { FileRenderer, RendererContext, RendererInstance } from "./types.js";
import { renderWorkbook } from "./xlsx-renderer.js";

/**
 * OpenDocument spreadsheet previews (implementation specification §4.6, milestone M12): the ODS
 * parser feeds the shared workbook grid. Saved display text is shown as written by the producing
 * application; the viewer's locale only affects text this library formats itself.
 */
export class OdsRenderer implements FileRenderer {
  readonly formats: readonly FileFormat[] = ["ods", "ots"];
  render({ bytes, file, signal }: RendererContext): RendererInstance {
    const locale = typeof navigator !== "undefined" && navigator.language ? navigator.language : "en-US";
    return renderWorkbook(parseOds(bytes, file.name, { signal }), { locale });
  }
}
