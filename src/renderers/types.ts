import type { FileFormat } from "../detection/detect-format.js";
import type { PreviewFile } from "../api/types.js";

export interface RendererContext { readonly file: PreviewFile; readonly format: FileFormat; readonly bytes: Uint8Array; readonly signal: AbortSignal; }
export interface RendererInstance {
  readonly element: HTMLElement;
  setZoom?(zoom: number): void;
  fitToWidth?(): void;
  destroy?(): void;
}
export interface FileRenderer { readonly formats: readonly FileFormat[]; render(context: RendererContext): Promise<RendererInstance> | RendererInstance; }
