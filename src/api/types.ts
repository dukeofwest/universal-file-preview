import type { RendererRegistry } from "../renderers/registry.js";

export interface PreviewFile {
  readonly name: string;
  readonly source: string | Blob | ArrayBuffer | Uint8Array;
  readonly mimeType?: string;
}
export interface PreviewOptions {
  readonly container: HTMLElement;
  readonly file: PreviewFile;
  readonly toolbar?: boolean;
  readonly zoom?: boolean;
  readonly download?: boolean;
  readonly maxFileSize?: number;
  readonly registry?: RendererRegistry;
}
