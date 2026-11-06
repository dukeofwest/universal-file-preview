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
  /**
   * Nonce for the viewer's fallback `<style>` element. The viewer adopts a constructable stylesheet where the
   * browser supports it, which needs no nonce and no `'unsafe-inline'`; only browsers without constructable
   * stylesheets get a `<style>` element, which a Content Security Policy without `'unsafe-inline'` allows only
   * with this nonce.
   */
  readonly styleNonce?: string;
}
