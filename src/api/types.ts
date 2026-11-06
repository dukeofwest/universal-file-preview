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
  /**
   * Parse DOCX, XLSX, PPTX, ODS, OTS and FODS files in a Web Worker, so parsing does not block the page and can
   * be stopped by a timeout. Without this option, parsing runs on the main thread exactly as before.
   */
  readonly worker?: PreviewWorkerOptions;
}

export interface PreviewWorkerOptions {
  /**
   * Creates the parser worker, a module worker running the package's `./worker` entry point, for example
   * `() => new Worker(new URL("…/worker.js", import.meta.url), { type: "module" })`. The library never creates
   * workers itself (no `blob:` or `data:` scripts). Called again after a worker has been ended.
   */
  readonly create: () => Worker;
  /** Milliseconds from create() until the worker is ready. Default 10,000; 1,000 to 60,000. */
  readonly startupTimeout?: number;
  /**
   * Milliseconds from sending a file to the worker until its result arrives; the worker is then ended and the
   * load fails with "resource-limit". Default 30,000; 1,000 to 300,000. Receiving, checking and rendering the
   * result happen afterwards on the main thread and are not covered by this timeout.
   */
  readonly timeout?: number;
  /**
   * When the worker cannot be created or does not become ready: "error" (default) fails the load;
   * "main-thread" parses on the main thread instead, without a timeout. Never applies to cancelled loads or to
   * failures after the worker was ready.
   */
  readonly fallback?: "error" | "main-thread";
}

export interface PreviewLoadOptions {
  /** Cancels this load when aborted, as a new load or destroy() does; the load then resolves without showing an error. */
  readonly signal?: AbortSignal;
}
