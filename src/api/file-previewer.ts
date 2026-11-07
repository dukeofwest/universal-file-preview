import { detectFormat } from "../detection/detect-format.js";
import { PreviewError } from "../errors.js";
import { viewerLocale } from "../parsers/parse-model.js";
import { createDefaultRendererRegistry, isBuiltInDocumentRenderer, type ParsedModelRenderer, type RendererRegistry } from "../renderers/registry.js";
import type { RendererContext, RendererInstance } from "../renderers/types.js";
import { isWorkerFormat, type WorkerFormat } from "../worker/protocol.js";
import { WorkerClient, WorkerStartupError } from "../worker/worker-client.js";
import type { PreviewFile, PreviewLoadOptions, PreviewOptions } from "./types.js";
import { attachViewerStyles, type ViewerStyles } from "./viewer-styles.js";

/** How long the download URL stays valid after the click (enough for the browser to start the download). */
const DOWNLOAD_URL_LIFETIME_MS = 30_000;

export class FilePreviewer {
  readonly #root: HTMLDivElement;
  readonly #status: HTMLDivElement;
  readonly #toolbar: HTMLDivElement;
  readonly #content: HTMLDivElement;
  readonly #registry: RendererRegistry;
  readonly #maxFileSize: number;
  readonly #styles: ViewerStyles;
  /** Worker mode (options.worker): one parser worker per viewer, created on first use. */
  readonly #worker: WorkerClient | undefined;
  readonly #fallback: "error" | "main-thread";
  #controller: AbortController | undefined;
  #instance: RendererInstance | undefined;
  #generation = 0;
  #zoom = 1;
  #destroyed = false;
  #downloadUrl: string | undefined;
  #title: HTMLSpanElement | undefined;

  constructor(private readonly options: PreviewOptions) {
    this.#registry = options.registry ?? createDefaultRendererRegistry();
    this.#maxFileSize = options.maxFileSize ?? 128 * 1024 * 1024;
    if (!Number.isSafeInteger(this.#maxFileSize) || this.#maxFileSize < 1) throw new RangeError("maxFileSize must be a positive safe integer.");
    if (options.styleNonce !== undefined && (typeof options.styleNonce !== "string" || options.styleNonce === "")) throw new TypeError("styleNonce must be a non-empty string.");
    const worker = options.worker;
    if (worker !== undefined && (typeof worker !== "object" || worker === null)) throw new TypeError("worker must be an object.");
    this.#fallback = worker?.fallback ?? "error";
    if (this.#fallback !== "error" && this.#fallback !== "main-thread") throw new RangeError('worker.fallback must be "error" or "main-thread".');
    // Validates create and the timeouts; no worker is created until a file needs one.
    this.#worker = worker ? new WorkerClient({ create: worker.create, maxFileSize: this.#maxFileSize, ...(worker.startupTimeout !== undefined ? { startupTimeout: worker.startupTimeout } : {}), ...(worker.timeout !== undefined ? { timeout: worker.timeout } : {}) }) : undefined;
    this.#root = document.createElement("div"); this.#root.className = "ufp-root";
    this.#toolbar = document.createElement("div"); this.#toolbar.className = "ufp-toolbar"; this.#toolbar.setAttribute("role", "toolbar");
    this.#status = document.createElement("div"); this.#status.className = "ufp-status"; this.#status.setAttribute("role", "status"); this.#status.setAttribute("aria-live", "polite");
    this.#content = document.createElement("div"); this.#content.className = "ufp-content";
    this.#root.append(this.#toolbar, this.#status, this.#content);
    this.#styles = attachViewerStyles(options.container, options.styleNonce);
    if (this.#styles.element) options.container.append(this.#styles.element);
    options.container.append(this.#root);
    this.#buildToolbar();
  }

  open(): Promise<void> { return this.load(this.options.file); }

  async load(file: PreviewFile, options: PreviewLoadOptions = {}): Promise<void> {
    this.#assertAlive();
    const external = options.signal;
    if (external !== undefined && (typeof external !== "object" || external === null || typeof external.addEventListener !== "function")) throw new TypeError("signal must be an AbortSignal.");
    this.#controller?.abort();
    this.#instance?.destroy?.(); this.#instance = undefined;
    this.#currentFile = undefined; this.#currentBytes = undefined;
    this.#content.replaceChildren(); this.#status.textContent = "Loading preview..."; this.#status.className = "ufp-status";
    this.#controller = new AbortController();
    const controller = this.#controller, generation = ++this.#generation;
    // The caller's signal cancels this load like a new load does; its listener is removed when the load ends.
    const cancel = () => controller.abort(external?.reason);
    if (external?.aborted) controller.abort(external.reason);
    else external?.addEventListener("abort", cancel, { once: true });
    try {
      const { bytes, mimeType } = await readSource(file, controller.signal, this.#maxFileSize);
      this.#checkCurrent(generation, controller.signal);
      const format = detectFormat(file.name, file.mimeType ?? mimeType, bytes);
      const renderer = this.#registry.get(format);
      if (!renderer) throw new PreviewError(`Preview is not available for ${format === "unknown" ? "this file type" : format.toUpperCase()}.`, "unsupported-format");
      const context: RendererContext = { file, format, bytes, signal: controller.signal };
      const instance = this.#worker && isBuiltInDocumentRenderer(renderer) && isWorkerFormat(format)
        ? await this.#renderInWorker(this.#worker, renderer, format, context, generation)
        : await renderer.render(context);
      if (controller.signal.aborted || generation !== this.#generation) { instance.destroy?.(); return; }
      this.#instance = instance; this.#content.replaceChildren(instance.element);
      if (this.#zoom !== 1) instance.setZoom?.(this.#zoom);
      this.#status.textContent = `${file.name} - ${format.toUpperCase()}`;
      this.#status.className = "ufp-status ufp-status-ready";
      if (this.#title) this.#title.textContent = file.name, this.#title.title = file.name;
      this.#currentFile = file; this.#currentBytes = bytes;
    } catch (error) {
      if (generation !== this.#generation) return;
      const previewError = error instanceof PreviewError ? error : controller.signal.aborted
        ? new PreviewError("Preview loading was cancelled.", "aborted", { cause: error })
        : new PreviewError("The file could not be previewed.", "parse-error", { cause: error });
      // Cancelled by the caller's signal (a new load or destroy() has already moved the generation on): not an error.
      if (previewError.code === "aborted") { this.#status.textContent = "Preview loading was cancelled."; this.#status.className = "ufp-status"; return; }
      this.#status.textContent = previewError.message; this.#status.className = "ufp-status ufp-status-error";
      throw previewError;
    } finally {
      external?.removeEventListener("abort", cancel);
    }
  }

  /**
   * Worker mode: parses in the worker, then renders the validated model on the main thread. A start-up failure
   * of a load that is still current falls back to main-thread parsing when the options allow it; cancellation
   * never does.
   */
  async #renderInWorker(worker: WorkerClient, renderer: ParsedModelRenderer, format: WorkerFormat, context: RendererContext, generation: number): Promise<RendererInstance> {
    let model;
    try { model = await worker.parse(format, context.bytes, context.file.name, { signal: context.signal, locale: viewerLocale() }); }
    catch (error) {
      if (error instanceof WorkerStartupError && this.#fallback === "main-thread" && !context.signal.aborted && generation === this.#generation) return renderer.render(context);
      throw error;
    }
    this.#checkCurrent(generation, context.signal);
    return renderer.renderParsed(model, context);
  }

  #currentFile: PreviewFile | undefined;
  #currentBytes: Uint8Array | undefined;
  #downloadTimer: ReturnType<typeof setTimeout> | undefined;

  setZoom(value: number): void {
    this.#assertAlive();
    if (!Number.isFinite(value)) throw new RangeError("Zoom must be a finite number.");
    this.#zoom = Math.min(4, Math.max(0.1, value)); this.#instance?.setZoom?.(this.#zoom);
    this.#root.dispatchEvent(new Event("ufp-zoom"));
  }
  fitToWidth(): void { this.#assertAlive(); this.#zoom = 1; this.#instance?.fitToWidth?.(); this.#root.dispatchEvent(new Event("ufp-zoom")); }
  fitToPage(): void { this.fitToWidth(); }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true; this.#generation++; this.#controller?.abort(); this.#worker?.destroy(); this.#instance?.destroy?.();
    this.#revokeDownload();
    this.#styles.release(); this.#root.remove();
  }

  #buildToolbar(): void {
    if (this.options.toolbar === false) { this.#toolbar.hidden = true; return; }
    const title = document.createElement("span"); title.className = "ufp-file-title"; title.textContent = this.options.file.name; title.title = this.options.file.name; this.#title = title; this.#toolbar.append(title);
    if (this.options.zoom !== false) {
      this.#button("Zoom out", "-", () => this.setZoom(this.#zoom - 0.1));
      const zoomValue = document.createElement("span"); zoomValue.className = "ufp-zoom-value"; zoomValue.textContent = "100%";
      this.#root.addEventListener("ufp-zoom", () => { zoomValue.textContent = `${Math.round(this.#zoom * 100)}%`; }); this.#toolbar.append(zoomValue);
      this.#button("Zoom in", "+", () => this.setZoom(this.#zoom + 0.1));
      this.#button("Fit to width", "Fit", () => this.fitToWidth());
    }
    if (this.options.download !== false) this.#button("Download file", "Download", () => this.#download());
  }
  #button(label: string, text: string, action: () => void): void {
    const button = document.createElement("button"); button.type = "button"; button.className = "ufp-button"; button.textContent = text; button.setAttribute("aria-label", label); button.onclick = action; this.#toolbar.append(button);
  }
  /**
   * Saves the original file. The blob is always application/octet-stream, whatever type the source declared, so
   * the browser saves it instead of displaying it; its URL is revoked shortly after the download starts.
   */
  #download(): void {
    if (!this.#currentFile || !this.#currentBytes) return;
    this.#revokeDownload();
    const bytes = new Uint8Array(this.#currentBytes.byteLength); bytes.set(this.#currentBytes);
    const url = URL.createObjectURL(new Blob([bytes.buffer], { type: "application/octet-stream" }));
    this.#downloadUrl = url;
    const link = document.createElement("a"); link.href = url; link.download = this.#currentFile.name; link.click();
    this.#downloadTimer = setTimeout(() => { if (this.#downloadUrl === url) this.#revokeDownload(); }, DOWNLOAD_URL_LIFETIME_MS);
  }
  #revokeDownload(): void {
    clearTimeout(this.#downloadTimer); this.#downloadTimer = undefined;
    if (this.#downloadUrl) URL.revokeObjectURL(this.#downloadUrl);
    this.#downloadUrl = undefined;
  }
  #assertAlive(): void { if (this.#destroyed) throw new PreviewError("This previewer has been destroyed.", "renderer-error"); }
  #checkCurrent(generation: number, signal: AbortSignal): void { if (signal.aborted || generation !== this.#generation) throw new PreviewError("Preview loading was cancelled.", "aborted"); }
}

async function readSource(file: PreviewFile, signal: AbortSignal, maxSize: number): Promise<{ bytes: Uint8Array; mimeType: string }> {
  if (typeof file.source === "string") {
    let response: Response;
    try { response = await fetch(file.source, { signal }); }
    catch (cause) { if (signal.aborted) throw cause; throw new PreviewError("Unable to fetch the preview file.", "network-error", { cause }); }
    if (!response.ok) throw new PreviewError(`File request failed with status ${response.status}.`, "network-error");
    const length = Number(response.headers.get("content-length"));
    if (Number.isFinite(length) && length > maxSize) throw new PreviewError("File exceeds the configured size limit.", "resource-limit");
    if (!response.body) {
      const buffer = new Uint8Array(await response.arrayBuffer());
      if (buffer.length > maxSize) throw new PreviewError("File exceeds the configured size limit.", "resource-limit");
      return { bytes: buffer, mimeType: response.headers.get("content-type") ?? "" };
    }
    const reader = response.body.getReader(), chunks: Uint8Array[] = [];
    let total = 0;
    try {
      while (true) {
        if (signal.aborted) throw new PreviewError("Preview loading was cancelled.", "aborted");
        const { done, value } = await reader.read(); if (done) break;
        total += value.length; if (total > maxSize) { await reader.cancel(); throw new PreviewError("File exceeds the configured size limit.", "resource-limit"); }
        chunks.push(value);
      }
    } finally { reader.releaseLock(); }
    const bytes = new Uint8Array(total); let offset = 0; for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
    return { bytes, mimeType: response.headers.get("content-type") ?? "" };
  }
  if (file.source instanceof Blob) {
    if (file.source.size > maxSize) throw new PreviewError("File exceeds the configured size limit.", "resource-limit");
    return { bytes: new Uint8Array(await file.source.arrayBuffer()), mimeType: file.source.type };
  }
  const bytes = file.source instanceof Uint8Array ? file.source : new Uint8Array(file.source);
  if (bytes.byteLength > maxSize) throw new PreviewError("File exceeds the configured size limit.", "resource-limit");
  return { bytes, mimeType: "" };
}
