import { detectFormat } from "../detection/detect-format.js";
import { PreviewError } from "../errors.js";
import { createDefaultRendererRegistry, type RendererRegistry } from "../renderers/registry.js";
import type { RendererInstance } from "../renderers/types.js";
import type { PreviewFile, PreviewOptions } from "./types.js";

export class FilePreviewer {
  readonly #root: HTMLDivElement;
  readonly #status: HTMLDivElement;
  readonly #toolbar: HTMLDivElement;
  readonly #content: HTMLDivElement;
  readonly #registry: RendererRegistry;
  readonly #maxFileSize: number;
  readonly #style: HTMLStyleElement;
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
    this.#root = document.createElement("div"); this.#root.className = "ufp-root";
    this.#toolbar = document.createElement("div"); this.#toolbar.className = "ufp-toolbar"; this.#toolbar.setAttribute("role", "toolbar");
    this.#status = document.createElement("div"); this.#status.className = "ufp-status"; this.#status.setAttribute("role", "status"); this.#status.setAttribute("aria-live", "polite");
    this.#content = document.createElement("div"); this.#content.className = "ufp-content";
    this.#root.append(this.#toolbar, this.#status, this.#content);
    this.#style = document.createElement("style"); this.#style.textContent = `${styles}
.ufp-word-page{position:relative}.ufp-word-page p,.ufp-word-page h2{margin:0;font-size:inherit;font-weight:inherit;white-space:pre-wrap}.ufp-word-page table{margin:0}.ufp-word-page td{padding:0}
`;
    options.container.append(this.#style, this.#root);
    this.#buildToolbar();
  }

  open(): Promise<void> { return this.load(this.options.file); }

  async load(file: PreviewFile): Promise<void> {
    this.#assertAlive();
    this.#controller?.abort();
    this.#instance?.destroy?.(); this.#instance = undefined;
    this.#currentFile = undefined; this.#currentBytes = undefined;
    this.#content.replaceChildren(); this.#status.textContent = "Loading preview..."; this.#status.className = "ufp-status";
    this.#controller = new AbortController();
    const controller = this.#controller, generation = ++this.#generation;
    try {
      const { bytes, mimeType } = await readSource(file, controller.signal, this.#maxFileSize);
      this.#checkCurrent(generation, controller.signal);
      const format = detectFormat(file.name, file.mimeType ?? mimeType, bytes);
      const renderer = this.#registry.get(format);
      if (!renderer) throw new PreviewError(`Preview is not available for ${format === "unknown" ? "this file type" : format.toUpperCase()}.`, "unsupported-format");
      const instance = await renderer.render({ file, format, bytes, signal: controller.signal });
      if (controller.signal.aborted || generation !== this.#generation) { instance.destroy?.(); return; }
      this.#instance = instance; this.#content.replaceChildren(instance.element);
      if (this.#zoom !== 1) instance.setZoom?.(this.#zoom);
      this.#status.textContent = `${file.name} - ${format.toUpperCase()}`;
      this.#status.className = "ufp-status ufp-status-ready";
      if (this.#title) this.#title.textContent = file.name, this.#title.title = file.name;
      this.#currentFile = file; this.#currentBytes = bytes; this.#currentMime = mimeType;
    } catch (error) {
      if (generation !== this.#generation) return;
      const previewError = error instanceof PreviewError ? error : controller.signal.aborted
        ? new PreviewError("Preview loading was cancelled.", "aborted", { cause: error })
        : new PreviewError("The file could not be previewed.", "parse-error", { cause: error });
      if (previewError.code === "aborted") return;
      this.#status.textContent = previewError.message; this.#status.className = "ufp-status ufp-status-error";
      throw previewError;
    }
  }

  #currentFile: PreviewFile | undefined;
  #currentBytes: Uint8Array | undefined;
  #currentMime = "application/octet-stream";

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
    this.#destroyed = true; this.#generation++; this.#controller?.abort(); this.#instance?.destroy?.();
    if (this.#downloadUrl) URL.revokeObjectURL(this.#downloadUrl);
    this.#style.remove(); this.#root.remove();
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
  #download(): void {
    if (!this.#currentFile || !this.#currentBytes) return;
    if (this.#downloadUrl) URL.revokeObjectURL(this.#downloadUrl);
    const bytes = new Uint8Array(this.#currentBytes.byteLength); bytes.set(this.#currentBytes);
    this.#downloadUrl = URL.createObjectURL(new Blob([bytes.buffer], { type: this.#currentMime }));
    const link = document.createElement("a"); link.href = this.#downloadUrl; link.download = this.#currentFile.name; link.click();
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

const styles = `
@font-face{font-family:"ufp Franklin Gothic Book";src:local("Segoe UI"),local("SegoeUI");size-adjust:97.6%;ascent-override:93.9%;descent-override:22.3%;line-gap-override:0%}@font-face{font-family:"ufp Franklin Gothic Book";font-style:italic;src:local("Segoe UI Italic"),local("SegoeUI-Italic");size-adjust:100%;ascent-override:91.7%;descent-override:21.7%;line-gap-override:0%}@font-face{font-family:"ufp Franklin Gothic Book Alt";src:local("Arial"),local("ArialMT"),local("Helvetica"),local("Liberation Sans"),local("Arimo");size-adjust:96%;ascent-override:95.5%;descent-override:22.6%;line-gap-override:0%}@font-face{font-family:"ufp Gill Sans MT";src:local("Corbel");size-adjust:99.2%;ascent-override:93.6%;descent-override:23.2%;line-gap-override:0%;font-feature-settings:"lnum" 1,"tnum" 1}@font-face{font-family:"ufp Gill Sans MT";font-weight:bold;src:local("Corbel Bold"),local("Corbel-Bold");size-adjust:110.4%;ascent-override:84.1%;descent-override:20.8%;line-gap-override:0%;font-feature-settings:"lnum" 1,"tnum" 1}@font-face{font-family:"ufp Gill Sans MT";font-style:italic;src:local("Corbel Italic"),local("Corbel-Italic");size-adjust:94.2%;ascent-override:98.6%;descent-override:24.4%;line-gap-override:0%;font-feature-settings:"lnum" 1,"tnum" 1}@font-face{font-family:"ufp Gill Sans MT Alt";src:local("Arial"),local("ArialMT"),local("Helvetica"),local("Liberation Sans"),local("Arimo");size-adjust:91.6%;ascent-override:101.4%;descent-override:25.1%;line-gap-override:0%}@font-face{font-family:"ufp Gill Sans MT Alt";font-weight:bold;src:local("Arial Bold"),local("Arial-BoldMT"),local("Helvetica Bold"),local("Liberation Sans Bold"),local("Arimo Bold");size-adjust:100.5%;ascent-override:92.4%;descent-override:22.9%;line-gap-override:0%}@font-face{font-family:"ufp Gill Sans MT Alt";font-style:italic;src:local("Arial Italic"),local("Arial-ItalicMT"),local("Helvetica Oblique"),local("Liberation Sans Italic"),local("Arimo Italic");size-adjust:83.7%;ascent-override:111.0%;descent-override:27.5%;line-gap-override:0%}
.ufp-root{font:14px/1.45 system-ui,sans-serif;color:#20242a;background:#f3f5f7;border:1px solid #d5d9de;border-radius:8px;display:flex;flex-direction:column;min-width:0;min-height:240px;height:100%;overflow:hidden}
.ufp-root *{box-sizing:border-box}.ufp-toolbar{display:flex;align-items:center;gap:8px;padding:8px 10px;background:#fff;border-bottom:1px solid #d5d9de;min-height:46px}.ufp-toolbar[hidden]{display:none}.ufp-file-title{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-right:auto;font-weight:600}.ufp-button{font:inherit;color:inherit;background:#fff;border:1px solid #c8cdd3;border-radius:5px;padding:5px 9px;cursor:pointer}.ufp-button:hover{background:#edf1f5}.ufp-button:focus-visible,.ufp-sheet-tab:focus-visible{outline:2px solid #2864dc;outline-offset:2px}.ufp-zoom-value{min-width:42px;text-align:center;font-variant-numeric:tabular-nums}.ufp-status{padding:9px 12px;color:#555;background:#fafbfc}.ufp-status-ready{font-size:12px;color:#56606b;border-bottom:1px solid #e1e4e8}.ufp-status-error{color:#8e1c1c;background:#fff1f0}.ufp-content{flex:1;min-height:0;overflow:auto;padding:16px}.ufp-image{min-height:100%;display:flex;justify-content:center;align-items:flex-start}.ufp-image img{display:block;max-width:none;height:auto;object-fit:contain}.ufp-word-document{display:flex;flex-direction:column;align-items:center;gap:18px}.ufp-word-page{background:white;color:#111;max-width:100%;min-height:700px;margin:0 auto;padding:68px 76px;box-shadow:0 2px 10px #0001;overflow-wrap:anywhere}.ufp-word-header,.ufp-word-footer{flex:none}.ufp-word-body{min-height:1em}.ufp-word-page p{margin:0 0 8pt;white-space:pre-wrap}.ufp-word-page table{border-collapse:collapse;table-layout:fixed;width:100%;margin:8pt 0}.ufp-word-page td{padding:5px;vertical-align:top;overflow-wrap:normal;word-break:normal}.ufp-word-page .ufp-word-image{max-width:100%;height:auto;vertical-align:middle}.ufp-word-list-label{white-space:pre}.ufp-sheet-view{height:100%;display:flex;flex-direction:column;min-height:0}.ufp-sheet-tabs{display:flex;overflow-x:auto;border-bottom:1px solid #cbd0d6;background:#fff}.ufp-sheet-tab{border:0;border-bottom:2px solid transparent;background:transparent;padding:8px 12px;font:inherit;white-space:nowrap;cursor:pointer}.ufp-sheet-tab[aria-selected=true]{border-bottom-color:#2365c7;color:#174d96}.ufp-grid-viewport{flex:1;min-height:160px;overflow:auto;background:white}.ufp-grid-scroll{position:relative;width:100%;height:100%;overflow:auto;background:#fff}.ufp-grid-canvas{position:relative}.ufp-grid-scroll *{box-sizing:border-box}.ufp-grid-pane-top{position:sticky;top:0;z-index:3;background:#fff}.ufp-grid-pane-corner{position:sticky;left:0;z-index:2;background:#fff}.ufp-grid-pane-left{position:sticky;left:0;z-index:2;background:#fff}.ufp-grid-body{position:absolute;inset:0;z-index:1}.ufp-cell-header{position:absolute;overflow:hidden;white-space:nowrap;background:#f5f5f5;color:#444;text-align:center;border-right:1px solid #d4d4d4;border-bottom:1px solid #d4d4d4}.ufp-cell-corner{background:#f5f5f5}.ufp-grid-cell{position:absolute;display:flex;flex-direction:column;overflow:hidden;pointer-events:none}.ufp-grid-cell a{pointer-events:auto}.ufp-cell-text{min-width:0}.ufp-gridline,.ufp-cell-fill,.ufp-cell-border,.ufp-data-bar,.ufp-cell-icon,.ufp-filter-button,.ufp-cell-comment{position:absolute;pointer-events:none}.ufp-cell-comment{pointer-events:auto}.ufp-grid-selection{position:absolute;border:2px solid #107c41;pointer-events:none;z-index:2}.ufp-input-message{position:absolute;z-index:3;width:200px;padding:3px 6px;background:#ffffe1;border:1px solid #767676;color:#000;font:12px/1.25 "Segoe UI",Tahoma,sans-serif;white-space:pre-wrap;overflow-wrap:break-word;pointer-events:none}.ufp-sheet-drawing{position:absolute;z-index:1;overflow:visible;pointer-events:none;container-type:inline-size}.ufp-sheet-slicer{position:absolute;z-index:1;overflow:hidden;pointer-events:none}.ufp-slicer-caption,.ufp-slicer-item{overflow:hidden;white-space:nowrap;text-overflow:ellipsis}.ufp-presentation{max-width:1100px;margin:0 auto}.ufp-slide-nav{display:flex;justify-content:center;align-items:center;gap:12px;margin:0 0 12px}.ufp-slide-nav button{font:inherit;padding:5px 9px}.ufp-slide{width:min(100%,960px);min-height:540px;margin:auto;padding:54px;background:white;box-shadow:0 2px 12px #0002;transform-origin:top center}.ufp-slide h1{font-size:2em}.ufp-slide p{white-space:pre-wrap}
.ufp-slide{position:relative;width:min(100%,960px);height:auto;min-height:0;padding:0;overflow:hidden;container-type:inline-size;background:#fff;isolation:isolate}.ufp-slide h1{font-size:2em}.ufp-slide>p{white-space:pre-wrap}.ufp-pptx-image,.ufp-pptx-shape,.ufp-pptx-text,.ufp-pptx-table,.ufp-pptx-line{box-sizing:border-box}.ufp-pptx-image img{display:block}.ufp-pptx-text{line-height:normal;color:#000}.ufp-pptx-paragraph{flex:none;min-height:0;white-space:inherit;overflow-wrap:anywhere}.ufp-pptx-table{overflow:hidden}.ufp-pptx-cell{min-width:0;min-height:0;overflow:hidden;box-sizing:border-box;overflow-wrap:anywhere}.ufp-pptx-cell p{margin:0}.ufp-pptx-bullet{white-space:pre}
`;
