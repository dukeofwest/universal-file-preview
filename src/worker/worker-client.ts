import { PreviewError } from "../errors.js";
import { PROTOCOL_VERSION, readWorkerMessage, validateModel, type DocumentModel, type ParseRequest, type WorkerFormat, type WorkerMessage } from "./protocol.js";

/**
 * The viewer's side of the parser worker. Internal module; FilePreviewer creates one client per viewer.
 *
 * Lifecycle: [none] -create()-> [starting] -"ready"-> [idle] -parse-> [busy] -result/error-> [idle].
 * Any failure, a timeout, cancellation and destroy() terminate the worker and return to [none]; the next
 * parse creates a new worker. One parse is in flight at a time: a new parse cancels the previous one.
 *
 * Failures before "ready" (create() throwing, error or messageerror events, an invalid message, the start-up
 * timeout) reject with WorkerStartupError, so the caller can apply its fallback. Cancellation always rejects
 * with "aborted", never with WorkerStartupError, so it cannot trigger a fallback.
 */

export interface WorkerClientOptions {
  /** Creates the worker (the application's factory; the library never creates workers itself). */
  readonly create: () => Worker;
  /** Milliseconds from create() until "ready". Default 10,000; 1,000 to 60,000. */
  readonly startupTimeout?: number;
  /** Milliseconds from posting a parse request until its result or error. Default 30,000; 1,000 to 300,000. */
  readonly timeout?: number;
  /** Largest input sent to the worker (the viewer's maxFileSize). */
  readonly maxFileSize: number;
}

export interface WorkerParseOptions { readonly signal?: AbortSignal; readonly locale?: string }

/** The worker could not be started; the caller may fall back to parsing on the main thread. */
export class WorkerStartupError extends PreviewError {}

export const DEFAULT_STARTUP_TIMEOUT = 10_000, DEFAULT_PARSE_TIMEOUT = 30_000;

function timeoutOption(value: number | undefined, fallback: number, max: number, name: string): number {
  if (value === undefined) return fallback;
  if (!Number.isSafeInteger(value) || value < 1000 || value > max) throw new RangeError(`${name} must be an integer from 1,000 to ${max.toLocaleString("en-US")} milliseconds.`);
  return value;
}

type State = "none" | "starting" | "idle" | "busy";
interface Pending { readonly id: number; readonly format: WorkerFormat; readonly resolve: (model: DocumentModel) => void; readonly reject: (error: unknown) => void }
interface Startup { readonly resolve: () => void; readonly reject: (error: unknown) => void }

const aborted = (cause?: unknown) => new PreviewError("Preview loading was cancelled.", "aborted", cause === undefined ? undefined : { cause });

export class WorkerClient {
  readonly #create: () => Worker;
  readonly #startupTimeout: number;
  readonly #timeout: number;
  readonly #maxFileSize: number;
  #state: State = "none";
  #worker: Worker | undefined;
  #nextId = 1;
  #calls = 0;
  #startup: Startup | undefined;
  #pending: Pending | undefined;
  #timer: ReturnType<typeof setTimeout> | undefined;
  #destroyed = false;
  readonly #onMessage = (event: MessageEvent) => this.#message(event.data);
  readonly #onError = (event: Event) => this.#failure(new PreviewError("The parser worker failed.", "renderer-error", { cause: event }));
  readonly #onMessageError = (event: Event) => this.#failure(new PreviewError("A message from the parser worker could not be read.", "renderer-error", { cause: event }));

  constructor(options: WorkerClientOptions) {
    if (typeof options.create !== "function") throw new TypeError("worker.create must be a function.");
    this.#create = options.create;
    this.#startupTimeout = timeoutOption(options.startupTimeout, DEFAULT_STARTUP_TIMEOUT, 60_000, "worker.startupTimeout");
    this.#timeout = timeoutOption(options.timeout, DEFAULT_PARSE_TIMEOUT, 300_000, "worker.timeout");
    this.#maxFileSize = options.maxFileSize;
  }

  /** The lifecycle state (for tests and diagnostics). */
  get state(): State { return this.#state; }

  /**
   * Parses `bytes` in the worker and returns the validated model. The bytes are copied into a buffer that is
   * transferred to the worker; the caller's array is not detached.
   */
  async parse(format: WorkerFormat, bytes: Uint8Array, name: string, options: WorkerParseOptions = {}): Promise<DocumentModel> {
    if (this.#destroyed) throw aborted();
    const { signal } = options;
    if (signal?.aborted) throw aborted(signal.reason);
    if (bytes.byteLength > this.#maxFileSize) throw new PreviewError("File exceeds the configured size limit.", "resource-limit");
    // One parse at a time: a running parse is cancelled by terminating its worker.
    if (this.#state === "busy" || this.#state === "starting") this.#terminate(aborted());
    const call = ++this.#calls;
    const onAbort = () => { if (call === this.#calls) this.#terminate(aborted(signal?.reason)); };
    signal?.addEventListener("abort", onAbort, { once: true });
    try {
      if (this.#state === "none") await this.#start();
      // A newer parse started while this one waited for "ready": this one is cancelled.
      if (this.#destroyed || signal?.aborted || call !== this.#calls) throw aborted(signal?.reason);
      return await this.#post(format, bytes, name, options.locale);
    } finally {
      signal?.removeEventListener("abort", onAbort);
    }
  }

  /** Terminates the worker and cancels any parse in flight. Idempotent. */
  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#terminate(aborted());
  }

  #start(): Promise<void> {
    let worker: Worker;
    try { worker = this.#create(); }
    catch (error) { throw new WorkerStartupError("The parser worker could not be created.", "renderer-error", { cause: error }); }
    if (!worker || typeof worker.postMessage !== "function" || typeof worker.addEventListener !== "function") throw new WorkerStartupError("worker.create did not return a Worker.", "renderer-error");
    this.#worker = worker;
    this.#state = "starting";
    worker.addEventListener("message", this.#onMessage);
    worker.addEventListener("error", this.#onError);
    worker.addEventListener("messageerror", this.#onMessageError);
    return new Promise<void>((resolve, reject) => {
      this.#startup = { resolve, reject };
      this.#timer = setTimeout(() => this.#terminate(new WorkerStartupError(`The parser worker did not start within ${this.#startupTimeout.toLocaleString("en-US")} ms.`, "resource-limit")), this.#startupTimeout);
    });
  }

  #post(format: WorkerFormat, bytes: Uint8Array, name: string, locale: string | undefined): Promise<DocumentModel> {
    const worker = this.#worker!, id = this.#nextId++;
    // A copy of exactly the view's bytes in a buffer of its own, transferred: the caller keeps its array.
    const buffer = bytes.slice().buffer;
    const request: ParseRequest = { v: PROTOCOL_VERSION, type: "parse", id, format, name, ...(locale !== undefined ? { locale } : {}), bytes: buffer };
    return new Promise<DocumentModel>((resolve, reject) => {
      this.#pending = { id, format, resolve, reject };
      this.#state = "busy";
      try { worker.postMessage(request, [buffer]); }
      catch (error) { this.#terminate(new PreviewError("The file could not be sent to the parser worker.", "renderer-error", { cause: error })); return; }
      this.#timer = setTimeout(() => this.#terminate(new PreviewError(`Parsing did not finish within ${this.#timeout.toLocaleString("en-US")} ms.`, "resource-limit")), this.#timeout);
    });
  }

  #message(data: unknown): void {
    const message = readWorkerMessage(data);
    if (this.#state === "starting") {
      if (message?.type !== "ready") { this.#failure(new PreviewError("The parser worker sent an invalid message.", "renderer-error")); return; }
      clearTimeout(this.#timer); this.#timer = undefined;
      this.#state = "idle";
      const startup = this.#startup!; this.#startup = undefined;
      startup.resolve();
      return;
    }
    if (!message || message.type === "ready" || message.id >= this.#nextId) { this.#failure(new PreviewError("The parser worker sent an invalid message.", "renderer-error")); return; }
    const pending = this.#pending;
    // An id that was issued but is no longer pending: a cancelled or timed-out request. Ignored.
    if (!pending || message.id !== pending.id) return;
    this.#settle(pending, message);
  }

  #settle(pending: Pending, message: Exclude<WorkerMessage, { type: "ready" }>): void {
    clearTimeout(this.#timer); this.#timer = undefined;
    this.#pending = undefined;
    this.#state = "idle";
    if (message.type === "error") { pending.reject(new PreviewError(message.error.message, message.error.code)); return; }
    let model: DocumentModel;
    try { model = validateModel(message.model, pending.format); }
    catch (error) { pending.reject(error); return; }
    pending.resolve(model);
  }

  /** A failure signal from the worker: before "ready" a start-up failure, afterwards a failure of the parse in flight. */
  #failure(error: PreviewError): void {
    this.#terminate(this.#state === "starting" ? new WorkerStartupError(error.message, error.code, { cause: error }) : error);
  }

  /** Terminates the worker (if any), clears the timer, removes every listener and rejects what is waiting. */
  #terminate(error: unknown): void {
    clearTimeout(this.#timer); this.#timer = undefined;
    const worker = this.#worker, startup = this.#startup, pending = this.#pending;
    this.#worker = undefined; this.#startup = undefined; this.#pending = undefined;
    this.#state = "none";
    if (worker) {
      worker.removeEventListener("message", this.#onMessage);
      worker.removeEventListener("error", this.#onError);
      worker.removeEventListener("messageerror", this.#onMessageError);
      worker.terminate();
    }
    startup?.reject(error);
    pending?.reject(error);
  }
}
