import { PreviewError } from "../errors.js";
import { parseModel } from "../parsers/parse-model.js";
import { errorMessage, PROTOCOL_VERSION, readParseRequest, type ResultMessage } from "./protocol.js";
import { prepareTransfer } from "./transfer.js";

/**
 * The parser worker: answers parse requests from the viewer with a model or an error (protocol version 1).
 * Internal module; the package's worker entry point (worker.ts) installs it on the worker's global scope.
 * It parses with the same parsers and limits as main-thread mode; it has no DOM and needs none.
 */

/** The part of a worker's global scope the parser worker uses (DedicatedWorkerGlobalScope, or a test double). */
export interface ParserWorkerScope {
  addEventListener(type: "message", listener: (event: { readonly data: unknown }) => void): void;
  postMessage(message: unknown, transfer?: Transferable[]): void;
}

const DEFAULT_LOCALE = "en-US";

export function installParserWorker(scope: ParserWorkerScope): void {
  scope.addEventListener("message", event => {
    const request = readParseRequest(event.data);
    if (!request) {
      // A request the viewer did not send in this shape: answered only when it carries a usable id.
      const id = (event.data as { id?: unknown } | null)?.id;
      if (Number.isSafeInteger(id) && (id as number) >= 1) scope.postMessage(errorMessage(id as number, new PreviewError("The parser worker received an invalid request.", "renderer-error")));
      return;
    }
    let message: ResultMessage, transfer: ArrayBuffer[];
    try {
      const model = parseModel(request.format, new Uint8Array(request.bytes), request.name, { locale: request.locale ?? DEFAULT_LOCALE });
      transfer = prepareTransfer(model, request.bytes);
      message = { v: PROTOCOL_VERSION, type: "result", id: request.id, model };
    } catch (error) {
      scope.postMessage(errorMessage(request.id, error));
      return;
    }
    scope.postMessage(message, transfer);
  });
  scope.postMessage({ v: PROTOCOL_VERSION, type: "ready" });
}
