import { installParserWorker, type ParserWorkerScope } from "./worker/worker-main.js";

/**
 * The parser worker script (package entry point "./worker"). The application creates it as a module worker
 * and passes the factory to FilePreviewer's `worker.create` option, for example:
 *
 *   new Worker(new URL("…/universal-file-preview/dist/worker.js", import.meta.url), { type: "module" })
 *
 * It answers parse requests from FilePreviewer's worker mode and runs only inside a worker.
 */
const scope = globalThis as unknown as ParserWorkerScope & { WorkerGlobalScope?: abstract new () => unknown };
if (typeof scope.WorkerGlobalScope !== "function" || !(globalThis instanceof scope.WorkerGlobalScope)) {
  throw new Error("The universal-file-preview parser worker must be loaded as a Web Worker (new Worker(url, { type: \"module\" })), not imported.");
}
installParserWorker(scope);
