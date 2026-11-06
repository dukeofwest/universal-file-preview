import { installParserWorker, type ParserWorkerScope } from "./worker/worker-main.js";

/**
 * The parser worker script (package entry point "./worker"). Load it as a module worker created by the
 * application, for example:
 *
 *   new Worker(new URL("@dukeofwest/universal-file-preview/worker", import.meta.url), { type: "module" })
 *
 * It answers parse requests from FilePreviewer's worker mode; importing it anywhere else has no use.
 */
installParserWorker(globalThis as unknown as ParserWorkerScope);
