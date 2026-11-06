// The parser worker inside a Node worker thread (tests only): installParserWorker on a scope that maps the Web
// Worker API onto parentPort. With workerData.report, every result is followed by a report of the transferred
// buffers' lengths as seen by the worker after sending (0 = detached), on a separate message.
import { parentPort, workerData } from "node:worker_threads";
import { installParserWorker } from "../../dist/worker/worker-main.js";

installParserWorker({
  addEventListener(type, listener) { parentPort.on("message", data => listener({ data })); },
  postMessage(message, transfer = []) {
    parentPort.postMessage(message, transfer);
    if (workerData?.report && message?.type === "result") parentPort.postMessage({ report: true, transferred: transfer.length, lengthsAfterSending: transfer.map(buffer => buffer.byteLength) });
  },
});
