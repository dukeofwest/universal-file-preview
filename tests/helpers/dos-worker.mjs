// Runs one denial-of-service corpus file in a worker thread with a capped heap (tests/dos-budget.security.mjs).
import { parentPort } from "node:worker_threads";
import { parseModel } from "../../dist/parsers/parse-model.js";
import { sanitizeSvg } from "../../dist/renderers/image-url.js";

parentPort.once("message", ({ kind, name, bytes }) => {
  // Self-test of the watchdog in tests/dos-budget.security.mjs: a worker that never answers.
  if (kind === "watchdog-self-test") for (;;) { /* never returns */ }
  const started = performance.now();
  let outcome, code, message;
  try {
    if (kind === "svg") sanitizeSvg(bytes);
    else parseModel(kind, bytes, name, { locale: "en-US" });
    outcome = "handled";
  } catch (error) {
    outcome = error?.name === "PreviewError" ? "refused" : "crashed";
    code = error?.code; message = String(error?.message ?? error).slice(0, 300);
  }
  const ms = Math.round(performance.now() - started);
  parentPort.postMessage({ outcome, code, message, ms, heapUsedMiB: Math.round(process.memoryUsage().heapUsed / 1048576) });
});
