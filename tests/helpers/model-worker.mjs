// Worker side of tests/model-budget.test.mjs: parses a document in a worker thread with the same parsers and
// budget as the main thread, checks the finished model again, and posts it back (structured clone).
import { parentPort } from "node:worker_threads";
import { parseDocx, parsePptx, parseXlsx } from "../../dist/parsers/ooxml.js";
import { parseOds } from "../../dist/parsers/odf/ods.js";
import { checkModel, modelBackstopRejections, modelBudgetMisses } from "../../dist/model-budget.js";

const parsers = { docx: parseDocx, pptx: parsePptx, xlsx: parseXlsx, ods: parseOds };

parentPort.on("message", ({ id, kind, bytes, name }) => {
  try {
    const model = parsers[kind](bytes, name);
    const usage = checkModel(model);
    parentPort.postMessage({ id, model, usage, rejections: modelBackstopRejections(), misses: modelBudgetMisses() });
  } catch (error) {
    parentPort.postMessage({ id, error: { name: error?.name, code: error?.code, message: error?.message }, rejections: modelBackstopRejections(), misses: modelBudgetMisses() });
  }
});
