import { PreviewError, type PreviewErrorCode } from "../errors.js";
import { checkModel } from "../model-budget.js";
import type { Presentation, Workbook, WordDocument } from "../model.js";
import type { ExtendedWorkbook } from "../model-extensions.js";

/**
 * Messages between the viewer and its parser worker (version 1), and their run-time validation. Internal
 * module. Both sides validate everything they receive: a message that does not match its exact shape is
 * invalid, and the viewer then terminates the worker.
 */

export const PROTOCOL_VERSION = 1;
export const WORKER_FORMATS = ["docx", "xlsx", "pptx", "ods", "ots", "fods"] as const;
export type WorkerFormat = typeof WORKER_FORMATS[number];
export type DocumentModel = WordDocument | Presentation | Workbook | ExtendedWorkbook;

/** Longest file name sent to the worker, and longest error message accepted from it. */
export const MAX_NAME_LENGTH = 1024, MAX_ERROR_MESSAGE_LENGTH = 500;
/** Largest input the worker accepts (the viewer's maxFileSize is checked before sending). */
export const MAX_WORKER_INPUT_BYTES = 1024 * 1024 * 1024;
const LOCALE = /^[A-Za-z0-9-]{1,64}$/;
const ERROR_CODES: ReadonlySet<string> = new Set<PreviewErrorCode>(["unsupported-format", "invalid-file", "resource-limit", "network-error", "parse-error", "unsupported-feature", "renderer-error", "aborted"]);

export interface ReadyMessage { readonly v: 1; readonly type: "ready" }
export interface ParseRequest { readonly v: 1; readonly type: "parse"; readonly id: number; readonly format: WorkerFormat; readonly name: string; readonly locale?: string; readonly bytes: ArrayBuffer }
export interface ResultMessage { readonly v: 1; readonly type: "result"; readonly id: number; readonly model: unknown }
export interface ErrorMessage { readonly v: 1; readonly type: "error"; readonly id: number; readonly error: { readonly code: PreviewErrorCode; readonly message: string } }
export type WorkerMessage = ReadyMessage | ResultMessage | ErrorMessage;

function isRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}
/** The object has exactly these own keys (optional ones may be missing). */
function hasKeys(value: Record<string, unknown>, required: readonly string[], optional: readonly string[] = []): boolean {
  const keys = Object.keys(value);
  return required.every(key => Object.hasOwn(value, key)) && keys.every(key => required.includes(key) || optional.includes(key));
}
const isId = (value: unknown): value is number => Number.isSafeInteger(value) && (value as number) >= 1;
export const isWorkerFormat = (value: unknown): value is WorkerFormat => typeof value === "string" && (WORKER_FORMATS as readonly string[]).includes(value);

/** A message from the worker in its exact shape, or undefined when it is invalid. Error codes outside the known set become "renderer-error"; messages are truncated. */
export function readWorkerMessage(data: unknown): WorkerMessage | undefined {
  if (!isRecord(data) || data.v !== PROTOCOL_VERSION) return undefined;
  if (data.type === "ready") return hasKeys(data, ["v", "type"]) ? { v: 1, type: "ready" } : undefined;
  if (data.type === "result") return hasKeys(data, ["v", "type", "id", "model"]) && isId(data.id) ? { v: 1, type: "result", id: data.id, model: data.model } : undefined;
  if (data.type === "error") {
    if (!hasKeys(data, ["v", "type", "id", "error"]) || !isId(data.id) || !isRecord(data.error) || !hasKeys(data.error, ["code", "message"])) return undefined;
    const { code, message } = data.error;
    if (typeof code !== "string" || typeof message !== "string") return undefined;
    return { v: 1, type: "error", id: data.id, error: { code: ERROR_CODES.has(code) ? code as PreviewErrorCode : "renderer-error", message: message.slice(0, MAX_ERROR_MESSAGE_LENGTH) } };
  }
  return undefined;
}

/** A parse request in its exact shape (checked by the worker), or undefined when it is invalid. */
export function readParseRequest(data: unknown): ParseRequest | undefined {
  if (!isRecord(data) || data.v !== PROTOCOL_VERSION || data.type !== "parse" || !hasKeys(data, ["v", "type", "id", "format", "name", "bytes"], ["locale"])) return undefined;
  const { id, format, name, locale, bytes } = data;
  if (!isId(id) || !isWorkerFormat(format) || typeof name !== "string" || name.length > MAX_NAME_LENGTH) return undefined;
  if (locale !== undefined && (typeof locale !== "string" || !LOCALE.test(locale))) return undefined;
  if (!(bytes instanceof ArrayBuffer) || bytes.byteLength > MAX_WORKER_INPUT_BYTES) return undefined;
  return { v: 1, type: "parse", id, format, name, ...(locale !== undefined ? { locale } : {}), bytes };
}

/** The error as the worker reports it: a known code and a bounded message. */
export function errorMessage(id: number, error: unknown): ErrorMessage {
  const code = error instanceof PreviewError && ERROR_CODES.has(error.code) ? error.code : "renderer-error";
  const message = error instanceof PreviewError ? error.message : "The file could not be previewed.";
  return { v: 1, type: "error", id, error: { code, message: message.slice(0, MAX_ERROR_MESSAGE_LENGTH) } };
}

/** The model kind each worker format produces. */
export function expectedKind(format: WorkerFormat): DocumentModel["kind"] {
  return format === "docx" ? "word" : format === "pptx" ? "presentation" : "workbook";
}

const invalidModel = (where: string) => new PreviewError(`The parser worker returned an invalid document model (${where}).`, "renderer-error");

/**
 * Checks a model received from the worker before it is rendered: the output limits and plain-data rules
 * (checkModel: "resource-limit" when a limit is exceeded), then the structure renderers rely on, including
 * every kind or type value that selects a rendering path ("renderer-error" otherwise).
 */
export function validateModel(model: unknown, format: WorkerFormat): DocumentModel {
  try { checkModel(model); }
  catch (error) {
    if (error instanceof PreviewError && error.code === "resource-limit") throw error;
    throw new PreviewError("The parser worker returned an invalid document model.", "renderer-error", { cause: error });
  }
  if (!isRecord(model) || model.kind !== expectedKind(format)) throw invalidModel("kind");
  if (model.kind === "workbook") checkWorkbook(model);
  else if (model.kind === "word") checkWord(model);
  else checkPresentation(model);
  return model as unknown as DocumentModel;
}

type Check = (value: unknown, where: string) => void;
const string: Check = (value, where) => { if (typeof value !== "string") throw invalidModel(where); };
const number: Check = (value, where) => { if (typeof value !== "number") throw invalidModel(where); };
const record = (value: unknown, where: string): Record<string, unknown> => { if (!isRecord(value)) throw invalidModel(where); return value; };
const list = (value: unknown, where: string): unknown[] => { if (!Array.isArray(value)) throw invalidModel(where); return value; };
const optionalList = (value: unknown, where: string): unknown[] => value === undefined ? [] : list(value, where);
const bytes: Check = (value, where) => { if (!(value instanceof Uint8Array)) throw invalidModel(where); };
const oneOf = (value: unknown, allowed: readonly string[], where: string) => { if (typeof value !== "string" || !allowed.includes(value)) throw invalidModel(where); };

function checkImage(value: unknown, where: string): void {
  const image = record(value, where);
  bytes(image.bytes, `${where}.bytes`); string(image.mimeType, `${where}.mimeType`);
}
function checkParagraph(value: unknown, where: string): void {
  const paragraph = record(value, where);
  for (const [index, item] of list(paragraph.runs, `${where}.runs`).entries()) {
    const run = record(item, `${where}.runs[${index}]`);
    if (run.text !== undefined) string(run.text, `${where}.runs[${index}].text`);
    if (run.image !== undefined) checkImage(run.image, `${where}.runs[${index}].image`);
  }
}
function checkParagraphs(value: unknown, where: string): void { for (const [index, item] of list(value, where).entries()) checkParagraph(item, `${where}[${index}]`); }
function checkBlocks(value: unknown, where: string): void {
  for (const [index, item] of list(value, where).entries()) {
    const block = record(item, `${where}[${index}]`);
    oneOf(block.type, ["paragraph", "table", "pageBreak"], `${where}[${index}].type`);
    if (block.type === "paragraph") checkParagraph(block.value, `${where}[${index}].value`);
    if (block.type === "table") {
      for (const [rowIndex, row] of list(record(block.value, `${where}[${index}].value`).rows, `${where}[${index}].rows`).entries()) {
        for (const [cellIndex, cell] of list(record(row, `${where}[${index}].rows[${rowIndex}]`).cells, `${where}[${index}].rows[${rowIndex}].cells`).entries()) {
          checkParagraphs(record(cell, "table cell").paragraphs, `${where}[${index}].rows[${rowIndex}].cells[${cellIndex}].paragraphs`);
        }
      }
    }
  }
}
function checkWord(model: Record<string, unknown>): void {
  string(model.title, "title");
  checkBlocks(model.blocks, "blocks");
  for (const key of ["headers", "footers", "firstPageHeaders", "firstPageFooters"]) if (model[key] !== undefined) checkBlocks(model[key], key);
}
function checkElements(value: unknown, where: string): void {
  for (const [index, item] of optionalList(value, where).entries()) {
    const element = record(item, `${where}[${index}]`), at = `${where}[${index}]`;
    oneOf(element.kind, ["shape", "line", "text", "image", "table"], `${at}.kind`);
    for (const key of ["left", "top", "width", "height"]) number(element[key], `${at}.${key}`);
    if (element.kind === "image") checkImage(element.image, `${at}.image`);
    if ((element.kind === "text" || element.kind === "shape") && element.paragraphs !== undefined) checkParagraphs(element.paragraphs, `${at}.paragraphs`);
    if (element.kind === "table") {
      list(element.columns, `${at}.columns`);
      for (const [rowIndex, row] of list(element.rows, `${at}.rows`).entries()) {
        for (const [cellIndex, cell] of list(record(row, `${at}.rows[${rowIndex}]`).cells, `${at}.rows[${rowIndex}].cells`).entries()) {
          checkParagraphs(record(cell, "table cell").paragraphs, `${at}.rows[${rowIndex}].cells[${cellIndex}].paragraphs`);
        }
      }
    }
  }
}
function checkPresentation(model: Record<string, unknown>): void {
  string(model.title, "title");
  for (const [index, item] of list(model.slides, "slides").entries()) {
    const slide = record(item, `slides[${index}]`);
    string(slide.title, `slides[${index}].title`);
    checkParagraphs(slide.paragraphs, `slides[${index}].paragraphs`);
    checkElements(slide.elements, `slides[${index}].elements`);
  }
}
function checkWorkbook(model: Record<string, unknown>): void {
  string(model.title, "title");
  const sheets = list(model.sheets, "sheets");
  if (!sheets.length) throw invalidModel("sheets");
  if (model.activeSheet !== undefined) number(model.activeSheet, "activeSheet");
  for (const [index, item] of sheets.entries()) {
    const sheet = record(item, `sheets[${index}]`), at = `sheets[${index}]`;
    string(sheet.name, `${at}.name`); number(sheet.maxRow, `${at}.maxRow`); number(sheet.maxColumn, `${at}.maxColumn`);
    for (const [cellIndex, value] of list(sheet.cells, `${at}.cells`).entries()) {
      const cell = record(value, `${at}.cells[${cellIndex}]`);
      string(cell.ref, `${at}.cells[${cellIndex}].ref`); string(cell.value, `${at}.cells[${cellIndex}].value`);
      for (const [runIndex, run] of optionalList(cell.runs, `${at}.cells[${cellIndex}].runs`).entries()) string(record(run, "run").text, `${at}.cells[${cellIndex}].runs[${runIndex}].text`);
    }
    for (const [drawingIndex, value] of optionalList(sheet.drawings, `${at}.drawings`).entries()) {
      const drawing = record(value, `${at}.drawings[${drawingIndex}]`), where = `${at}.drawings[${drawingIndex}]`;
      oneOf(drawing.kind, ["image", "shapes", "slicer"], `${where}.kind`);
      record(drawing.anchor, `${where}.anchor`);
      if (drawing.kind === "image") checkImage(drawing, where);
      if (drawing.kind === "shapes") { record(drawing.frame, `${where}.frame`); checkElements(drawing.elements, `${where}.elements`); }
      if (drawing.kind === "slicer") for (const [itemIndex, slicerItem] of list(drawing.items, `${where}.items`).entries()) string(record(slicerItem, "slicer item").text, `${where}.items[${itemIndex}].text`);
    }
    if (sheet.conditional !== undefined) {
      for (const [ref, value] of Object.entries(record(sheet.conditional, `${at}.conditional`))) {
        const entry = record(value, `${at}.conditional.${ref}`);
        if (entry.icon !== undefined) { const icon = record(entry.icon, "icon"); string(icon.set, `${at}.conditional.${ref}.icon.set`); number(icon.index, `${at}.conditional.${ref}.icon.index`); }
      }
    }
  }
}
