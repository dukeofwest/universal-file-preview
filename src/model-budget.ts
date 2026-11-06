import { PreviewError } from "./errors.js";

/**
 * Limits on the size of a parsed document model, enforced while the model is built (ModelBudget) and checked
 * again on the finished model (checkModel). Internal module (not exported from the package entry point).
 *
 * The accounting follows what structured cloning copies when a model crosses threads: strings count once per
 * occurrence (cloning writes every occurrence of a primitive string); typed arrays count once per identity, and
 * so do objects and arrays that contain objects (cloning preserves identity, so a shared subtree is copied once).
 * Leaf objects (whose fields are all primitives, such as a cell or a run) count once per reference: tracking the
 * identity of millions of leaves would cost more memory than the model itself, a leaf cannot be part of a cycle,
 * and counting a shared leaf again only overestimates. ModelBudget and checkModel use the same rule, so the
 * budget charged during parsing is never lower than what the walk measures. A small file can therefore not
 * produce a model that is expensive to keep, clone or render, in either execution mode.
 */
export interface ModelLimits {
  /** Spreadsheet cells across the workbook (XLSX and OpenDocument). */
  readonly cells: number;
  /** Objects and arrays. Target: about 512 MiB of JavaScript heap at about 75 bytes per object (measured). */
  readonly objects: number;
  /** UTF-16 code units of all strings, counted per occurrence. */
  readonly stringUnits: number;
  /** Code units of any one string (cell text has its own, lower limit of 32,767 in the parsers). */
  readonly stringLength: number;
  /** Bytes of all distinct Uint8Arrays (embedded images), the same as the package decompression budget. */
  readonly binaryBytes: number;
  /** Nesting depth of objects and arrays. */
  readonly depth: number;
}

export const MODEL_LIMITS: ModelLimits = {
  cells: 500_000,
  objects: 6_000_000,
  stringUnits: 64 * 1024 * 1024,
  stringLength: 1024 * 1024,
  binaryBytes: 256 * 1024 * 1024,
  depth: 64,
};

export interface ModelUsage { readonly cells: number; readonly objects: number; readonly stringUnits: number; readonly binaryBytes: number }

const exceeded = (what: string, limit: number) => new PreviewError(`The document exceeds the ${what} limit (${limit.toLocaleString("en-US")}).`, "resource-limit");

/** The fields (or elements) of a plain object or array. */
function fieldsOf(value: object): unknown[] {
  if (Array.isArray(value)) return value;
  const record = value as Record<string, unknown>, keys = Object.keys(record), fields = new Array<unknown>(keys.length);
  for (let index = 0; index < keys.length; index++) fields[index] = record[keys[index]!];
  return fields;
}
/** True when every field is a primitive: such an object is counted per reference, without identity tracking. */
function isLeaf(fields: readonly unknown[]): boolean {
  for (const field of fields) if (typeof field === "object" && field !== null) return false;
  return true;
}

/** Charged by a parser while it builds one model; throws before the model grows past a limit. */
export class ModelBudget {
  #cells = 0;
  #objects = 0;
  #stringUnits = 0;
  #binaryBytes = 0;
  // Sets, not WeakSets: the budget lives only for one parse, and WeakSets with millions of entries make garbage
  // collection very slow (measured: 4,000,000 objects took 84 s in a WeakSet, 2 s in a Set).
  readonly #binary = new Set<Uint8Array>();
  readonly #objectsSeen = new Set<object>();

  constructor(readonly limits: ModelLimits = MODEL_LIMITS) {}

  /** Spreadsheet cells added to the model. */
  cells(count = 1): void {
    if ((this.#cells += count) > this.limits.cells) throw new PreviewError(`The spreadsheet has more than ${this.limits.cells.toLocaleString("en-US")} cells.`, "resource-limit");
  }

  /** New objects or arrays added to the model. */
  objects(count = 1): void {
    if ((this.#objects += count) > this.limits.objects) throw exceeded("object", this.limits.objects);
  }

  /** One occurrence of a string placed into the model. Returns the string, for use in expressions. */
  text(value: string): string {
    if (value.length > this.limits.stringLength) throw exceeded("string length", this.limits.stringLength);
    if ((this.#stringUnits += value.length) > this.limits.stringUnits) throw exceeded("text", this.limits.stringUnits);
    return value;
  }

  /** Binary data placed into the model; the same array placed several times is counted once (cloning copies it once). */
  binary(bytes: Uint8Array): Uint8Array {
    if (this.#binary.has(bytes)) return bytes;
    this.#binary.add(bytes);
    if ((this.#binaryBytes += bytes.byteLength) > this.limits.binaryBytes) throw exceeded("binary data", this.limits.binaryBytes);
    return bytes;
  }

  /**
   * A part of the model, charged as checkModel counts it: strings per occurrence, Uint8Arrays once, leaf objects
   * per reference, other objects once (a subtree charged before, for example a shared runs array or a sheet, is
   * skipped). Property names are not charged (the parsers charge names they create from document content).
   * Depth and value types are left to checkModel.
   */
  add(value: unknown): void {
    const stack: unknown[] = [value];
    while (stack.length) {
      const item = stack.pop();
      if (typeof item === "string") { this.text(item); continue; }
      if (typeof item !== "object" || item === null) continue;
      if (item instanceof Uint8Array) { this.binary(item); continue; }
      const fields = fieldsOf(item);
      if (isLeaf(fields)) { this.objects(); for (const field of fields) if (typeof field === "string") this.text(field); continue; }
      if (this.#objectsSeen.has(item)) continue;
      this.#objectsSeen.add(item);
      this.objects();
      for (const field of fields) stack.push(field);
    }
  }

  /**
   * A record the parser has just filled in (a cell, a conditional-format entry), referenced once by its
   * container: the record and its fields are charged as add() would. A record changed later is charged again,
   * which only overestimates. Its container is then charged with container(), which does not count it again.
   */
  record(value: object): void {
    this.objects();
    for (const key of Object.keys(value)) this.add((value as Record<string, unknown>)[key]);
  }

  /**
   * A finished container (a sheet): the container is charged once and marked, so later add() calls on the
   * whole model skip it. Fields named in `recorded` hold records charged with record(): only the array or
   * object holding them is counted here. Other fields are charged with add().
   */
  container(value: object, recorded: readonly string[]): void {
    if (this.#objectsSeen.has(value)) return;
    this.#objectsSeen.add(value);
    this.objects();
    for (const key of Object.keys(value)) {
      const field = (value as Record<string, unknown>)[key];
      if (recorded.includes(key) && typeof field === "object" && field !== null) { this.#objectsSeen.add(field); this.objects(); }
      else this.add(field);
    }
  }

  get usage(): ModelUsage { return { cells: this.#cells, objects: this.#objects, stringUnits: this.#stringUnits, binaryBytes: this.#binaryBytes }; }
}

let backstopRejections = 0, budgetMisses = 0;
/** How often checkModel has rejected a model. The parsers' budgets are meant to stop oversized models first. */
export function modelBackstopRejections(): number { return backstopRejections; }
/**
 * How often a finished model measured more than its parser charged (a charge point missing in a parser).
 * Diagnostic for tests: the backstop still enforces the limits.
 */
export function modelBudgetMisses(): number { return budgetMisses; }

/**
 * Checks a finished model against the limits with one iterative walk: only plain objects, arrays, strings,
 * finite numbers, booleans, null, undefined and Uint8Array; no cycles; bounded depth; and the object, string
 * and binary limits counted as described above. Throws PreviewError("resource-limit") or
 * PreviewError("invalid-file"); returns the measured usage. `charged` is what the parser's ModelBudget
 * counted for the same model; measuring more than that counts as a budget miss (modelBudgetMisses).
 */
export function checkModel(model: unknown, limits: ModelLimits = MODEL_LIMITS, charged?: ModelUsage): ModelUsage {
  let usage: ModelUsage;
  try { usage = walk(model, limits); }
  catch (error) { backstopRejections++; throw error; }
  if (charged && (usage.objects > charged.objects || usage.stringUnits > charged.stringUnits || usage.binaryBytes > charged.binaryBytes)) budgetMisses++;
  return usage;
}

function walk(model: unknown, limits: ModelLimits): ModelUsage {
  let objects = 0, stringUnits = 0, binaryBytes = 0;
  // Identity is tracked only for typed arrays and for objects that contain objects (see the module comment).
  const seen = new Set<object>(), onPath = new Set<object>();
  // Parallel stacks of objects still to visit and their depths; depth -1 marks the end of an object's children.
  const values: object[] = [], depths: number[] = [];
  const string = (value: string) => {
    if (value.length > limits.stringLength) throw exceeded("string length", limits.stringLength);
    if ((stringUnits += value.length) > limits.stringUnits) throw exceeded("text", limits.stringUnits);
  };
  /** A field or element: primitives are checked here, objects are queued. */
  const visit = (value: unknown, depth: number) => {
    if (typeof value === "string") string(value);
    else if (typeof value === "number") { if (!Number.isFinite(value)) throw new PreviewError("The document model contains a non-finite number.", "invalid-file"); }
    else if (typeof value === "object" && value !== null) { values.push(value); depths.push(depth); }
    else if (value !== null && value !== undefined && typeof value !== "boolean") throw new PreviewError(`The document model contains a ${typeof value}.`, "invalid-file");
  };
  visit(model, 0);
  while (values.length) {
    const value = values.pop()!, depth = depths.pop()!;
    if (depth < 0) { onPath.delete(value); continue; }
    if (value instanceof Uint8Array) {
      if (seen.has(value)) continue; // copied once by structured cloning
      seen.add(value);
      if ((binaryBytes += value.byteLength) > limits.binaryBytes) throw exceeded("binary data", limits.binaryBytes);
      continue;
    }
    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== Array.prototype && prototype !== null) throw new PreviewError("The document model contains an unexpected object type.", "invalid-file");
    if (depth + 1 > limits.depth) throw exceeded("nesting depth", limits.depth);
    const fields = fieldsOf(value);
    if (!isLeaf(fields)) {
      if (onPath.has(value)) throw new PreviewError("The document model contains a cycle.", "invalid-file");
      if (seen.has(value)) continue; // shared subtree: copied once by structured cloning
      seen.add(value);
      onPath.add(value);
      values.push(value); depths.push(-1);
    }
    if (++objects > limits.objects) throw exceeded("object", limits.objects);
    // Property names are not counted: fixed field names are bounded by the object limit, and parsers charge
    // the names they create from document content (for example cell references used as keys).
    for (let index = fields.length - 1; index >= 0; index--) visit(fields[index], depth + 1);
  }
  return { cells: 0, objects, stringUnits, binaryBytes };
}
