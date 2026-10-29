/** Small Excel formula evaluator used for conditional formatting and for formula cells saved without cached values. */
import { dateToSerial, formatGeneral, serialToDate } from "./xlsx-number-format.js";

export class FormulaError { constructor(readonly code: string) {} }
export type FormulaScalar = number | string | boolean | FormulaError | null;
export type FormulaRange = { kind: "range"; sheet: string; top: number; left: number; bottom: number; right: number };
export type FormulaValue = FormulaScalar | FormulaRange;
export interface FormulaContext {
  readonly sheet: string;
  readonly row: number;
  readonly column: number;
  cell(sheet: string, row: number, column: number): FormulaScalar;
  name(name: string, sheet: string): string | undefined;
  readonly today?: number;
}

type Ref = { sheet?: string; row: number; column: number; absRow: boolean; absColumn: boolean };
type Node =
  | { k: "num"; v: number } | { k: "str"; v: string } | { k: "bool"; v: boolean } | { k: "err"; v: string }
  | { k: "ref"; from: Ref; to?: Ref; wholeColumns?: boolean; wholeRows?: boolean }
  | { k: "name"; v: string; sheet?: string } | { k: "call"; name: string; args: Node[] }
  | { k: "unary"; op: string; arg: Node } | { k: "bin"; op: string; left: Node; right: Node } | { k: "pct"; arg: Node } | { k: "array"; rows: Node[][] };

const MAX_ROW = 1_048_576, MAX_COLUMN = 16_384;
const errors = { value: new FormulaError("#VALUE!"), div0: new FormulaError("#DIV/0!"), ref: new FormulaError("#REF!"), name: new FormulaError("#NAME?"), na: new FormulaError("#N/A"), num: new FormulaError("#NUM!") };
const astCache = new Map<string, Node | FormulaError>();

export function parseFormula(source: string): Node | FormulaError {
  const cached = astCache.get(source);
  if (cached) return cached;
  let result: Node | FormulaError;
  try { result = new Parser(source.startsWith("=") ? source.slice(1) : source).parse(); }
  catch { result = errors.name; }
  if (astCache.size > 5000) astCache.clear();
  astCache.set(source, result);
  return result;
}

/** Evaluates `formula` for the context cell. Relative references are shifted by (context cell - origin), as for conditional formats. */
export function evaluateFormula(formula: string, context: FormulaContext, origin: { row: number; column: number } = context): FormulaValue {
  const ast = parseFormula(formula);
  if (ast instanceof FormulaError) return ast;
  return new Evaluator(context, context.row - origin.row, context.column - origin.column, 0).eval(ast);
}

export function isTruthy(value: FormulaValue): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  return false;
}

class Parser {
  private at = 0;
  constructor(private readonly source: string) {}
  parse(): Node {
    const node = this.comparison();
    this.space();
    if (this.at < this.source.length) throw new Error("Unexpected input");
    return node;
  }
  private space() { while (/\s/.test(this.source[this.at] ?? "")) this.at++; }
  private peek(text: string) { this.space(); return this.source.startsWith(text, this.at); }
  private take(text: string) { if (this.peek(text)) { this.at += text.length; return true; } return false; }
  private comparison(): Node {
    let left = this.concat();
    for (;;) {
      const op = ["<>", "<=", ">=", "=", "<", ">"].find(candidate => this.peek(candidate));
      if (!op) return left;
      this.at += op.length; left = { k: "bin", op, left, right: this.concat() };
    }
  }
  private concat(): Node { let left = this.additive(); while (this.take("&")) left = { k: "bin", op: "&", left, right: this.additive() }; return left; }
  private additive(): Node {
    let left = this.multiplicative();
    for (;;) { const op = this.peek("+") ? "+" : this.peek("-") ? "-" : undefined; if (!op) return left; this.at++; left = { k: "bin", op, left, right: this.multiplicative() }; }
  }
  private multiplicative(): Node {
    let left = this.power();
    for (;;) { const op = this.peek("*") ? "*" : this.peek("/") ? "/" : undefined; if (!op) return left; this.at++; left = { k: "bin", op, left, right: this.power() }; }
  }
  private power(): Node { let left = this.unary(); while (this.take("^")) left = { k: "bin", op: "^", left, right: this.unary() }; return left; }
  private unary(): Node {
    if (this.take("-")) return { k: "unary", op: "-", arg: this.unary() };
    if (this.take("+")) return this.unary();
    let node = this.primary();
    while (this.take("%")) node = { k: "pct", arg: node };
    return node;
  }
  private primary(): Node {
    this.space();
    const rest = this.source.slice(this.at);
    if (this.take("(")) { const node = this.comparison(); if (!this.take(")")) throw new Error("Missing )"); return node; }
    if (this.take("{")) {
      const rows: Node[][] = [[]];
      if (!this.peek("}")) for (;;) {
        rows.at(-1)!.push(this.unary());
        if (this.take(",")) continue;
        if (this.take(";")) { rows.push([]); continue; }
        break;
      }
      if (!this.take("}")) throw new Error("Missing }");
      return { k: "array", rows };
    }
    if (rest.startsWith('"')) {
      let value = "", index = 1;
      for (;;) {
        const quote = rest.indexOf('"', index);
        if (quote < 0) throw new Error("Unclosed string");
        value += rest.slice(index, quote);
        if (rest[quote + 1] === '"') { value += '"'; index = quote + 2; continue; }
        this.at += quote + 1; return { k: "str", v: value };
      }
    }
    const number = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/.exec(rest);
    if (number && !/^\d+:/.test(rest)) { this.at += number[0].length; return { k: "num", v: Number(number[0]) }; }
    const error = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!)/i.exec(rest);
    if (error) { this.at += error[0].length; return { k: "err", v: error[0].toUpperCase() }; }
    // Optional sheet prefix: 'Sheet name'! or Sheet1!
    let sheet: string | undefined, prefixLength = 0;
    const quoted = /^'((?:[^']|'')+)'!/.exec(rest), bare = /^([\p{L}\p{N}_.]+)!/u.exec(rest);
    if (quoted) { sheet = quoted[1]!.replace(/''/g, "'"); prefixLength = quoted[0].length; }
    else if (bare) { sheet = bare[1]!; prefixLength = bare[0].length; }
    const body = rest.slice(prefixLength);
    const cell = "(\\$?)([A-Za-z]{1,3})(\\$?)(\\d+)";
    const range = new RegExp(`^${cell}(?::${cell})?(?![\\p{L}\\p{N}_.(])`, "u").exec(body);
    if (range) {
      const make = (absColumn: string, column: string, absRow: string, row: string): Ref => ({ ...(sheet ? { sheet } : {}), column: columnNumber(column), row: Number(row), absColumn: !!absColumn, absRow: !!absRow });
      const from = make(range[1]!, range[2]!, range[3]!, range[4]!);
      if (from.column <= MAX_COLUMN && from.row <= MAX_ROW && from.row > 0) {
        this.at += prefixLength + range[0].length;
        return range[5] !== undefined ? { k: "ref", from, to: make(range[5]!, range[6]!, range[7]!, range[8]!) } : { k: "ref", from };
      }
    }
    const columns = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\p{L}\p{N}_.(])/u.exec(body);
    if (columns) {
      this.at += prefixLength + columns[0].length;
      return { k: "ref", wholeColumns: true, from: { ...(sheet ? { sheet } : {}), column: columnNumber(columns[2]!), row: 1, absColumn: !!columns[1], absRow: true }, to: { column: columnNumber(columns[4]!), row: MAX_ROW, absColumn: !!columns[3], absRow: true } };
    }
    const rows = /^(\$?)(\d+):(\$?)(\d+)/.exec(body);
    if (rows) {
      this.at += prefixLength + rows[0].length;
      return { k: "ref", wholeRows: true, from: { ...(sheet ? { sheet } : {}), column: 1, row: Number(rows[2]), absColumn: true, absRow: !!rows[1] }, to: { column: MAX_COLUMN, row: Number(rows[4]), absColumn: true, absRow: !!rows[3] } };
    }
    const identifier = /^[\p{L}_\\][\p{L}\p{N}_.\\?]*/u.exec(body);
    if (identifier) {
      this.at += prefixLength + identifier[0].length;
      const upper = identifier[0].toUpperCase();
      if (this.peek("(")) {
        this.at++;
        const args: Node[] = [];
        if (!this.peek(")")) for (;;) {
          if (this.peek(",") || this.peek(")")) args.push({ k: "str", v: "" }); else args.push(this.comparison());
          if (this.take(",")) continue;
          break;
        }
        if (!this.take(")")) throw new Error("Missing ) in call");
        return { k: "call", name: upper.replace(/^_XLFN\.|^_XLWS\.|^_XLL\./, ""), args };
      }
      if (upper === "TRUE" || upper === "FALSE") return { k: "bool", v: upper === "TRUE" };
      return { k: "name", v: identifier[0], ...(sheet ? { sheet } : {}) };
    }
    throw new Error(`Unexpected token at ${this.at}`);
  }
}

function columnNumber(letters: string): number { let value = 0; for (const char of letters.toUpperCase()) value = value * 26 + char.charCodeAt(0) - 64; return value; }

class Evaluator {
  constructor(private readonly context: FormulaContext, private readonly dRow: number, private readonly dColumn: number, private readonly depth: number) {}

  eval(node: Node): FormulaValue {
    switch (node.k) {
      case "num": return node.v;
      case "str": return node.v;
      case "bool": return node.v;
      case "err": return new FormulaError(node.v);
      case "array": return this.eval(node.rows[0]?.[0] ?? { k: "num", v: 0 });
      case "ref": return this.reference(node);
      case "name": return this.name(node.v, node.sheet);
      case "pct": { const value = toNumber(this.scalar(node.arg)); return value instanceof FormulaError ? value : value / 100; }
      case "unary": { const value = toNumber(this.scalar(node.arg)); return value instanceof FormulaError ? value : -value; }
      case "bin": return this.binary(node.op, this.scalar(node.left), this.scalar(node.right));
      case "call": return this.call(node.name, node.args);
    }
  }

  private shift(ref: Ref): { row: number; column: number } {
    return { row: ref.absRow ? ref.row : ref.row + this.dRow, column: ref.absColumn ? ref.column : ref.column + this.dColumn };
  }

  private reference(node: Extract<Node, { k: "ref" }>): FormulaValue {
    const sheet = node.from.sheet ?? this.context.sheet;
    const from = this.shift(node.from), to = node.to ? this.shift(node.to) : from;
    if (node.wholeColumns) { from.row = 1; to.row = MAX_ROW; }
    if (node.wholeRows) { from.column = 1; to.column = MAX_COLUMN; }
    if ([from.row, to.row].some(row => row < 1 || row > MAX_ROW) || [from.column, to.column].some(column => column < 1 || column > MAX_COLUMN)) return errors.ref;
    if (!node.to) return this.context.cell(sheet, from.row, from.column);
    return { kind: "range", sheet, top: Math.min(from.row, to.row), left: Math.min(from.column, to.column), bottom: Math.max(from.row, to.row), right: Math.max(from.column, to.column) };
  }

  private name(name: string, sheet?: string): FormulaValue {
    if (this.depth > 16) return errors.ref;
    const formula = this.context.name(name, sheet ?? this.context.sheet);
    if (formula === undefined) return errors.name;
    const ast = parseFormula(formula);
    if (ast instanceof FormulaError) return ast;
    // Defined names store relative references relative to A1; they resolve against the cell being evaluated.
    return new Evaluator(this.context, this.context.row - 1, this.context.column - 1, this.depth + 1).eval(ast);
  }

  /** Implicit intersection: a range used where a single value is expected picks the cell in the current row/column. */
  private scalar(node: Node): FormulaScalar { return this.intersect(this.eval(node)); }
  private intersect(value: FormulaValue): FormulaScalar {
    if (!isRange(value)) return value;
    if (value.top === value.bottom && value.left === value.right) return this.context.cell(value.sheet, value.top, value.left);
    if (value.left === value.right && this.context.row >= value.top && this.context.row <= value.bottom) return this.context.cell(value.sheet, this.context.row, value.left);
    if (value.top === value.bottom && this.context.column >= value.left && this.context.column <= value.right) return this.context.cell(value.sheet, value.top, this.context.column);
    return errors.value;
  }

  private binary(op: string, left: FormulaScalar, right: FormulaScalar): FormulaScalar {
    if (left instanceof FormulaError) return left;
    if (right instanceof FormulaError) return right;
    if (op === "&") return toText(left) + toText(right);
    if (["=", "<>", "<", ">", "<=", ">="].includes(op)) {
      const order = compare(left, right);
      return op === "=" ? order === 0 : op === "<>" ? order !== 0 : op === "<" ? order < 0 : op === ">" ? order > 0 : op === "<=" ? order <= 0 : order >= 0;
    }
    const a = toNumber(left), b = toNumber(right);
    if (a instanceof FormulaError) return a;
    if (b instanceof FormulaError) return b;
    if (op === "+") return a + b;
    if (op === "-") return a - b;
    if (op === "*") return a * b;
    if (op === "/") return b === 0 ? errors.div0 : a / b;
    if (op === "^") { const result = a ** b; return Number.isFinite(result) ? result : errors.num; }
    return errors.value;
  }

  /** Flattens arguments for aggregate functions: ranges contribute numbers only, direct scalars are coerced. */
  private numbers(args: Node[]): number[] | FormulaError {
    const result: number[] = [];
    for (const arg of args) {
      const value = this.eval(arg);
      if (isRange(value)) { for (const cell of this.cells(value)) { if (cell instanceof FormulaError) return cell; if (typeof cell === "number") result.push(cell); } continue; }
      if (value === null) continue;
      const number = toNumber(value);
      if (number instanceof FormulaError) return number;
      result.push(number);
    }
    return result;
  }
  private *cells(range: FormulaRange): Generator<FormulaScalar> {
    const bottom = Math.min(range.bottom, range.top + 100_000), right = Math.min(range.right, range.left + 1000);
    for (let row = range.top; row <= bottom; row++) for (let column = range.left; column <= right; column++) yield this.context.cell(range.sheet, row, column);
  }
  private values(args: Node[]): FormulaScalar[] {
    const result: FormulaScalar[] = [];
    for (const arg of args) { const value = this.eval(arg); if (isRange(value)) result.push(...this.cells(value)); else result.push(value); }
    return result;
  }

  private call(name: string, args: Node[]): FormulaValue {
    const arg = (index: number): FormulaScalar => args[index] ? this.scalar(args[index]!) : null;
    const num = (index: number, fallback?: number): number | FormulaError => args[index] === undefined && fallback !== undefined ? fallback : toNumber(arg(index));
    const aggregate = (fn: (values: number[]) => FormulaScalar) => { const values = this.numbers(args); return values instanceof FormulaError ? values : fn(values); };
    const math = (fn: (...values: number[]) => FormulaScalar, count = 1) => {
      const values: number[] = [];
      for (let index = 0; index < count; index++) { const value = num(index); if (value instanceof FormulaError) return value; values.push(value); }
      return fn(...values);
    };
    switch (name) {
      case "SUM": return aggregate(values => values.reduce((sum, value) => sum + value, 0));
      case "AVERAGE": return aggregate(values => values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : errors.div0);
      case "MIN": return aggregate(values => values.length ? Math.min(...values) : 0);
      case "MAX": return aggregate(values => values.length ? Math.max(...values) : 0);
      case "PRODUCT": return aggregate(values => values.reduce((product, value) => product * value, 1));
      case "MEDIAN": return aggregate(values => { if (!values.length) return errors.num; const sorted = [...values].sort((a, b) => a - b), middle = sorted.length >> 1; return sorted.length % 2 ? sorted[middle]! : (sorted[middle - 1]! + sorted[middle]!) / 2; });
      case "COUNT": return this.values(args).filter(value => typeof value === "number").length;
      case "COUNTA": return this.values(args).filter(value => value !== null && value !== "").length;
      case "COUNTBLANK": return this.values(args).filter(value => value === null || value === "").length;
      case "ABS": return math(Math.abs);
      case "INT": return math(Math.floor);
      case "TRUNC": return math((value, digits) => Math.trunc(value * 10 ** (digits ?? 0)) / 10 ** (digits ?? 0), args.length > 1 ? 2 : 1);
      case "SIGN": return math(Math.sign);
      case "SQRT": return math(value => value < 0 ? errors.num : Math.sqrt(value));
      case "POWER": return math((base, exponent) => base ** exponent!, 2);
      case "PI": return Math.PI;
      case "MOD": return math((value, divisor) => divisor === 0 ? errors.div0 : value - divisor! * Math.floor(value / divisor!), 2);
      case "ROUND": return math((value, digits) => roundHalfAway(value, digits ?? 0), 2);
      case "ROUNDUP": return math((value, digits) => { const factor = 10 ** (digits ?? 0); return Math.sign(value) * Math.ceil(Math.abs(value) * factor - 1e-9) / factor; }, 2);
      case "ROUNDDOWN": return math((value, digits) => { const factor = 10 ** (digits ?? 0); return Math.sign(value) * Math.floor(Math.abs(value) * factor + 1e-9) / factor; }, 2);
      case "CEILING": case "CEILING.MATH": return math((value, step) => !step ? 0 : Math.ceil(value / step) * step, args.length > 1 ? 2 : 1);
      case "FLOOR": case "FLOOR.MATH": return math((value, step) => !step ? 0 : Math.floor(value / step) * step, args.length > 1 ? 2 : 1);
      case "AND": case "OR": case "XOR": {
        const values = this.values(args).filter(value => value !== null && typeof value !== "string");
        for (const value of values) if (value instanceof FormulaError) return value;
        const flags = values.map(value => isTruthy(value as FormulaValue));
        if (!flags.length) return errors.value;
        return name === "AND" ? flags.every(Boolean) : name === "OR" ? flags.some(Boolean) : flags.filter(Boolean).length % 2 === 1;
      }
      case "NOT": { const value = arg(0); if (value instanceof FormulaError) return value; const number = toNumber(value); return number instanceof FormulaError ? number : number === 0; }
      case "TRUE": return true;
      case "FALSE": return false;
      case "IF": { const condition = arg(0); if (condition instanceof FormulaError) return condition; const truth = typeof condition === "string" ? errors.value : isTruthy(condition); if (truth instanceof FormulaError) return truth; return truth ? (args[1] ? this.eval(args[1]) : true) : (args[2] ? this.eval(args[2]) : false); }
      case "IFERROR": { const value = this.eval(args[0]!); return value instanceof FormulaError ? (args[1] ? this.eval(args[1]) : 0) : value; }
      case "IFNA": { const value = this.eval(args[0]!); return value instanceof FormulaError && value.code === "#N/A" ? (args[1] ? this.eval(args[1]) : 0) : value; }
      case "IFS": { for (let index = 0; index + 1 < args.length; index += 2) { const condition = this.scalar(args[index]!); if (condition instanceof FormulaError) return condition; if (isTruthy(condition)) return this.eval(args[index + 1]!); } return errors.na; }
      case "ISBLANK": return arg(0) === null;
      case "ISNUMBER": return typeof arg(0) === "number";
      case "ISTEXT": return typeof arg(0) === "string";
      case "ISNONTEXT": return typeof arg(0) !== "string";
      case "ISLOGICAL": return typeof arg(0) === "boolean";
      case "ISERROR": return arg(0) instanceof FormulaError;
      case "ISERR": { const value = arg(0); return value instanceof FormulaError && value.code !== "#N/A"; }
      case "ISNA": { const value = arg(0); return value instanceof FormulaError && value.code === "#N/A"; }
      case "ISEVEN": return math(value => Math.floor(Math.abs(value)) % 2 === 0);
      case "ISODD": return math(value => Math.floor(Math.abs(value)) % 2 === 1);
      case "NA": return errors.na;
      case "N": { const value = arg(0); return typeof value === "number" ? value : typeof value === "boolean" ? Number(value) : value instanceof FormulaError ? value : 0; }
      case "ROW": case "COLUMN": {
        if (!args.length) return name === "ROW" ? this.context.row : this.context.column;
        const value = this.eval(args[0]!);
        if (isRange(value)) return name === "ROW" ? value.top : value.left;
        const ref = args[0]!.k === "ref" ? this.shift((args[0] as Extract<Node, { k: "ref" }>).from) : undefined;
        return ref ? (name === "ROW" ? ref.row : ref.column) : errors.value;
      }
      case "ROWS": case "COLUMNS": { const value = this.eval(args[0]!); return isRange(value) ? (name === "ROWS" ? value.bottom - value.top + 1 : value.right - value.left + 1) : 1; }
      case "LEN": return toText(arg(0)).length;
      case "LEFT": { const count = num(1, 1); return count instanceof FormulaError ? count : toText(arg(0)).slice(0, Math.max(0, count)); }
      case "RIGHT": { const count = num(1, 1); return count instanceof FormulaError ? count : count <= 0 ? "" : toText(arg(0)).slice(-count); }
      case "MID": { const start = num(1), count = num(2); return start instanceof FormulaError ? start : count instanceof FormulaError ? count : toText(arg(0)).slice(start - 1, start - 1 + count); }
      case "UPPER": return toText(arg(0)).toUpperCase();
      case "LOWER": return toText(arg(0)).toLowerCase();
      case "PROPER": return toText(arg(0)).toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_all, before: string, letter: string) => before + letter.toUpperCase());
      case "TRIM": return toText(arg(0)).trim().replace(/ +/g, " ");
      case "CONCATENATE": case "CONCAT": return this.values(args).map(toText).join("");
      case "TEXTJOIN": { const delimiter = toText(arg(0)), skip = isTruthy(arg(1)); return this.values(args.slice(2)).map(toText).filter(text => !skip || text !== "").join(delimiter); }
      case "REPT": { const count = num(1); return count instanceof FormulaError ? count : toText(arg(0)).repeat(Math.max(0, Math.min(32767, count))); }
      case "EXACT": return toText(arg(0)) === toText(arg(1));
      case "FIND": case "SEARCH": {
        const needle = toText(arg(0)), haystack = toText(arg(1)), start = num(2, 1);
        if (start instanceof FormulaError) return start;
        const index = name === "FIND" ? haystack.indexOf(needle, start - 1) : haystack.toLowerCase().indexOf(needle.toLowerCase(), start - 1);
        return index < 0 ? errors.value : index + 1;
      }
      case "SUBSTITUTE": return toText(arg(0)).split(toText(arg(1))).join(toText(arg(2)));
      case "VALUE": return toNumber(arg(0));
      case "TODAY": return this.context.today ?? todaySerial();
      case "NOW": return this.context.today ?? todaySerial();
      case "DATE": return math((year, month, day) => { const fullYear = year < 1900 ? year + 1900 : year; return dateToSerial(fullYear, month! - 1, day!); }, 3);
      case "YEAR": case "MONTH": case "DAY": case "WEEKDAY": {
        const serial = num(0);
        if (serial instanceof FormulaError) return serial;
        const date = serialToDate(Math.floor(serial));
        if (name === "YEAR") return date.year;
        if (name === "MONTH") return date.month + 1;
        if (name === "DAY") return date.day;
        const type = num(1, 1);
        if (type instanceof FormulaError) return type;
        return type === 2 ? (date.weekday + 6) % 7 + 1 : type === 3 ? (date.weekday + 6) % 7 : date.weekday + 1;
      }
      case "EOMONTH": case "EDATE": return math((serial, months) => {
        const date = serialToDate(Math.floor(serial));
        const target = new Date(Date.UTC(date.year, date.month + Math.trunc(months!) + (name === "EOMONTH" ? 1 : 0), name === "EOMONTH" ? 0 : date.day));
        return dateToSerial(target.getUTCFullYear(), target.getUTCMonth(), target.getUTCDate());
      }, 2);
      case "COUNTIF": case "SUMIF": case "AVERAGEIF": {
        const range = this.eval(args[0]!), criteria = arg(1), sumRange = args[2] ? this.eval(args[2]) : range;
        if (!isRange(range)) return errors.value;
        const matches = criterion(criteria);
        const cells = [...this.cells(range)], sums = isRange(sumRange) ? [...this.cells({ ...sumRange, bottom: sumRange.top + range.bottom - range.top, right: sumRange.left + range.right - range.left })] : cells;
        const hits = cells.map(matches);
        if (name === "COUNTIF") return hits.filter(Boolean).length;
        const values = sums.filter((value, index) => hits[index] && typeof value === "number") as number[];
        if (name === "SUMIF") return values.reduce((sum, value) => sum + value, 0);
        return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : errors.div0;
      }
      case "CHOOSE": { const index = num(0); return index instanceof FormulaError ? index : args[index] ? this.eval(args[index]!) : errors.value; }
      case "TEXT": { const value = arg(0); return typeof value === "number" ? formatGeneral(value) : toText(value); }
      default: return errors.name;
    }
  }
}

function isRange(value: FormulaValue): value is FormulaRange { return typeof value === "object" && value !== null && !(value instanceof FormulaError) && value.kind === "range"; }

function toNumber(value: FormulaScalar): number | FormulaError {
  if (value instanceof FormulaError) return value;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (value === null || value === "") return 0;
  const number = Number(value.trim().replace(/%$/, ""));
  return Number.isFinite(number) && value.trim() !== "" ? (value.trim().endsWith("%") ? number / 100 : number) : errors.value;
}

function toText(value: FormulaScalar): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return formatGeneral(value);
  if (value instanceof FormulaError) return value.code;
  return value;
}

/** Excel ordering for comparisons: numbers < text < logicals; empty cells compare as 0 or "". */
function compare(left: FormulaScalar, right: FormulaScalar): number {
  if (left === null) left = typeof right === "string" ? "" : typeof right === "boolean" ? false : 0;
  if (right === null) right = typeof left === "string" ? "" : typeof left === "boolean" ? false : 0;
  const rank = (value: FormulaScalar) => typeof value === "number" ? 0 : typeof value === "string" ? 1 : 2;
  if (rank(left) !== rank(right)) return rank(left) - rank(right);
  if (typeof left === "string" && typeof right === "string") { const a = left.toLowerCase(), b = right.toLowerCase(); return a < b ? -1 : a > b ? 1 : 0; }
  const a = Number(left), b = Number(right);
  return Math.abs(a - b) < 1e-12 * Math.max(1, Math.abs(a), Math.abs(b)) ? 0 : a < b ? -1 : 1;
}

function criterion(value: FormulaScalar): (cell: FormulaScalar) => boolean {
  if (typeof value !== "string") return cell => compare(cell, value) === 0 && cell !== null;
  const match = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(value)!, op = match[1] ?? "=", operand = match[2]!;
  const numeric = operand.trim() !== "" && Number.isFinite(Number(operand)) ? Number(operand) : undefined;
  const pattern = numeric === undefined && /[*?]/.test(operand) ? new RegExp(`^${operand.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".")}$`, "is") : undefined;
  return cell => {
    if (pattern && (op === "=" || op === "<>")) { const hit = typeof cell === "string" && pattern.test(cell); return op === "=" ? hit : !hit; }
    const target: FormulaScalar = numeric ?? operand;
    if (op === "=" && operand === "") return cell === null || cell === "";
    if (numeric !== undefined && typeof cell !== "number") return op === "<>";
    const order = compare(cell, target);
    return op === "=" ? order === 0 : op === "<>" ? order !== 0 : op === "<" ? order < 0 : op === ">" ? order > 0 : op === "<=" ? order <= 0 : order >= 0;
  };
}

function roundHalfAway(value: number, digits: number): number {
  const factor = 10 ** digits;
  return Math.sign(value) * Math.round(Math.abs(value) * factor + 1e-9) / factor;
}

function todaySerial(): number {
  const now = new Date();
  return dateToSerial(now.getFullYear(), now.getMonth(), now.getDate());
}
