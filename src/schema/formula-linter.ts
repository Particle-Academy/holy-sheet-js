import type { Cell } from "../workbook/cell";
import { CellAddress } from "../workbook/cell-address";
import type { Workbook } from "../workbook/workbook";
import { isNumericString } from "../util";
import { Normalizer } from "./normalizer";
import type { FormulaProblem } from "./types";

const ERR_VALUE = "#VALUE!";
const ERR_REF = "#REF!";
const ERR_NAME = "#NAME?";
const ERR_DIV0 = "#DIV/0!";
const ERR_CIRC = "#CIRC!";

class LinterError extends Error {
  constructor(public errorCode: string) {
    super(errorCode);
  }
}

interface Token {
  type: "NUMBER" | "STRING" | "IDENT" | "OP";
  value: string;
}
interface Cursor {
  tokens: Token[];
  pos: number;
}

type Val = number | string | boolean | null | Val[];

/**
 * Evaluates every formula and reports Excel-style errors. Mirrors PHP
 * `Schema\FormulaLinter` rule-for-rule.
 */
export class FormulaLinter {
  private index = new Map<string, Cell>();
  private cache = new Map<string, Val>();

  lint(schema: unknown): FormulaProblem[] {
    const workbook = new Normalizer().normalize(schema);
    this.index = this.buildIndex(workbook);
    this.cache = new Map();
    const issues: FormulaProblem[] = [];

    for (const sheet of workbook.sheets) {
      for (const [address, cell] of sheet.cells) {
        if (cell.formula === null) continue;
        const key = sheet.name + "!" + address;
        const result = this.evaluate(cell.formula, sheet.name, [key]);
        this.cache.set(key, result);

        if (this.isError(result)) {
          issues.push({
            sheet: sheet.name,
            address,
            formula: cell.formula,
            error: result as string,
            hint: this.hint(result as string, cell.formula, sheet.name),
          });
        }
      }
    }
    return issues;
  }

  private buildIndex(wb: Workbook): Map<string, Cell> {
    const index = new Map<string, Cell>();
    for (const sheet of wb.sheets) {
      for (const [address, cell] of sheet.cells) {
        index.set(sheet.name + "!" + address, cell);
      }
    }
    return index;
  }

  private isError(v: Val): boolean {
    return typeof v === "string" && v.length > 0 && v[0] === "#";
  }

  private evaluate(formula: string, defaultSheet: string, stack: string[]): Val {
    try {
      const c: Cursor = { tokens: this.tokenize(formula), pos: 0 };
      const result = this.parseExpr(c, defaultSheet, stack);
      if (c.pos < c.tokens.length) return ERR_NAME;
      return result;
    } catch (e) {
      if (e instanceof LinterError) return e.errorCode;
      return ERR_NAME;
    }
  }

  // ---- Tokenizer ----

  private tokenize(src: string): Token[] {
    const tokens: Token[] = [];
    let i = 0;
    const len = src.length;
    const isSpace = (ch: string) => /\s/.test(ch);
    const isDigit = (ch: string) => ch >= "0" && ch <= "9";
    const isAlpha = (ch: string) => /[A-Za-z]/.test(ch);
    const isAlnum = (ch: string) => /[A-Za-z0-9]/.test(ch);

    while (i < len) {
      const ch = src[i]!;
      if (isSpace(ch)) {
        i++;
        continue;
      }
      if (isDigit(ch) || (ch === "." && i + 1 < len && isDigit(src[i + 1]!))) {
        const start = i;
        while (i < len && (isDigit(src[i]!) || src[i] === ".")) i++;
        tokens.push({ type: "NUMBER", value: src.slice(start, i) });
        continue;
      }
      if (ch === '"') {
        i++;
        const start = i;
        while (i < len && src[i] !== '"') i++;
        const value = src.slice(start, i);
        i++; // closing quote
        tokens.push({ type: "STRING", value });
        continue;
      }
      if (isAlpha(ch) || ch === "_" || ch === "$") {
        const start = i;
        while (i < len && (isAlnum(src[i]!) || src[i] === "_" || src[i] === "$" || src[i] === ".")) i++;
        tokens.push({ type: "IDENT", value: src.slice(start, i) });
        continue;
      }
      if (i + 1 < len) {
        const two = src.slice(i, i + 2);
        if (two === "<=" || two === ">=" || two === "<>") {
          tokens.push({ type: "OP", value: two });
          i += 2;
          continue;
        }
      }
      if ("+-*/^%&=<>(),:!".includes(ch)) {
        tokens.push({ type: "OP", value: ch });
        i++;
        continue;
      }
      throw new LinterError(ERR_NAME);
    }
    return tokens;
  }

  // ---- Parser (recursive descent) ----

  private peek(c: Cursor): Token | undefined {
    return c.tokens[c.pos];
  }
  private isOp(c: Cursor, ...vals: string[]): boolean {
    const t = c.tokens[c.pos];
    return !!t && t.type === "OP" && vals.includes(t.value);
  }

  private parseExpr(c: Cursor, sheet: string, stack: string[]): Val {
    let left = this.parseConcat(c, sheet, stack);
    while (this.isOp(c, "=", "<", ">", "<=", ">=", "<>")) {
      const op = c.tokens[c.pos]!.value;
      c.pos++;
      const right = this.parseConcat(c, sheet, stack);
      left = this.compare(left, right, op);
    }
    return left;
  }

  private parseConcat(c: Cursor, sheet: string, stack: string[]): Val {
    let left = this.parseArith(c, sheet, stack);
    while (this.isOp(c, "&")) {
      c.pos++;
      const right = this.parseArith(c, sheet, stack);
      left = this.coerceString(left) + this.coerceString(right);
    }
    return left;
  }

  private parseArith(c: Cursor, sheet: string, stack: string[]): Val {
    let left = this.parseTerm(c, sheet, stack);
    while (this.isOp(c, "+", "-")) {
      const op = c.tokens[c.pos]!.value;
      c.pos++;
      const right = this.parseTerm(c, sheet, stack);
      const a = this.coerceNumber(left);
      const b = this.coerceNumber(right);
      left = op === "+" ? a + b : a - b;
    }
    return left;
  }

  private parseTerm(c: Cursor, sheet: string, stack: string[]): Val {
    let left = this.parseUnary(c, sheet, stack);
    while (this.isOp(c, "*", "/", "%", "^")) {
      const op = c.tokens[c.pos]!.value;
      c.pos++;
      const right = this.parseUnary(c, sheet, stack);
      const a = this.coerceNumber(left);
      const b = this.coerceNumber(right);
      switch (op) {
        case "*":
          left = a * b;
          break;
        case "/":
          if (b === 0) throw new LinterError(ERR_DIV0);
          left = a / b;
          break;
        case "%":
          left = (a / 100.0) * b;
          break;
        case "^":
          left = a ** b;
          break;
      }
    }
    return left;
  }

  private parseUnary(c: Cursor, sheet: string, stack: string[]): Val {
    if (this.isOp(c, "-")) {
      c.pos++;
      const v = this.parseUnary(c, sheet, stack);
      return -this.coerceNumber(v);
    }
    if (this.isOp(c, "+")) {
      c.pos++;
      return this.parseUnary(c, sheet, stack);
    }
    return this.parsePrimary(c, sheet, stack);
  }

  private parsePrimary(c: Cursor, sheet: string, stack: string[]): Val {
    const tok = this.peek(c);
    if (!tok) throw new LinterError(ERR_NAME);

    if (tok.type === "OP" && tok.value === "(") {
      c.pos++;
      const val = this.parseExpr(c, sheet, stack);
      this.expectOp(c, ")");
      return val;
    }

    if (tok.type === "NUMBER") {
      c.pos++;
      return parseFloat(tok.value);
    }
    if (tok.type === "STRING") {
      c.pos++;
      return tok.value;
    }

    if (tok.type === "IDENT") {
      const next = c.tokens[c.pos + 1];
      const next2 = c.tokens[c.pos + 2];
      // Sheet!Ref
      if (next && next.type === "OP" && next.value === "!" && next2 && next2.type === "IDENT") {
        const sheetName = this.cleanSheetName(tok.value);
        c.pos += 2;
        const startTok = c.tokens[c.pos]!.value;
        c.pos++;
        return this.resolveRefOrRange(startTok, sheetName, c, stack);
      }
      // Function call
      if (next && next.type === "OP" && next.value === "(") {
        const name = tok.value.toUpperCase();
        c.pos += 2;
        const args: Val[] = [];
        if (!this.isOp(c, ")")) {
          for (;;) {
            args.push(this.parseExpr(c, sheet, stack));
            if (this.isOp(c, ",")) {
              c.pos++;
              continue;
            }
            break;
          }
        }
        this.expectOp(c, ")");
        return this.callFunction(name, args);
      }
      // Boolean literals
      const upper = tok.value.toUpperCase();
      if (upper === "TRUE" || upper === "FALSE") {
        c.pos++;
        return upper === "TRUE";
      }
      // Bare cell reference
      c.pos++;
      return this.resolveRefOrRange(tok.value, sheet, c, stack);
    }

    throw new LinterError(ERR_NAME);
  }

  private resolveRefOrRange(startRef: string, defaultSheet: string, c: Cursor, stack: string[]): Val {
    const cleanStart = this.cleanRef(startRef);
    if (cleanStart === null) throw new LinterError(ERR_REF);

    const next = c.tokens[c.pos + 1];
    if (this.isOp(c, ":") && next && next.type === "IDENT") {
      const endRef = next.value;
      c.pos += 2;
      const cleanEnd = this.cleanRef(endRef);
      if (cleanEnd === null) throw new LinterError(ERR_REF);
      return this.resolveRange(defaultSheet, cleanStart, cleanEnd, stack);
    }
    return this.resolveCell(defaultSheet, cleanStart, stack);
  }

  private cleanRef(ref: string): string | null {
    const r = ref.replace(/\$/g, "").toUpperCase();
    return /^[A-Z]+\d+$/.test(r) ? r : null;
  }

  private cleanSheetName(name: string): string {
    if (name.length >= 2 && name[0] === "'" && name[name.length - 1] === "'") {
      return name.slice(1, -1);
    }
    return name;
  }

  private resolveCell(sheet: string, a1: string, stack: string[]): Val {
    const key = sheet + "!" + a1;
    if (stack.includes(key)) return ERR_CIRC;
    if (this.cache.has(key)) return this.cache.get(key)!;
    const cell = this.index.get(key);
    if (!cell) return null;
    if (cell.formula !== null) {
      const result = this.evaluate(cell.formula, sheet, [...stack, key]);
      this.cache.set(key, result);
      return result;
    }
    return cell.value;
  }

  private resolveRange(sheet: string, start: string, end: string, stack: string[]): Val[] {
    const a = CellAddress.parse(start);
    const b = CellAddress.parse(end);
    if (a === null || b === null) throw new LinterError(ERR_REF);
    const col1 = Math.min(a[0], b[0]);
    const row1 = Math.min(a[1], b[1]);
    const col2 = Math.max(a[0], b[0]);
    const row2 = Math.max(a[1], b[1]);

    const values: Val[] = [];
    for (let r = row1; r <= row2; r++) {
      for (let col = col1; col <= col2; col++) {
        values.push(this.resolveCell(sheet, CellAddress.letter(col) + r, stack));
      }
    }
    return values;
  }

  private expectOp(c: Cursor, op: string): void {
    const t = c.tokens[c.pos];
    if (!t || t.type !== "OP" || t.value !== op) throw new LinterError(ERR_NAME);
    c.pos++;
  }

  // ---- Coercion / comparison / functions ----

  private coerceNumber(v: Val): number {
    if (this.isError(v)) throw new LinterError(v as string);
    if (Array.isArray(v)) {
      for (const item of v) {
        if (typeof item === "number") return item;
        if (typeof item === "string" && isNumericString(item)) return Number(item);
      }
      throw new LinterError(ERR_VALUE);
    }
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v ? 1 : 0;
    if (v === null) return 0;
    if (typeof v === "string" && isNumericString(v)) return Number(v);
    throw new LinterError(ERR_VALUE);
  }

  private coerceString(v: Val): string {
    if (this.isError(v)) throw new LinterError(v as string);
    if (v === null) return "";
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (Array.isArray(v)) return v.map((x) => this.coerceString(x)).join("");
    return String(v);
  }

  private compare(a: Val, b: Val, op: string): boolean {
    const cmp = this.cmpVal(a, b);
    switch (op) {
      case "=":
        return cmp === 0;
      case "<>":
        return cmp !== 0;
      case "<":
        return cmp < 0;
      case ">":
        return cmp > 0;
      case "<=":
        return cmp <= 0;
      case ">=":
        return cmp >= 0;
      default:
        return false;
    }
  }

  private cmpVal(a: Val, b: Val): number {
    const an = typeof a === "number" || (typeof a === "string" && isNumericString(a));
    const bn = typeof b === "number" || (typeof b === "string" && isNumericString(b));
    if (an && bn) {
      const x = Number(a);
      const y = Number(b);
      return x < y ? -1 : x > y ? 1 : 0;
    }
    const sa = String(a);
    const sb = String(b);
    return sa < sb ? -1 : sa > sb ? 1 : 0;
  }

  private callFunction(name: string, args: Val[]): Val {
    const flat = this.flatten(args);
    switch (name) {
      case "SUM":
        return flat.filter((v) => this.isNumericLike(v)).reduce((s: number, v) => s + this.maybeNum(v), 0);
      case "AVERAGE":
      case "AVG":
        return this.avg(flat);
      case "COUNT":
        return flat.filter((v) => this.isNumericLike(v)).length;
      case "COUNTA":
        return flat.filter((v) => v !== null && v !== "").length;
      case "MIN":
        return this.minMax(flat, true);
      case "MAX":
        return this.minMax(flat, false);
      case "IF":
        return this.ifFn(args);
      case "ROUND":
        return phpRound(this.coerceNumber(args[0] ?? 0), Math.trunc(this.coerceNumber(args[1] ?? 0)));
      case "ABS":
        return Math.abs(this.coerceNumber(args[0] ?? 0));
      case "LEN":
        return this.coerceString(args[0] ?? "").length;
      case "UPPER":
        return this.coerceString(args[0] ?? "").toUpperCase();
      case "LOWER":
        return this.coerceString(args[0] ?? "").toLowerCase();
      case "CONCAT":
      case "CONCATENATE":
        return flat.map((v) => this.coerceString(v)).join("");
      case "TRUE":
        return true;
      case "FALSE":
        return false;
      default:
        return ERR_NAME;
    }
  }

  private flatten(args: Val[]): Val[] {
    const out: Val[] = [];
    for (const a of args) {
      if (Array.isArray(a)) for (const x of a) out.push(x);
      else out.push(a);
    }
    return out;
  }

  private isNumericLike(v: Val): boolean {
    return (
      typeof v === "number" ||
      (typeof v === "string" && isNumericString(v)) ||
      typeof v === "boolean"
    );
  }

  private maybeNum(v: Val): number {
    return typeof v === "boolean" ? (v ? 1 : 0) : Number(v);
  }

  private avg(flat: Val[]): number | string {
    const nums = flat.filter((v) => this.isNumericLike(v));
    if (nums.length === 0) return ERR_DIV0;
    return nums.reduce((s: number, v) => s + this.maybeNum(v), 0) / nums.length;
  }

  private minMax(flat: Val[], min: boolean): number {
    const nums = flat.filter((v) => this.isNumericLike(v));
    if (nums.length === 0) return 0;
    const vals = nums.map((v) => this.maybeNum(v));
    return min ? Math.min(...vals) : Math.max(...vals);
  }

  private ifFn(args: Val[]): Val {
    if (args.length < 2) return ERR_VALUE;
    const cond = args[0]!;
    const truthy =
      typeof cond === "boolean"
        ? cond
        : cond !== null && cond !== 0 && cond !== "" && cond !== "FALSE";
    return truthy ? args[1]! : (args[2] ?? false);
  }

  // ---- Hints ----

  private hint(error: string, formula: string, sheet: string): string {
    switch (error) {
      case ERR_VALUE:
        return this.hintValue(formula, sheet);
      case ERR_REF:
        return "A cell reference points to a cell that doesn't exist in the workbook. Check column letters and row numbers.";
      case ERR_NAME:
        return "The formula references an unknown function or has a syntax error. Holy Sheet supports: SUM, AVERAGE, COUNT, COUNTA, MIN, MAX, IF, ROUND, ABS, LEN, UPPER, LOWER, CONCAT.";
      case ERR_DIV0:
        return "Division by zero — the divisor evaluated to 0.";
      case ERR_CIRC:
        return "Circular reference — the formula directly or transitively depends on its own cell.";
      default:
        return "Formula evaluation failed.";
    }
  }

  private hintValue(formula: string, sheet: string): string {
    const re = /(?:([A-Za-z][A-Za-z0-9_]*)!)?\$?([A-Z]+)\$?(\d+)/gi;
    const offenders: string[] = [];
    for (const m of formula.matchAll(re)) {
      const sheetName = m[1] ? m[1] : sheet;
      const a1 = m[2]!.toUpperCase() + m[3]!;
      const key = sheetName + "!" + a1;
      const cell = this.index.get(key);
      if (!cell) continue;
      const value = this.cache.has(key)
        ? this.cache.get(key)!
        : cell.formula !== null
          ? null
          : cell.value;
      if (typeof value === "string" && !isNumericString(value) && value !== "") {
        const row = parseInt(m[3]!, 10);
        const col = m[2]!;
        const next = sheetName + "!" + col.toUpperCase() + (row + 1);
        const nextCell = this.index.get(next);
        let suggestion = "";
        if (nextCell && isNumericString(nextCell.value as never)) {
          suggestion = ` Did you mean ${col.toUpperCase()}${row + 1}? (it holds ${String(nextCell.value)})`;
        }
        offenders.push(`${a1} = "${value}" (string)${suggestion}`);
      }
    }
    if (offenders.length > 0) {
      return "Arithmetic on a non-numeric cell: " + offenders.join("; ");
    }
    return "A non-numeric value was used in arithmetic. Check that all referenced cells contain numbers.";
  }
}

/** PHP round() — half away from zero. */
function phpRound(x: number, n: number): number {
  const factor = 10 ** n;
  return (Math.sign(x) * Math.round(Math.abs(x) * factor)) / factor;
}
