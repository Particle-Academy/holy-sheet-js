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

/** The functions callFunction() evaluates. Anything else is #NAME?. */
const FUNCTIONS = new Set([
  "SUM", "AVERAGE", "AVG", "COUNT", "COUNTA", "MIN", "MAX", "IF", "ROUND",
  "ABS", "LEN", "UPPER", "LOWER", "CONCAT", "CONCATENATE", "TRUE", "FALSE",
]);

/**
 * A sheet qualifier, quoted ('My Sheet'!) or bare (Sheet2!). Group 1 is the name
 * as written, quotes included.
 */
const SHEET_QUALIFIER = "('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!";

class LinterError extends Error {
  /**
   * @param cause Why a #NAME? arose in THIS formula: "syntax" or "function".
   *   Null for an error that was computed, or inherited from another cell.
   * @param detail The syntax fault, or the unknown function's name.
   */
  constructor(
    public errorCode: string,
    public cause: string | null = null,
    public detail = "",
  ) {
    super(errorCode);
  }

  static syntax(detail: string): LinterError {
    return new LinterError(ERR_NAME, "syntax", detail);
  }
}

interface Token {
  type: "NUMBER" | "STRING" | "IDENT" | "SHEET" | "OP";
  value: string;
}
interface Cursor {
  tokens: Token[];
  pos: number;
}

type Val = number | string | boolean | null | Val[];

// PHP's ctype_* classes, which are ASCII-only. A Unicode-aware test here would
// accept a no-break space or a letter like "é" that the PHP linter rejects.
const isSpace = (ch: string) => ch === " " || (ch >= "\t" && ch <= "\r");
const isDigit = (ch: string) => ch >= "0" && ch <= "9";
const isAlpha = (ch: string) => (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z");
const isAlnum = (ch: string) => isAlpha(ch) || isDigit(ch);

/**
 * Evaluates every formula and reports Excel-style errors. Mirrors PHP
 * `Schema\FormulaLinter` rule-for-rule.
 */
export class FormulaLinter {
  private index = new Map<string, Cell>();
  private cache = new Map<string, Val>();
  /**
   * Every sheet name, keyed by its lower-cased form. Excel matches sheet names
   * without regard to case, and a reference to a sheet that is not here is
   * `#REF!` rather than a range of blanks.
   */
  private sheetNames = new Map<string, string>();
  /**
   * Why a cell's own formula produced #NAME?, keyed `Sheet!A1`: a syntax
   * error, an unknown function, or a linter fault. A cell with #NAME? and no
   * entry here inherited it from a cell it references.
   */
  private nameCauses = new Map<string, { cause: string; detail: string }>();
  /**
   * Greater than 0 while parsing an IF branch that is not taken. The parser
   * evaluates as it goes, so a skipped branch is still PARSED (its syntax
   * errors and unknown functions are reported) but nothing in it is
   * EVALUATED: no operator runs, no cell is read, and every value is null.
   */
  private skipping = 0;

  lint(schema: unknown): FormulaProblem[] {
    const workbook = new Normalizer().normalize(schema);
    this.sheetNames = new Map(workbook.sheets.map((sheet) => [sheet.name.toLowerCase(), sheet.name]));
    this.nameCauses = new Map();
    this.skipping = 0;
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
            hint: this.hint(result as string, key, cell.formula, sheet.name),
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
    const key = stack[stack.length - 1]!;
    this.nameCauses.delete(key);
    try {
      const c: Cursor = { tokens: this.tokenize(formula), pos: 0 };
      const result = this.parseExpr(c, defaultSheet, stack);
      // If anything follows the parsed expression, the formula was malformed.
      if (c.pos < c.tokens.length) {
        throw LinterError.syntax(`unexpected ${this.describeToken(c.tokens[c.pos]!)} after a complete expression`);
      }
      return result;
    } catch (e) {
      if (e instanceof LinterError) {
        if (e.cause !== null) this.nameCauses.set(key, { cause: e.cause, detail: e.detail });
        return e.errorCode;
      }
      // A bug in this linter, not a fault in the formula. Still reported, so it
      // is not silently passed, but the hint says whose fault it is.
      this.nameCauses.set(key, { cause: "internal", detail: e instanceof Error ? e.constructor.name : typeof e });
      return ERR_NAME;
    }
  }

  private describeToken(tok: Token): string {
    switch (tok.type) {
      case "STRING":
        return 'text "' + tok.value + '"';
      case "NUMBER":
        return "number " + tok.value;
      default:
        return "'" + tok.value + "'";
    }
  }

  // ---- Tokenizer ----

  private tokenize(src: string): Token[] {
    const tokens: Token[] = [];
    let i = 0;
    const len = src.length;

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
      // String literal, with Excel's "" for a literal quote: "a""b" is a"b.
      // A literal that never closes is a syntax error, not text running to the
      // end of the formula.
      if (ch === '"') {
        i++;
        let value = "";
        for (;;) {
          if (i >= len) throw LinterError.syntax('a text literal opens with " and never closes');
          if (src[i] === '"') {
            if (i + 1 < len && src[i + 1] === '"') {
              value += '"';
              i += 2;
              continue;
            }
            i++;
            break;
          }
          value += src[i];
          i++;
        }
        tokens.push({ type: "STRING", value });
        continue;
      }
      // Quoted sheet name: 'My Sheet', with Excel's '' for a literal quote.
      // Kept with its quotes; cleanSheetName() unwraps it. A quote that never
      // closes is a syntax error, not a name running to the end.
      if (ch === "'") {
        const start = i;
        i++;
        for (;;) {
          if (i >= len) throw LinterError.syntax("a quoted sheet name opens with ' and never closes");
          if (src[i] === "'") {
            if (i + 1 < len && src[i + 1] === "'") {
              i += 2;
              continue;
            }
            i++;
            break;
          }
          i++;
        }
        tokens.push({ type: "SHEET", value: src.slice(start, i) });
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
      // Unknown character: fail soft as #NAME?. Named only when it is ASCII,
      // as PHP does (a lone byte of a multi-byte character there is not valid
      // UTF-8, and the hint is byte-identical across the three linters).
      throw LinterError.syntax(
        ch.charCodeAt(0) < 128 ? `unexpected character '${ch}'` : "unexpected non-ASCII character outside a text literal",
      );
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
      left = this.skipping > 0 ? null : this.compare(left, right, op);
    }
    return left;
  }

  private parseConcat(c: Cursor, sheet: string, stack: string[]): Val {
    let left = this.parseArith(c, sheet, stack);
    while (this.isOp(c, "&")) {
      c.pos++;
      const right = this.parseArith(c, sheet, stack);
      left = this.skipping > 0 ? null : this.coerceString(left) + this.coerceString(right);
    }
    return left;
  }

  private parseArith(c: Cursor, sheet: string, stack: string[]): Val {
    let left = this.parseTerm(c, sheet, stack);
    while (this.isOp(c, "+", "-")) {
      const op = c.tokens[c.pos]!.value;
      c.pos++;
      const right = this.parseTerm(c, sheet, stack);
      if (this.skipping > 0) {
        left = null;
        continue;
      }
      const a = this.coerceNumber(left);
      const b = this.coerceNumber(right);
      left = op === "+" ? a + b : a - b;
    }
    return left;
  }

  private parseTerm(c: Cursor, sheet: string, stack: string[]): Val {
    let left = this.parsePower(c, sheet, stack);
    while (this.isOp(c, "*", "/")) {
      const op = c.tokens[c.pos]!.value;
      c.pos++;
      const right = this.parsePower(c, sheet, stack);
      if (this.skipping > 0) {
        left = null;
        continue;
      }
      const a = this.coerceNumber(left);
      const b = this.coerceNumber(right);
      if (op === "*") {
        left = a * b;
      } else {
        if (b === 0) throw new LinterError(ERR_DIV0);
        left = a / b;
      }
    }
    return left;
  }

  /**
   * `^` binds tighter than `*` and `/` (=2*3^2 is 18) and, as in Excel, is
   * left-associative (=2^3^2 is 64). Negation binds tighter than `^`, so
   * =-2^2 is 4: Excel's order, not mathematics'.
   */
  private parsePower(c: Cursor, sheet: string, stack: string[]): Val {
    let left = this.parseUnary(c, sheet, stack);
    while (this.isOp(c, "^")) {
      c.pos++;
      const right = this.parseUnary(c, sheet, stack);
      left = this.skipping > 0 ? null : this.coerceNumber(left) ** this.coerceNumber(right);
    }
    return left;
  }

  private parseUnary(c: Cursor, sheet: string, stack: string[]): Val {
    if (this.isOp(c, "-")) {
      c.pos++;
      const v = this.parseUnary(c, sheet, stack);
      return this.skipping > 0 ? null : -this.coerceNumber(v);
    }
    if (this.isOp(c, "+")) {
      c.pos++;
      return this.parseUnary(c, sheet, stack);
    }
    return this.parsePercent(c, sheet, stack);
  }

  /**
   * Excel's `%` is a POSTFIX operator that divides by 100: `50%` is 0.5 and
   * `A1%` is A1/100. It binds tighter than `^`, `*` and `/`. There is no
   * binary `%`, so `10%2` leaves a stray `2` and is a syntax error.
   */
  private parsePercent(c: Cursor, sheet: string, stack: string[]): Val {
    let value = this.parsePrimary(c, sheet, stack);
    while (this.isOp(c, "%")) {
      c.pos++;
      value = this.skipping > 0 ? null : this.coerceNumber(value) / 100.0;
    }
    return value;
  }

  private parsePrimary(c: Cursor, sheet: string, stack: string[]): Val {
    const tok = this.peek(c);
    if (!tok) throw LinterError.syntax("the formula ends where a value was expected");

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

    // A quoted name is only ever a sheet qualifier. On its own it is not a
    // value, so it falls through to #NAME? below.
    if (tok.type === "SHEET" || tok.type === "IDENT") {
      const next = c.tokens[c.pos + 1];
      const next2 = c.tokens[c.pos + 2];
      // Sheet!Ref: (SHEET | IDENT) '!' IDENT
      if (next && next.type === "OP" && next.value === "!" && next2 && next2.type === "IDENT") {
        const sheetName = this.sheetNames.get(this.cleanSheetName(tok.value).toLowerCase());
        if (sheetName === undefined && this.skipping === 0) throw new LinterError(ERR_REF);
        c.pos += 2;
        const startTok = c.tokens[c.pos]!.value;
        c.pos++;
        return this.resolveRefOrRange(startTok, sheetName ?? tok.value, c, stack);
      }
      if (tok.type === "SHEET") {
        throw LinterError.syntax(`the quoted name ${tok.value} is not followed by ! and a cell reference`);
      }
      // Function call
      if (next && next.type === "OP" && next.value === "(") {
        const name = tok.value.toUpperCase();
        // Excel resolves names before it evaluates anything, so an unknown
        // function is #NAME? whatever its arguments hold and whichever IF
        // branch it sits in.
        if (!FUNCTIONS.has(name)) throw new LinterError(ERR_NAME, "function", name);
        if (name === "IF") return this.parseIf(c, sheet, stack);
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
        return this.skipping > 0 ? null : this.callFunction(name, args);
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

    throw LinterError.syntax(`unexpected ${this.describeToken(tok)} where a value was expected`);
  }

  /**
   * IF, evaluated lazily as Excel does: `=IF(A2=0,0,1/A2)` with A2 = 0 is 0,
   * not #DIV/0!. Every argument is parsed, so a syntax error or an unknown
   * function anywhere in it is still reported, but only the condition and the
   * branch it selects are evaluated.
   *
   * The order of the verdicts is deliberate. A syntax error anywhere wins,
   * because Excel refuses a formula it cannot parse and computes nothing.
   * Then fewer than two arguments is #VALUE!. Then an error in the condition
   * propagates, and neither branch is evaluated.
   */
  private parseIf(c: Cursor, sheet: string, stack: string[]): Val {
    c.pos += 2; // consume IF '('
    const args: Val[] = [];
    let take: number | null = null; // the argument the condition selects: 1 or 2
    let conditionError: LinterError | null = null; // an error the condition evaluated to
    if (!this.isOp(c, ")")) {
      for (;;) {
        const i = args.length;
        const skip = i > 0 && i !== take;
        if (skip) this.skipping++;
        try {
          args.push(this.parseExpr(c, sheet, stack));
        } finally {
          if (skip) this.skipping--;
        }
        if (i === 0 && this.skipping === 0) {
          try {
            take = this.toLogical(args[0]!) ? 1 : 2;
          } catch (e) {
            if (!(e instanceof LinterError)) throw e;
            conditionError = e;
          }
        }
        if (this.isOp(c, ",")) {
          c.pos++;
          continue;
        }
        break;
      }
    }
    this.expectOp(c, ")");

    if (this.skipping > 0) return null;
    if (args.length < 2) return ERR_VALUE;
    if (conditionError !== null) throw conditionError;
    // `??` as PHP's: a blank else branch reads as FALSE, as a missing one does.
    return take === 1 ? args[1]! : (args[2] ?? false);
  }

  private resolveRefOrRange(startRef: string, defaultSheet: string, c: Cursor, stack: string[]): Val {
    const next = c.tokens[c.pos + 1];
    const isRange = this.isOp(c, ":") && !!next && next.type === "IDENT";
    if (this.skipping > 0) {
      // In an untaken IF branch: consume a range's `:` and end so the parse
      // carries on, but read nothing. A #REF!, or an error held by the cell,
      // belongs to a branch Excel never evaluates.
      if (isRange) c.pos += 2;
      return null;
    }
    const cleanStart = this.cleanRef(startRef);
    if (cleanStart === null) throw new LinterError(ERR_REF);

    if (isRange) {
      const endRef = next!.value;
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

  /**
   * Unwrap the single quotes Excel puts around a sheet name that needs them,
   * and turn its doubled quote back into one: 'Q3 ''Final''' is Q3 'Final'.
   */
  private cleanSheetName(name: string): string {
    if (name.length >= 2 && name[0] === "'" && name[name.length - 1] === "'") {
      // split/join rather than replaceAll: the package targets a lib before ES2021.
      return name.slice(1, -1).split("''").join("'");
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
    if (!t) throw LinterError.syntax(`expected '${op}', but the formula ends`);
    if (t.type !== "OP" || t.value !== op) {
      throw LinterError.syntax(`expected '${op}', found ${this.describeToken(t)}`);
    }
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

  /**
   * Excel's comparison. An error on either side propagates. Otherwise the
   * types rank number < text < logical, so ="5">9 is TRUE: numeric TEXT is
   * text, never a number. Text compares without regard to case. A blank
   * takes the other side's type (0, "" or FALSE); two blanks are equal.
   */
  private cmpVal(a: Val, b: Val): number {
    if (this.isError(a)) throw new LinterError(a as string);
    if (this.isError(b)) throw new LinterError(b as string);
    if (Array.isArray(a) || Array.isArray(b)) {
      // A range. Implicit intersection is out of scope; unchanged from before.
      return compareCodePoints(String(a), String(b));
    }
    const x = a ?? blankLike(b);
    const y = b ?? blankLike(x);
    const rank = (v: Val) => (typeof v === "boolean" ? 2 : typeof v === "string" ? 1 : 0);
    if (rank(x) !== rank(y)) return rank(x) < rank(y) ? -1 : 1;
    if (typeof x === "string") {
      return compareCodePoints(x.toLowerCase(), String(y).toLowerCase());
    }
    const p = Number(x);
    const q = Number(y);
    return p < q ? -1 : p > q ? 1 : 0;
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
      case "ROUND":
        return excelRound(this.coerceNumber(args[0] ?? 0), Math.trunc(this.coerceNumber(args[1] ?? 0)));
      case "ABS":
        return Math.abs(this.coerceNumber(args[0] ?? 0));
      case "LEN":
        // String length IS Excel's LEN: both count UTF-16 code units, so "é"
        // is 1 and an emoji outside the Basic Multilingual Plane is 2.
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
        // Unreachable: parsePrimary checks FUNCTIONS before parsing args, and
        // IF never arrives here, because parseIf evaluates it lazily.
        throw new LinterError(ERR_NAME, "function", name);
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

  /**
   * Excel's reading of an IF condition: an error propagates, a blank is
   * FALSE, a number is TRUE unless it is 0, and text is TRUE or FALSE only
   * when it spells one of them (in any case). Any other text, "" included,
   * is #VALUE!.
   */
  private toLogical(v: Val): boolean {
    if (this.isError(v)) throw new LinterError(v as string);
    if (typeof v === "boolean") return v;
    if (v === null) return false;
    if (typeof v === "number") return v !== 0;
    if (typeof v === "string") {
      const upper = v.toUpperCase();
      if (upper === "TRUE") return true;
      if (upper === "FALSE") return false;
      throw new LinterError(ERR_VALUE);
    }
    // A range. Implicit intersection is out of scope; unchanged from before.
    return true;
  }

  // ---- Hints ----

  private hint(error: string, key: string, formula: string, sheet: string): string {
    switch (error) {
      case ERR_VALUE:
        return this.hintValue(formula, sheet);
      case ERR_REF:
        return this.hintRef(formula);
      case ERR_NAME:
        return this.hintName(key, formula, sheet);
      case ERR_DIV0:
        return "Division by zero — the divisor evaluated to 0.";
      case ERR_CIRC:
        return "Circular reference — the formula directly or transitively depends on its own cell.";
      default:
        return "Formula evaluation failed.";
    }
  }

  /**
   * #NAME? has three causes in this formula, plus inheritance, and each hint
   * says which one it is. A syntax error used to be told it "references an
   * unknown function", and an agent went hunting for a function that was fine.
   * The texts are byte-identical to PHP's and Python's: see PHP
   * `FormulaLinter::hintName` for why a syntax error keeps the #NAME? label.
   */
  private hintName(key: string, formula: string, sheet: string): string {
    const cause = this.nameCauses.get(key);
    switch (cause?.cause) {
      case "syntax":
        return (
          `Syntax error: ${cause.detail}. Excel cannot parse this formula. Check for an unbalanced parenthesis or quote, ` +
          'a missing operator or argument, or a character a formula cannot contain. Text goes in double quotes, and a quote inside text is written twice: "a""b".'
        );
      case "function":
        return `Unknown function ${cause.detail}. Holy Sheet supports: SUM, AVERAGE, COUNT, COUNTA, MIN, MAX, IF, ROUND, ABS, LEN, UPPER, LOWER, CONCAT.`;
      case "internal":
        return `The linter itself failed on this formula (${cause.detail}). That is a holy-sheet defect, not evidence that the formula is wrong.`;
      default:
        return this.hintInheritedName(formula, sheet);
    }
  }

  private hintInheritedName(formula: string, sheet: string): string {
    const re = new RegExp(`(?:${SHEET_QUALIFIER})?\\$?([A-Z]+)\\$?(\\d+)`, "gi");
    const sources: string[] = [];
    for (const m of formula.matchAll(re)) {
      const sheetName = m[1]
        ? (this.sheetNames.get(this.cleanSheetName(m[1]).toLowerCase()) ?? this.cleanSheetName(m[1]))
        : sheet;
      const a1 = m[2]!.toUpperCase() + m[3]!;
      if (this.cache.get(sheetName + "!" + a1) === ERR_NAME) {
        const source = m[1] ? m[1] + "!" + a1 : a1;
        if (!sources.includes(source)) sources.push(source);
      }
    }
    return sources.length === 0
      ? "This formula depends on a cell whose value is #NAME?. Fix that cell first."
      : `This formula depends on ${sources.join(", ")}, whose value is #NAME?. Fix that cell first.`;
  }

  private hintRef(formula: string): string {
    // Name the sheet when that is what is missing. "A cell doesn't exist" sends
    // an agent to check column letters in a sheet it never created.
    for (const m of formula.matchAll(new RegExp(SHEET_QUALIFIER, "g"))) {
      const name = this.cleanSheetName(m[1]!);
      if (!this.sheetNames.has(name.toLowerCase())) {
        return `The formula refers to a sheet named '${name}', and this workbook has no such sheet. Its sheets are: ${[...this.sheetNames.values()].join(", ")}. Quote a name that contains spaces or punctuation: 'My Sheet'!A1.`;
      }
    }
    return "A cell reference points to a cell that doesn't exist in the workbook. Check column letters and row numbers.";
  }

  private hintValue(formula: string, sheet: string): string {
    const re = new RegExp(`(?:${SHEET_QUALIFIER})?\\$?([A-Z]+)\\$?(\\d+)`, "gi");
    const offenders: string[] = [];
    for (const m of formula.matchAll(re)) {
      const sheetName = m[1]
        ? (this.sheetNames.get(this.cleanSheetName(m[1]).toLowerCase()) ?? this.cleanSheetName(m[1]))
        : sheet;
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

/** What a blank cell reads as when compared with `other`. */
function blankLike(other: Val): Val {
  if (typeof other === "boolean") return false;
  if (typeof other === "string") return "";
  return 0;
}

/**
 * Order two strings by code point, which is PHP's strcmp on UTF-8 and Python's
 * str comparison. JavaScript's `<` compares UTF-16 code units, which puts a
 * character outside the Basic Multilingual Plane BELOW U+E000..U+FFFF.
 */
function compareCodePoints(a: string, b: string): number {
  const x = Array.from(a);
  const y = Array.from(b);
  const n = Math.min(x.length, y.length);
  for (let i = 0; i < n; i++) {
    const p = x[i]!.codePointAt(0)!;
    const q = y[i]!.codePointAt(0)!;
    if (p !== q) return p < q ? -1 : 1;
  }
  return x.length < y.length ? -1 : x.length > y.length ? 1 : 0;
}

/**
 * Excel's ROUND (and PHP's round()): half away from zero, on the DECIMAL value
 * the number was written as. `Math.round(x * 10 ** n)` rounds the binary
 * double instead, so ROUND(1.005, 2) came out 1: 1.005 is 1.00499999... in
 * binary. Shifting the shortest round-trip decimal digits by `n` places keeps
 * the value exactly as written.
 */
function excelRound(x: number, n: number): number {
  if (!Number.isFinite(x)) return x;
  const [mantissa, exponent] = Math.abs(x).toExponential().split("e") as [string, string];
  const shifted = Number(`${mantissa}e${Number(exponent) + n}`);
  // More places than a double holds: nothing to round, as PHP's round() says.
  if (!Number.isFinite(shifted)) return x;
  const rounded = Math.round(shifted);
  const digits = String(rounded);
  // At 1e21 and above String() is exponential, which cannot take a second `e`.
  return Math.sign(x) * (digits.includes("e") ? rounded / 10 ** n : Number(`${digits}e${-n}`));
}
