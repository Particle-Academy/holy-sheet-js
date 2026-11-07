import { formatFloat } from "../util";
import { phpInt, phpIsNumeric } from "../reader/ods/ns";

/**
 * One portable table -> RFC 4180 CSV. A port of PHP `Exchange\Csv::toCsv()`
 * (with the parts of `Display`, `Portable::layout` and `Report` it uses), held
 * to PHP's bytes and report by tests/csv-parity.test.ts.
 *
 * Formula-injection escaping is ON by default: a `string` or `error` field
 * beginning with `=`, `+`, `-`, `@`, TAB or CR is prefixed with `'` before it
 * is quoted, and reported `csv-escaped`. Numbers are never escaped. Only the
 * host turns it off (`escapeFormulas: false`).
 *
 * The input is portable content's table shape: `{rows: [{cells: [{t, v,
 * colSpan?, rowSpan?, text?, format?}]}]}`, where `t` is `string`, `number`,
 * `bool`, `date`, `datetime`, `error` or `empty`. Lines end CRLF, every line
 * including the last.
 */

/** Bad options (`options_invalid`). Mirrors PHP `Exchange\OptionsInvalid`. */
export class OptionsInvalid extends Error {
  static readonly CODE = "options_invalid";
  constructor(message = "") {
    super("[holy-sheet] options_invalid" + (message === "" ? "" : ": " + message));
    this.name = "OptionsInvalid";
    Object.setPrototypeOf(this, OptionsInvalid.prototype);
  }
}

export interface CsvOptions {
  /** Prefix a leading `= + - @ TAB CR` in a string or error field with `'`. Default true. */
  escapeFormulas?: boolean;
  /** Write what a spreadsheet shows (`text`, or the number `format`) instead of the canonical value. Default false. */
  display?: boolean;
}

export interface ReportEntry {
  kind: string;
  count: number;
  where?: string[];
}

export interface ConversionReport {
  dropped: ReportEntry[];
  approximated: ReportEntry[];
  guard: { findings: Array<{ rule: string; count: number }> };
}

type Any = any; // eslint-disable-line @typescript-eslint/no-explicit-any

const LEADS = new Set(["=", "+", "-", "@", "\t", "\r"]);
const WHERE_CAP = 16;
const SYMBOLS: Record<string, string> = {
  USD: "$", EUR: "€", GBP: "£", JPY: "¥", CNY: "¥", INR: "₹",
  AUD: "A$", CAD: "C$", CHF: "CHF", KRW: "₩",
};
const DEFAULT_DECIMALS: Record<string, number> = { number: 0, currency: 2, percent: 1 };

export const Csv = {
  toCsv(table: Record<string, unknown>, options: CsvOptions | null = null): { text: string; report: ConversionReport } {
    const opts = (options ?? {}) as Record<string, unknown>;
    for (const [k, v] of Object.entries(opts)) {
      if ((k !== "escapeFormulas" && k !== "display") || typeof v !== "boolean") throw new OptionsInvalid(`bad csv option ${k}`);
    }
    const escape = opts["escapeFormulas"] !== false;
    const showDisplay = opts["display"] === true;

    // Only the two approximated kinds this conversion can report; sorted by kind as PHP's Report sorts them.
    const approximated = new Map<string, { count: number; where: string[] }>();
    const add = (kind: string, where: string): void => {
      const e = approximated.get(kind) ?? { count: 0, where: [] };
      e.count++;
      if (e.where.length < WHERE_CAP) e.where.push(where);
      approximated.set(kind, e);
    };

    const rows = Array.isArray(table["rows"]) ? (table["rows"] as Any[]) : Object.values((table["rows"] as Any) ?? {});
    const g = layout(rows);
    // A row's fields keep PHP's array order: a column first set past the
    // table's width is appended, which is what a ragged table gives in PHP.
    const grid: Map<number, string>[] = rows.map(() => {
      const line = new Map<number, string>();
      for (let c = 0; c < g.width; c++) line.set(c, "");
      return line;
    });

    rows.forEach((r: Any, ri: number) => {
      const cells = r?.cells;
      if (cells === null || typeof cells !== "object") {
        throw new TypeError(`[holy-sheet] csv: rows[${ri}] has no cells list`);
      }
      (Array.isArray(cells) ? cells : Object.values(cells)).forEach((cell: Any, ci: number) => {
        const cp = `/rows/${ri}/cells/${ci}`;
        const anchor = g.anchors[ri]?.[ci];
        if (anchor === undefined) {
          // PHP reads past its layout here (a warning, then a garbled row). Refuse instead.
          throw new TypeError(`[holy-sheet] csv: the table's rows do not form a grid at row ${ri}`);
        }
        const [ar, ac] = anchor;
        if ((cell["colSpan"] ?? 1) > 1 || (cell["rowSpan"] ?? 1) > 1) add("merge-flattened", cp);
        const t = phpString(cell["t"]);
        let f: string;
        if (showDisplay) {
          f = typeof cell["text"] === "string"
            ? cell["text"]
            : cell["format"] !== undefined && cell["format"] !== null
              ? display(t, cell["v"] ?? null, cell["format"])
              : canonical(t, cell["v"] ?? null);
        } else {
          f = canonical(t, cell["v"] ?? null);
        }
        if (escape && (t === "string" || t === "error") && f !== "" && LEADS.has(f[0]!)) {
          f = "'" + f;
          add("csv-escaped", cp);
        }
        if (/[",\r\n]/.test(f)) f = '"' + f.replace(/"/g, '""') + '"';
        let line = grid[ar];
        if (line === undefined) {
          line = new Map();
          grid[ar] = line;
        }
        line.set(ac, f);
      });
    });

    let text = "";
    for (const line of grid) text += [...(line ?? new Map()).values()].join(",") + "\r\n";

    const kinds = [...approximated.keys()].sort(compareCodeUnits);
    return {
      text,
      report: {
        dropped: [],
        approximated: kinds.map((kind) => ({ kind, count: approximated.get(kind)!.count, where: approximated.get(kind)!.where })),
        guard: { findings: [] },
      },
    };
  },
};

/** PHP `Portable::layout()`: where each cell's top-left lands, and the table's width. */
function layout(rows: Any[]): { ok: boolean; width: number; anchors: Array<Array<[number, number]>> } {
  const covered = new Map<number, Set<number>>();
  const anchors: Array<Array<[number, number]>> = [];
  let width = -1;
  const n = rows.length;
  for (let r = 0; r < n; r++) {
    const row = rows[r];
    const occ = new Set<number>(covered.get(r) ?? []);
    const coveredCount = occ.size;
    const placed: Array<[number, number]> = [];
    let c = 0;
    let span = 0;
    const cells = row !== null && typeof row === "object" && row.cells !== null && typeof row.cells === "object"
      ? (Array.isArray(row.cells) ? row.cells : Object.values(row.cells))
      : [];
    for (const cell of cells) {
      while (occ.has(c)) c++;
      const cs = phpIntCast(cell?.["colSpan"] ?? 1);
      const rs = phpIntCast(cell?.["rowSpan"] ?? 1);
      if (r + rs > n) return { ok: false, width, anchors };
      placed.push([r, c]);
      for (let dr = 0; dr < rs; dr++) {
        for (let dc = 0; dc < cs; dc++) {
          if (dr === 0) occ.add(c + dc);
          else {
            let set = covered.get(r + dr);
            if (!set) covered.set(r + dr, (set = new Set()));
            set.add(c + dc);
          }
        }
      }
      span += cs;
      c += cs;
    }
    anchors.push(placed);
    let max = 0;
    for (const k of occ) if (k + 1 > max) max = k + 1;
    const w = coveredCount + span;
    if (w !== max) return { ok: false, width, anchors };
    if (width === -1) width = w;
    else if (w !== width) return { ok: false, width, anchors };
  }
  return { ok: true, width: Math.max(width, 0), anchors };
}

/** PHP `Display::canonical()`. */
function canonical(t: string, v: unknown): string {
  if (t === "number") return formatFloat(typeof v === "number" ? v : 0);
  if (t === "bool") return phpTruthy(v) ? "TRUE" : "FALSE";
  if (t === "empty") return "";
  return typeof v === "string" ? v : phpJsonEncode(v);
}

/** PHP `Display::display()`: the English-only number display a spreadsheet shows. */
function display(t: string, v: unknown, format: unknown): string {
  const fmt = format !== null && typeof format === "object" && !Array.isArray(format) ? (format as Any) : null;
  const kind = fmt?.kind ?? null;

  if (t === "number" && (kind === "number" || kind === "currency" || kind === "percent") && typeof v === "number") {
    const places = typeof fmt.decimals === "number" && Number.isFinite(fmt.decimals) && Math.floor(fmt.decimals) === fmt.decimals
      ? fmt.decimals
      : DEFAULT_DECIMALS[kind]!;
    let negative = v < 0 || Object.is(v, -0);
    let d = plainDecimal(v);
    if (kind === "percent") d = shiftTwo(d);
    const rounded = roundHalfAway(d, places);
    if (/^[0.]+$/.test(rounded)) negative = false;
    const parts = rounded.split(".");
    let whole = parts[0]!;
    if (kind !== "percent") whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    const num = whole + (parts.length > 1 ? "." + parts[1] : "");
    const sign = negative ? "-" : "";
    if (kind === "percent") return sign + num + "%";
    if (kind === "currency") {
      const code = typeof fmt.currency === "string" ? fmt.currency : "USD";
      return sign + (SYMBOLS[code] ?? code + " ") + num;
    }
    return sign + num;
  }

  if ((t === "date" || t === "datetime") && (kind === "date" || kind === "datetime") && typeof v === "string") {
    const date = v.slice(0, 10);
    if (kind === "date") return date;
    return date + " " + (t === "datetime" ? v.slice(11, 19) : "00:00:00");
  }

  return canonical(t, v);
}

/** The shortest decimal string that round-trips to |v|, without an exponent. PHP `Display::plainDecimal()`. */
function plainDecimal(v: number): string {
  const s = String(Math.abs(v));
  const m = /^(\d+)(?:\.(\d+))?e([+-]?\d+)$/i.exec(s);
  if (!m) return s;
  const digits = m[1]! + (m[2] ?? "");
  const point = m[1]!.length + Number(m[3]);
  if (point <= 0) return "0." + "0".repeat(-point) + digits;
  if (point >= digits.length) return digits + "0".repeat(point - digits.length);
  return digits.slice(0, point) + "." + digits.slice(point);
}

/** x100 by moving the decimal point two places to the right. */
function shiftTwo(d: string): string {
  const [w, frac = ""] = d.split(".", 2) as [string, string?];
  const f = frac.padEnd(2, "0");
  let whole = (w + f.slice(0, 2)).replace(/^0+/, "");
  if (whole === "") whole = "0";
  return whole + (f.length > 2 ? "." + f.slice(2) : "");
}

/** Round a plain non-negative decimal string to `places`, half away from zero, on its digits. */
function roundHalfAway(d: string, places: number): string {
  const idx = d.indexOf(".");
  const whole = idx === -1 ? d : d.slice(0, idx);
  const frac = idx === -1 ? "" : d.slice(idx + 1);
  if (frac.length <= places) return whole + (places > 0 ? "." + frac.padEnd(places, "0") : "");
  const digits = (whole + frac.slice(0, places)).split("");
  if (Number(frac[places]) >= 5) {
    let i = digits.length - 1;
    while (i >= 0) {
      if (digits[i] === "9") {
        digits[i] = "0";
        i--;
      } else {
        digits[i] = String(Number(digits[i]) + 1);
        break;
      }
    }
    if (i < 0) digits.unshift("1");
  }
  const all = digits.join("");
  let w = all.slice(0, all.length - places);
  if (w === "") w = "0";
  return places > 0 ? w + "." + all.slice(-places) : w;
}

/** PHP's `(string)` of a scalar `t`. */
function phpString(v: unknown): string {
  if (typeof v === "string") return v;
  if (v === null || v === undefined || v === false) return "";
  if (v === true) return "1";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : formatFloat(v);
  return "";
}

/** PHP truthiness: `""`, `"0"`, `0`, `0.0`, `null`, `false` and `[]` are false. */
function phpTruthy(v: unknown): boolean {
  if (v === null || v === undefined || v === false || v === 0 || v === "" || v === "0") return false;
  if (Array.isArray(v)) return v.length > 0;
  if (typeof v === "object") return Object.keys(v as object).length > 0;
  return !Number.isNaN(v);
}

/** PHP `(int)` for a span: numbers truncate, numeric strings parse as PHP does, bools are 0/1. */
function phpIntCast(v: unknown): number {
  if (typeof v === "number") return Number.isFinite(v) ? Math.trunc(v) : 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (typeof v === "string") return phpIsNumeric(v) ? Math.trunc(Number(v.trim())) : phpInt(v);
  return v === null || v === undefined ? 0 : 1;
}

/**
 * PHP `json_encode()` with default flags, for the scalars a malformed cell can
 * carry: `/` is escaped and non-ASCII is `\uXXXX`, as PHP writes them. A float
 * with an integral value is written as an integer, because this port cannot
 * tell `1.0` from `1`.
 */
function phpJsonEncode(v: unknown): string {
  if (v === undefined) return "null";
  const json = JSON.stringify(v) ?? "null";
  return json
    .replace(/\//g, "\\/")
    .replace(/[\u0080-￿]/g, (ch) => "\\u" + ch.charCodeAt(0).toString(16).padStart(4, "0"));
}

/** UTF-16 code unit order, PHP `Json::sortedKeys`' comparison. */
function compareCodeUnits(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0;
}
