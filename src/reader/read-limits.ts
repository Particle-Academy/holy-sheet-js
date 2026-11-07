import { ArchiveRefused, ReadLimitExceeded, type ReadLimitName } from "../exceptions";
import { CellAddress } from "../workbook/cell-address";
import { el, els, type XmlNode } from "./xml";

/** Positive integers keyed by limit name; any subset overrides the defaults. */
export interface ReadLimitOptions {
  maxRows?: number;
  maxColumns?: number;
  maxCells?: number;
  maxSharedStrings?: number;
}

/**
 * The defaults. `maxRows` / `maxColumns` are Excel's own sheet limits;
 * `maxCells` admits the measured 300,000 x 50 workbook (15,000,050 cells) with
 * a third to spare; `maxSharedStrings` matches it. Mirrors PHP
 * `ReadLimits::DEFAULTS`, and the reasons are in the 2.13.0 CHANGELOG entry.
 */
export const DEFAULT_READ_LIMITS: Readonly<Required<ReadLimitOptions>> = Object.freeze({
  maxRows: 1048576,
  maxColumns: 16384,
  maxCells: 20000000,
  maxSharedStrings: 20000000,
});

/**
 * How big a spreadsheet a read will expand, beyond the archive guard's byte
 * limits. Mirrors PHP `Reader\ReadLimits`: per-sheet rows and columns, a
 * workbook-wide cell count, the shared-strings table; each checked before the
 * thing it counts is built. Configured like the archive limits: an unknown key
 * or anything but a positive integer is `ArchiveRefused("read_limits")`.
 *
 * @internal one instance per read; it carries the running cell count
 */
export class ReadLimits {
  private cells = 0;
  private readonly letterCache = new Map<string, number>();

  private constructor(private readonly limits: Required<ReadLimitOptions>) {}

  static from(options: unknown = {}): ReadLimits {
    if (options === null || typeof options !== "object" || Array.isArray(options)) throw new ArchiveRefused("read_limits");
    for (const [key, value] of Object.entries(options)) {
      if (!Object.prototype.hasOwnProperty.call(DEFAULT_READ_LIMITS, key) || !Number.isSafeInteger(value) || (value as number) < 1) {
        throw new ArchiveRefused("read_limits");
      }
    }
    return new ReadLimits({ ...DEFAULT_READ_LIMITS, ...(options as ReadLimitOptions) });
  }

  max(limit: ReadLimitName): number {
    return this.limits[limit];
  }

  sharedStrings(count: number): void {
    check("maxSharedStrings", this.limits.maxSharedStrings, count);
  }

  /** One sheet's extent and cell count, checked in order: rows, columns, cells. */
  sheet(rows: number, columns: number, cells: number): void {
    check("maxRows", this.limits.maxRows, rows);
    check("maxColumns", this.limits.maxColumns, columns);
    this.cells += cells;
    check("maxCells", this.limits.maxCells, this.cells);
  }

  /**
   * Count a parsed worksheet: the highest row a `<row r>` or a cell uses, the
   * highest column a cell or a `<col max>` uses, and every `<c>` element.
   */
  worksheet(xml: XmlNode): void {
    let rows = 0;
    let columns = 0;
    let cells = 0;
    for (const col of els(el(xml, "cols"), "col")) columns = Math.max(columns, attrNumber(col.attrs["max"]));
    for (const row of els(el(xml, "sheetData"), "row")) {
      rows = Math.max(rows, attrNumber(row.attrs["r"]));
      for (const c of els(row, "c")) {
        cells++;
        const r = c.attrs["r"] ?? "";
        const m = /^([A-Z]{1,7})([0-9]{1,15})$/.exec(r);
        if (m) {
          // The canonical form every writer emits; address() gives the same answer.
          let col = this.letterCache.get(m[1]!);
          if (col === undefined) {
            col = CellAddress.index(m[1]!);
            if (this.letterCache.size < 20000) this.letterCache.set(m[1]!, col);
          }
          columns = Math.max(columns, col + 1);
          rows = Math.max(rows, Number(m[2]));
          continue;
        }
        const at = readAddress(r);
        if (at !== null) {
          columns = Math.max(columns, at[0] + 1);
          rows = Math.max(rows, at[1]);
        }
      }
    }
    this.sheet(rows, columns, cells);
  }
}

function check(limit: ReadLimitName, max: number, actual: number): void {
  if (actual > max) throw new ReadLimitExceeded(limit, max, actual);
}

/** PHP's trim() set: space, tab, LF, CR, NUL, vertical tab. Never Unicode whitespace. */
function phpTrimAscii(s: string): string {
  return s.replace(/^[ \t\n\r\0\x0B]+|[ \t\n\r\0\x0B]+$/g, "");
}

/**
 * `[column index, row]` of an A1 address, or null. Mirrors PHP
 * `ReadLimits::address`: trimmed with trim()'s set, upper-cased ASCII only, at
 * most seven letters and fifteen digits so the numbers stay exact.
 */
export function readAddress(address: string): [number, number] | null {
  const m = /^([A-Z]{1,7})([0-9]{1,15})$/.exec(phpTrimAscii(address).replace(/[a-z]/g, (ch) => ch.toUpperCase()));
  return m ? [CellAddress.index(m[1]!), Number(m[2])] : null;
}

/** A row or column number from an attribute; mirrors PHP `ReadLimits::number`. */
function attrNumber(value: string | undefined): number {
  const v = phpTrimAscii(value ?? "");
  if (!/^[0-9]+$/.test(v)) return 0;
  return v.replace(/^0+/, "").length > 15 ? 1e15 : Number(v);
}
