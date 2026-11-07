import { ArchiveRefused, ReadLimitExceeded, type ReadLimitName } from "../exceptions";
import { CellAddress } from "../workbook/cell-address";
import { DEFAULT_ARCHIVE_LIMITS, resolveArchiveLimits, type ArchiveLimits } from "../zip/source-archive";
import { scanXml } from "./xml";

/** Positive integers keyed by limit name; any subset overrides the defaults. */
export interface ReadLimitOptions {
  maxRows?: number;
  maxColumns?: number;
  maxCells?: number;
  maxSharedStrings?: number;
}

/** What `read()` / `describe()` take as `limits`: the read limits and the archive limits. */
export type AllReadLimits = ReadLimitOptions & ArchiveLimits;

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

  private constructor(
    private readonly limits: Required<ReadLimitOptions>,
    /** The archive limits, resolved over their defaults. */
    readonly archive: Required<ArchiveLimits>,
  ) {}

  /**
   * Any of the nine limits, by name (since 2.14.0 the archive limits too).
   * A read limit that is not a positive integer, or a name that is no limit,
   * is `ArchiveRefused("read_limits")`; a bad archive limit is
   * `ArchiveRefused("archive_limits")`. The message names the setting. Every
   * key is checked before anything is opened. Mirrors PHP `ReadLimits::from`.
   */
  static from(options: unknown = {}): ReadLimits {
    if (options === null || typeof options !== "object" || Array.isArray(options)) throw new ArchiveRefused("read_limits");
    const read: Record<string, unknown> = {};
    const archive: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(options)) {
      if (Object.prototype.hasOwnProperty.call(DEFAULT_ARCHIVE_LIMITS, key)) {
        archive[key] = value;
        continue;
      }
      const known = Object.prototype.hasOwnProperty.call(DEFAULT_READ_LIMITS, key);
      if (!known || !Number.isSafeInteger(value) || (value as number) < 1) throw ArchiveRefused.badSetting("read_limits", key, known);
      read[key] = value;
    }
    return new ReadLimits({ ...DEFAULT_READ_LIMITS, ...(read as ReadLimitOptions) }, resolveArchiveLimits(archive));
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
   * Count a worksheet part BEFORE it is parsed: the highest row a `<row r>` or
   * a cell uses, the highest column a cell or a `<col max>` uses, and every
   * `<c>` element. Streams the text (`scanXml`), building no nodes: until
   * 2.14.0 this walked the parsed tree, so a limit refused a 10,000,000-cell
   * sheet only after the whole tree was built. Matches what the parse reads:
   * the root's first `cols` and `sheetData`, their `col` / `row` children, and
   * a row's `c` children, by local name. The cell count refuses the moment
   * the workbook passes `maxCells` (the message says "at least"). Mirrors PHP
   * `ReadLimits::worksheetPart`.
   */
  worksheetPart(src: string): void {
    let rows = 0;
    let columns = 0;
    let cells = 0;
    let seenCols = false;
    let seenSheetData = false;
    let inside: "cols" | "sheetData" | null = null;
    let inRow = false;
    scanXml(src, 3, (depth, name, attrs) => {
      if (depth === 1) {
        inside = null;
        if (name === "cols" && !seenCols) [inside, seenCols] = ["cols", true];
        else if (name === "sheetData" && !seenSheetData) [inside, seenSheetData] = ["sheetData", true];
        return;
      }
      if (depth === 2) {
        inRow = false;
        if (inside === "cols" && name === "col") columns = Math.max(columns, attrNumber(attrs["max"]));
        else if (inside === "sheetData" && name === "row") {
          rows = Math.max(rows, attrNumber(attrs["r"]));
          inRow = true;
        }
        return;
      }
      if (depth !== 3 || !inRow || name !== "c") return;
      if (++cells + this.cells > this.limits.maxCells) {
        throw new ReadLimitExceeded("maxCells", this.limits.maxCells, cells + this.cells);
      }
      [columns, rows] = this.cellExtent(attrs["r"] ?? "", columns, rows);
    });
    this.sheet(rows, columns, cells);
  }

  /** Count a shared-strings part's `si` entries BEFORE it is parsed. Mirrors PHP `ReadLimits::sharedStringsPart`. */
  sharedStringsPart(src: string): void {
    let count = 0;
    scanXml(src, 1, (depth, name) => {
      if (depth === 1 && name === "si") count++;
    });
    this.sharedStrings(count);
  }

  /** The extent one cell address adds: `[columns, rows]` after it. */
  private cellExtent(r: string, columns: number, rows: number): [number, number] {
    const m = /^([A-Z]{1,7})([0-9]{1,15})$/.exec(r);
    if (m) {
      // The canonical form every writer emits; address() gives the same answer.
      let col = this.letterCache.get(m[1]!);
      if (col === undefined) {
        col = CellAddress.index(m[1]!);
        if (this.letterCache.size < 20000) this.letterCache.set(m[1]!, col);
      }
      return [Math.max(columns, col + 1), Math.max(rows, Number(m[2]))];
    }
    const at = readAddress(r);
    return at === null ? [columns, rows] : [Math.max(columns, at[0] + 1), Math.max(rows, at[1])];
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
