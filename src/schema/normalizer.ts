import { Cell } from "../workbook/cell";
import { CellAddress } from "../workbook/cell-address";
import { ColumnWidths } from "./column-widths";
import { CellComment } from "../workbook/cell-comment";
import { CellFormat } from "../workbook/cell-format";
import { MergedRegion } from "../workbook/merged-region";
import { Sheet, type NormalizedChart, type PrintSetup } from "../workbook/sheet";
import { ChartRange } from "../workbook/chart-range";
import { Workbook } from "../workbook/workbook";
import { DateConverter } from "../writer/format/date-converter";
import { SchemaException } from "../exceptions";
import { isNumericString, isPlainObject, numericStringToNumber, typeOf } from "../util";
import type { CellPrimitive } from "./types";
import { Theme } from "./theme";

type Any = any; // input is loose agent JSON

/** Schema → canonical Workbook. Mirrors PHP `Schema\Normalizer`. */
export class Normalizer {
  normalize(schema: Any): Workbook {
    const sheets: Sheet[] = [];
    for (const sheetSchema of schema.sheets as Any[]) {
      sheets.push(this.normalizeSheet(sheetSchema));
    }
    return new Workbook(sheets, (schema.meta as Record<string, unknown>) ?? {});
  }

  private normalizeSheet(sheet: Any): Sheet {
    const name = String(sheet.name);

    // NOTE: there is no separate "cells-only" path. A sheet with no columns
    // and no rows builds an empty table below and then takes the overlay, which
    // is the same code the combined case runs.
    //
    // It USED to return here the moment `cells` was set, and that early return
    // silently discarded `columns`, `rows`, `totals` and `theme` — a four-row
    // table with one styled title cell wrote a one-cell workbook, while
    // `validate()` reported no errors because the validator explicitly permits
    // both together. The PHP twin had the identical bug in the identical place;
    // see `Schema/Normalizer.php`.
    const cells = new Map<string, Cell>();
    const columns: Any[] = sheet.columns ?? [];
    const rows: Any[] = sheet.rows ?? [];
    const themeKey = sheet.theme ?? "default";
    const theme = new Theme(themeKey);
    let headerOffset = 0;

    const columnFormats = new Map<number, CellFormat | null>();
    const columnByHeader = new Map<string, number>();
    columns.forEach((columnDef, colIdx) => {
      columnFormats.set(colIdx, this.columnFormat(columnDef));
      if (isPlainObject(columnDef) && columnDef.header !== undefined) {
        columnByHeader.set(String(columnDef.header), colIdx);
      }
    });

    if (columns.length > 0) {
      columns.forEach((columnDef, col) => {
        const address = CellAddress.letter(col) + "1";
        const header = isPlainObject(columnDef) ? (columnDef.header ?? "") : String(columnDef);
        cells.set(address, new Cell(address, String(header), null, theme.headerFormat()));
      });
      headerOffset = 1;
    }

    rows.forEach((row: Any[], r) => {
      row.forEach((value, c) => {
        const address = CellAddress.letter(c) + (r + 1 + headerOffset);
        const columnFormat = columnFormats.get(c) ?? null;
        const rowBand = theme.dataFormat(r);
        const merged =
          rowBand !== null ? (columnFormat ? columnFormat.mergeWith(rowBand) : rowBand) : columnFormat;
        cells.set(address, this.buildCell(address, value, merged, name));
      });
    });

    if (
      sheet.totals &&
      isPlainObject(sheet.totals) &&
      Object.keys(sheet.totals).length > 0 &&
      rows.length > 0 &&
      columns.length > 0
    ) {
      const totalsRow = rows.length + 1 + headerOffset;
      const totalsTheme = theme.totalsFormat();
      const labelAddr = CellAddress.letter(0) + totalsRow;
      cells.set(labelAddr, new Cell(labelAddr, "Total", null, totalsTheme));

      for (const [headerKey, aggOp] of Object.entries(sheet.totals as Record<string, Any>)) {
        if (!columnByHeader.has(headerKey)) continue;
        const colIdx = columnByHeader.get(headerKey)!;
        const colLetter = CellAddress.letter(colIdx);
        const rangeStart = colLetter + (headerOffset + 1);
        const rangeEnd = colLetter + (rows.length + headerOffset);
        const func = String(aggOp).toUpperCase();
        if (!["SUM", "AVG", "COUNT", "MIN", "MAX"].includes(func)) continue;
        const excelFunc = func === "AVG" ? "AVERAGE" : func;
        const address = colLetter + totalsRow;
        const colFmt = columnFormats.get(colIdx) ?? null;
        const combined = colFmt ? colFmt.mergeWith(totalsTheme) : totalsTheme;
        cells.set(address, new Cell(address, null, `${excelFunc}(${rangeStart}:${rangeEnd})`, combined));
      }
    }

    // The overlay. Explicit cells win at their address; a format is MERGED
    // rather than swapped, so a cell that only sets `bold` keeps the theme's
    // banding and the column's currency format instead of dropping to bare.
    if (sheet.cells != null) {
      for (const [address, cell] of this.normalizeCellMap(sheet.cells, name)) {
        const beneath = cells.get(address);

        cells.set(
          address,
          beneath?.format
            ? new Cell(
                cell.address,
                cell.value,
                cell.formula,
                beneath.format.mergeWith(cell.format),
                cell.comment,
                cell.cachedValue,
              )
            : cell,
        );
      }
    }

    return new Sheet(
      name,
      cells,
      this.normalizeMerges(sheet.mergedRegions ?? []),
      this.normalizeWidths(sheet),
      toInt(sheet.frozenRows ?? 0),
      toInt(sheet.frozenCols ?? 0),
      normalizeRowHeights(sheet.rowHeights),
      normalizePrint(sheet.print),
      normalizeCharts(sheet.charts, name, cells),
    );
  }

  private normalizeCellMap(map: Record<string, Any>, sheetName?: string): Map<string, Cell> {
    const cells = new Map<string, Cell>();
    for (const [address, cellData] of Object.entries(map)) {
      // A bare scalar cell, e.g. {A1: 42} or {A1: "=SUM(B1:B5)"}.
      //
      // Without this branch every non-object cell read as EMPTY: `cellData.value`
      // on a number is `undefined`, so `{A1: 42}` wrote a blank cell and a bare
      // formula string vanished outright — no error, just a hole in the sheet.
      // PHP has always had this branch; the port dropped it.
      if (!isCellSpec(cellData)) {
        const [bare, formula] = promoteFormula(cellData);
        cells.set(
          address,
          new Cell(
            address,
            formula !== null ? null : this.coerceValue(bare, null, address, sheetName),
            formula,
          ),
        );
        continue;
      }

      const format = isPlainObject(cellData.format) ? CellFormat.fromInput(cellData.format) : null;
      // `comment` takes either a string or an object. The string form was
      // accepted by the validator and dropped here — the same silent-drop shape.
      const comment = isPlainObject(cellData.comment)
        ? new CellComment(
            String(cellData.comment.text ?? ""),
            cellData.comment.author ?? null,
            cellData.comment.color ?? null,
          )
        : typeof cellData.comment === "string"
          ? new CellComment(cellData.comment)
          : null;
      const rawValue = cellData?.value ?? null;
      cells.set(
        address,
        new Cell(
          address,
          this.coerceValue(rawValue, format, address, sheetName),
          cellData?.formula ?? null,
          format,
          comment,
          cellData?.computedValue ?? null,
        ),
      );
    }
    return cells;
  }

  private normalizeMerges(list: Any[]): MergedRegion[] {
    const out: MergedRegion[] = [];
    for (const m of list) {
      if (m && m.start !== undefined && m.end !== undefined) {
        out.push(new MergedRegion(String(m.start), String(m.end)));
      }
    }
    return out;
  }

  /**
   * The two ways a width can be stated, folded into one map.
   *
   * `columns[].width` widens the column at that POSITION; the sheet-level
   * `columnWidths` map keys the same columns by 0-based index. Both have been
   * described in the shared `skills/holy-sheet.schema.json` ("Same as
   * columnWidths but per-column") for as long as the field has existed, and only
   * the map was ever read — so a width written the way the schema documents it
   * emitted no `<cols>` element at all, with no error and no repair note
   * (holy-sheet#8, fixed in PHP 2.4.0; this is the twin).
   *
   * The MAP IS APPLIED LAST and wins, matching PHP. It is the mechanism that
   * already worked, so anyone who reached for it to work around this bug must not
   * now find a leftover `width` quietly overriding it.
   */
  private normalizeWidths(sheet: Record<string, Any>): Map<number, number> {
    const out = new Map<number, number>();

    for (const [colIdx, columnDef] of (sheet.columns ?? []).entries()) {
      // A string column (`columns: ["Account"]`) is a documented shorthand and
      // carries no width — reading `.width` off it must not throw.
      if (columnDef === null || typeof columnDef !== "object") continue;
      if ((columnDef as Record<string, Any>).width === undefined) continue;

      const index = ColumnWidths.index(colIdx);
      const width = ColumnWidths.width((columnDef as Record<string, Any>).width);
      if (index === null || width === null) continue;
      out.set(index, width);
    }

    for (const [index, width] of this.normalizeColumnWidths(sheet.columnWidths ?? {})) {
      out.set(index, width);
    }

    // `<col>` elements are expected in ascending order, and merging two sources
    // means insertion order is no longer column order.
    return new Map([...out.entries()].sort((a, b) => a[0] - b[0]));
  }

  private normalizeColumnWidths(widths: Record<string, Any>): Map<number, number> {
    const out = new Map<number, number>();
    for (const [key, px] of Object.entries(widths)) {
      // An entry that is not a column index and a width is skipped. `parseInt`
      // turned "abc" into NaN and wrote a NaN column (PHP 2.3.4 skips it too).
      const index = ColumnWidths.index(key);
      const width = ColumnWidths.width(px);
      if (index === null || width === null) continue;
      out.set(index, width);
    }
    return out;
  }

  private columnFormat(columnDef: Any): CellFormat | null {
    if (typeof columnDef === "string") return null;
    const type = columnDef.type ?? "auto";
    const decimals = columnDef.decimals != null ? toInt(columnDef.decimals) : null;
    const currency = columnDef.currency ?? null;

    switch (type) {
      case "integer":
        return new CellFormat({ displayFormat: "number", decimals: 0 });
      case "number":
        return decimals !== null ? new CellFormat({ displayFormat: "number", decimals }) : null;
      case "percent":
        return new CellFormat({ displayFormat: "percentage", decimals: decimals ?? 1 });
      case "currency":
        return new CellFormat({ displayFormat: "currency", decimals: decimals ?? 2, currency });
      case "date":
        return new CellFormat({ displayFormat: "date" });
      case "datetime":
        return new CellFormat({ displayFormat: "datetime" });
      default:
        return null;
    }
  }

  private buildCell(
    address: string,
    value: Any,
    columnFormat: CellFormat | null,
    sheetName?: string,
  ): Cell {
    if (isCellSpec(value)) {
      const cellFormat = isPlainObject(value.format) ? CellFormat.fromInput(value.format) : null;
      const merged = columnFormat ? columnFormat.mergeWith(cellFormat) : cellFormat;
      const comment = isPlainObject(value.comment)
        ? new CellComment(
            String(value.comment.text ?? ""),
            value.comment.author ?? null,
            value.comment.color ?? null,
          )
        : null;
      const rawValue = value.value ?? null;
      return new Cell(
        address,
        this.coerceValue(rawValue, merged, address, sheetName),
        value.formula ?? null,
        merged,
        comment,
        value.computedValue ?? null,
      );
    }

    const [bare, formula] = promoteFormula(value);

    return new Cell(
      address,
      formula !== null ? null : this.coerceValue(bare, columnFormat, address, sheetName),
      formula,
      columnFormat,
    );
  }

  private coerceValue(
    value: Any,
    format: CellFormat | null,
    address = "?",
    sheetName?: string,
  ): CellPrimitive {
    if (value === null || value === undefined) return null;

    const df = format?.displayFormat;
    if (df === "date" || df === "datetime") {
      if (value instanceof Date || (typeof value === "string" && value.trim() !== "")) {
        return DateConverter.toSerial(value, df === "datetime");
      }
    }

    if (typeof value === "string" && isNumericString(value)) {
      return numericStringToNumber(value);
    }

    if (value instanceof Date) {
      return DateConverter.toSerial(value, true);
    }

    // A bigint is a primitive that converts losslessly and deliberately, and
    // PHP has no counterpart for it to diverge from. Unchanged.
    if (typeof value === "bigint") {
      return String(value);
    }

    // PHP's `instanceof Stringable` — an object that says how it becomes a
    // string. Two exclusions, and BOTH are load-bearing — the first draft of
    // this guard omitted the array one and its own comment claimed otherwise,
    // which the test caught:
    //
    //   - Gated on `typeof === "object"` so a FUNCTION does not qualify.
    //     `Function.prototype.toString` IS its own and returns the function's
    //     SOURCE TEXT, so a bare "has a custom toString" test writes a code
    //     listing into the cell.
    //   - `Array.isArray` excluded explicitly, because
    //     `Array.prototype.toString` is ALSO its own — and its joining is
    //     precisely the corruption this guard exists to stop. "Has a custom
    //     toString" is true of the one value that must never pass here.
    if (typeof value === "object" && !Array.isArray(value) && hasOwnToString(value)) {
      return String(value);
    }

    // REFUSE rather than cast. This used to read "not a primitive, so it must
    // be safe to stringify", and an ARRAY reaches it whenever a plain `rows`
    // form is nested one level too deep -- no object-cell syntax required,
    // which is exactly the shape an AI-generated edit produces.
    //
    // In THIS runtime the cast was the dangerous kind of wrong: `["x"]` became
    // "x", which is what the author probably meant, and `["x","y"]` became
    // "x,y" -- two cells merged into one, with the data intact enough to look
    // deliberate. The PHP twin wrote the visibly-wrong "Array", which someone
    // files a bug about. holy-sheet-js#2, mirroring holy-sheet#9 from MOIC
    // production.
    //
    // The path is the cell's `Sheet!A1` location, not the `sheets[0].rows[1]`
    // schema path Validator uses. Deliberate, and the one place the two
    // conventions differ: this fault is always ONE CELL, and the A1 location is
    // what an operator and an agent can both act on -- the production report
    // was that the error named no cell, not that it failed.
    if (typeof value !== "string" && typeof value !== "number" && typeof value !== "boolean") {
      const where = sheetName === undefined ? address : `${sheetName}!${address}`;
      const got = typeOf(value);

      throw SchemaException.fromErrors([
        {
          path: where,
          expected: "a string, number, bool, null, Date, or an object with its own toString",
          got,
          value,
          hint: Array.isArray(value)
            ? 'This cell holds an array. The usual cause is a `rows` entry nested one level too deep: `rows: [[["x"]]]` makes cell A1 an array, where `rows: [["x"]]` makes it the string "x". For a cell with formatting use the object form: {value: "x", format: {...}}.'
            : "Give the cell a scalar or a Date, or define toString() on the object so the conversion is yours rather than implicit.",
        },
      ]);
      // No second argument: `fromErrors()` takes only the error list and
      // ignores a message, matching PHP. `summarize()` already renders
      // "schema invalid at {path}: expected X, got Y", the format every other
      // schema error in this package uses.
    }

    return value;
  }
}

/**
 * Is this the `{value, format, comment, ...}` CELL SPEC, rather than a value
 * that merely happens to be an object?
 *
 * This replaced `isPlainObject`, which is `typeof === "object" && !null &&
 * !Array.isArray` and therefore true for EVERY non-array object — so a `Date`,
 * a class instance and a `Map` were all read as a cell spec, their `.value` was
 * `undefined`, and the cell came out EMPTY. Measured against the PHP reference
 * on `rows: [[new Date(2026, 0, 2)]]`: PHP wrote the serial 46024, this port
 * wrote an empty cell. A date in a plain row is the most ordinary non-string
 * cell there is, and the byte-parity suite never saw it because its fixtures
 * carry ISO date STRINGS rather than Date objects.
 *
 * PHP's test is `is_array($v) && !array_is_list($v)` — an associative array.
 * The faithful mirror is a plain object LITERAL, which is also what
 * `JSON.parse` produces for every schema that arrives as JSON. Anything with
 * its own prototype (Date, Map, a class instance) is a VALUE and belongs to
 * `coerceValue`; a null-prototype object is not a cell spec either, and is
 * refused there rather than silently emptied.
 */
function isCellSpec(value: unknown): value is Record<string, Any> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return false;
  return Object.getPrototypeOf(value) === Object.prototype;
}

/**
 * Does this object define how it becomes a string, rather than inheriting
 * `Object.prototype.toString` (which yields "[object Object]") or having no
 * `toString` at all (a null-prototype object, which THROWS on conversion)?
 *
 * Deliberately not `"toString" in value`: that is true for every object.
 */
function hasOwnToString(value: object): boolean {
  const fn = (value as { toString?: unknown }).toString;
  return typeof fn === "function" && fn !== Object.prototype.toString;
}

/**
 * A bare string cell value beginning with "=" (e.g. "=SUM(B2:B10)") is an Excel
 * formula, not literal text — promote it to a real formula cell.
 *
 * Only *bare* strings promote. An object cell ({value: "=x"} or {formula: "x"})
 * is always taken as the caller's explicit intent, so {value: "=literal"} is the
 * escape hatch for a genuine leading-"=" string. Mirrors PHP
 * `Normalizer::promoteFormula`, which the port had omitted entirely: the same
 * schema produced a formula in PHP and the literal text "=SUM(B2:B10)" in Node.
 */
function promoteFormula(value: Any): [Any, string | null] {
  if (typeof value === "string" && value.length > 1 && value[0] === "=") {
    return [null, value.slice(1)];
  }
  return [value, null];
}

function toInt(v: unknown): number {
  const n = Math.trunc(Number(v));
  return Number.isFinite(n) ? n : 0;
}

/** The chart types the writer draws. Mirrors PHP Normalizer::CHART_TYPES. */
export const CHART_TYPES = ["column", "bar", "line", "pie"] as const;

/** `charts`, normalised. Mirrors PHP Normalizer::normalizeCharts rule for rule. */
export function normalizeCharts(charts: unknown, sheetName: string, cells: Map<string, Cell>): NormalizedChart[] {
  if (!isPlainObject(charts) && !Array.isArray(charts)) return [];
  let lastCol = -1;
  for (const address of cells.keys()) {
    const parsed = CellAddress.parse(address);
    if (parsed !== null) lastCol = Math.max(lastCol, parsed[0]);
  }

  const out: NormalizedChart[] = [];
  for (const chart of Object.values(charts as Record<string, unknown>)) {
    if (!isPlainObject(chart) || !(CHART_TYPES as readonly unknown[]).includes(chart.type)) continue;
    const series: NormalizedChart["series"] = [];
    const rawSeries = isPlainObject(chart.series) || Array.isArray(chart.series) ? Object.values(chart.series as Record<string, unknown>) : [];
    for (const s of rawSeries) {
      const values = isPlainObject(s) || Array.isArray(s) ? ChartRange.parse((s as Record<string, unknown>).values, sheetName) : null;
      if (values === null) continue;
      const rec = s as Record<string, unknown>;
      const m = typeof rec.color === "string" ? /^#([0-9A-Fa-f]{6})$/.exec(rec.color) : null;
      series.push({ name: typeof rec.name === "string" ? rec.name : null, values, color: m ? m[1]!.toUpperCase() : null });
    }
    if (series.length === 0) continue;

    const stack = out.length;
    let from: [number, number] = [lastCol + 2, 1 + 18 * stack];
    let to: [number, number] = [from[0] + 8, from[1] + 16];
    if (isPlainObject(chart.position) || Array.isArray(chart.position)) {
      const p = chart.position as Record<string, unknown>;
      const a = typeof p.from === "string" ? CellAddress.parse(p.from) : null;
      const b = typeof p.to === "string" ? CellAddress.parse(p.to) : null;
      if (a !== null && b !== null && b[0] > a[0] && b[1] > a[1]) {
        from = [a[0], a[1] - 1];
        to = [b[0], b[1] - 1];
      }
    }

    out.push({
      type: chart.type as NormalizedChart["type"],
      title: typeof chart.title === "string" && chart.title !== "" ? chart.title : null,
      categories: ChartRange.parse(chart.categories, sheetName),
      series,
      from,
      to,
      legend: typeof chart.legend === "boolean" ? chart.legend : null,
      dataLabels: chart.dataLabels === true,
      format: typeof chart.format === "string" && chart.format !== "" ? chart.format : null,
    });
  }
  return out;
}

/** PHP `is_numeric`: a finite number or a numeric string. */
function phpNumeric(v: unknown): number | null {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  return isNumericString(v) ? numericStringToNumber(v as string) : null;
}

/** `rowHeights`: 1-based row number → pixels, sorted. Mirrors PHP Normalizer::normalizeRowHeights. */
export function normalizeRowHeights(map: unknown): Map<number, number> {
  const out = new Map<number, number>();
  if (!isPlainObject(map) && !Array.isArray(map)) return out;
  for (const [row, px] of Object.entries(map as Record<string, unknown>)) {
    const n = phpNumeric(px);
    if (/^[1-9]\d{0,6}$/.test(row) && n !== null && n >= 0) out.set(parseInt(row, 10), n);
  }
  return new Map([...out.entries()].sort((a, b) => a[0] - b[0]));
}

/** `print`, normalised; null when it asks for nothing. Mirrors PHP Normalizer::normalizePrint. */
export function normalizePrint(print: unknown): PrintSetup | null {
  if (!isPlainObject(print) && !Array.isArray(print)) return null;
  const p = print as Record<string, unknown>;
  const orientation = p.orientation === "portrait" || p.orientation === "landscape" ? p.orientation : null;
  const fitToWidth = p.fitToWidth === true;
  let footer: PrintSetup["footer"] = null;
  if (isPlainObject(p.footer) || Array.isArray(p.footer)) {
    const f = p.footer as Record<string, unknown>;
    const slot = (k: string): string => (typeof f[k] === "string" ? (f[k] as string) : "");
    footer = { left: slot("left"), center: slot("center"), right: slot("right") };
    if (footer.left === "" && footer.center === "" && footer.right === "") footer = null;
  }
  if (orientation === null && !fitToWidth && footer === null) return null;
  return { orientation, fitToWidth, footer };
}
