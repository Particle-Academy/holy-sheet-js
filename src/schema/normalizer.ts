import { Cell } from "../workbook/cell";
import { CellAddress } from "../workbook/cell-address";
import { ColumnWidths } from "./column-widths";
import { CellComment } from "../workbook/cell-comment";
import { CellFormat } from "../workbook/cell-format";
import { MergedRegion } from "../workbook/merged-region";
import { Sheet } from "../workbook/sheet";
import { Workbook } from "../workbook/workbook";
import { DateConverter } from "../writer/format/date-converter";
import { isNumericString, isPlainObject, numericStringToNumber } from "../util";
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
        cells.set(address, this.buildCell(address, value, merged));
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
      for (const [address, cell] of this.normalizeCellMap(sheet.cells)) {
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
      this.normalizeColumnWidths(sheet.columnWidths ?? {}),
      toInt(sheet.frozenRows ?? 0),
      toInt(sheet.frozenCols ?? 0),
    );
  }

  private normalizeCellMap(map: Record<string, Any>): Map<string, Cell> {
    const cells = new Map<string, Cell>();
    for (const [address, cellData] of Object.entries(map)) {
      // A bare scalar cell, e.g. {A1: 42} or {A1: "=SUM(B1:B5)"}.
      //
      // Without this branch every non-object cell read as EMPTY: `cellData.value`
      // on a number is `undefined`, so `{A1: 42}` wrote a blank cell and a bare
      // formula string vanished outright — no error, just a hole in the sheet.
      // PHP has always had this branch; the port dropped it.
      if (!isPlainObject(cellData)) {
        const [bare, formula] = promoteFormula(cellData);
        cells.set(
          address,
          new Cell(address, formula !== null ? null : this.coerceValue(bare, null), formula),
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
          this.coerceValue(rawValue, format),
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

  private buildCell(address: string, value: Any, columnFormat: CellFormat | null): Cell {
    if (isPlainObject(value)) {
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
        this.coerceValue(rawValue, merged),
        value.formula ?? null,
        merged,
        comment,
        value.computedValue ?? null,
      );
    }

    const [bare, formula] = promoteFormula(value);

    return new Cell(
      address,
      formula !== null ? null : this.coerceValue(bare, columnFormat),
      formula,
      columnFormat,
    );
  }

  private coerceValue(value: Any, format: CellFormat | null): CellPrimitive {
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

    if (typeof value !== "string" && typeof value !== "number" && typeof value !== "boolean") {
      return String(value);
    }

    return value;
  }
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
