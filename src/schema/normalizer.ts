import { Cell } from "../workbook/cell";
import { CellAddress } from "../workbook/cell-address";
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

    if (sheet.cells != null) {
      return new Sheet(
        name,
        this.normalizeCellMap(sheet.cells),
        this.normalizeMerges(sheet.mergedRegions ?? []),
        this.normalizeColumnWidths(sheet.columnWidths ?? {}),
        toInt(sheet.frozenRows ?? 0),
        toInt(sheet.frozenCols ?? 0),
      );
    }

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
      const format =
        cellData && isPlainObject(cellData.format) ? CellFormat.fromInput(cellData.format) : null;
      const comment =
        cellData && isPlainObject(cellData.comment)
          ? new CellComment(
              String(cellData.comment.text ?? ""),
              cellData.comment.author ?? null,
              cellData.comment.color ?? null,
            )
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
      out.set(parseInt(key, 10), Number(px));
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

    return new Cell(address, this.coerceValue(value, columnFormat), null, columnFormat);
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

function toInt(v: unknown): number {
  const n = Math.trunc(Number(v));
  return Number.isFinite(n) ? n : 0;
}
