import { Inference } from "../schema/inference";
import type { BuilderOptions, ColumnSchema, HolySheetSchema, SheetSchema } from "../schema/types";

type Any = any;

/**
 * Flat rows (+ optional headers) → Holy Sheet schema with inferred column types.
 * Mirrors PHP `Helpers\ArrayBuilder`.
 */
export const ArrayBuilder = {
  build(
    rows: Any[][],
    headers: string[] | null = null,
    sheetName = "Sheet 1",
    options: BuilderOptions = {},
  ): HolySheetSchema {
    let hdrs = headers;
    let data = rows;
    if (hdrs === null) {
      if (rows.length === 0) return wrap([], [], sheetName, options);
      hdrs = (rows[0] ?? []).map((v) => String(v));
      data = rows.slice(1);
    }
    const columns = inferColumns(data, hdrs, options);
    return wrap(columns, data, sheetName, options);
  },
};

function inferColumns(rows: Any[][], headers: string[], options: BuilderOptions): ColumnSchema[] {
  return headers.map((headerName, i) => {
    const columnValues = rows.map((row) => row[i] ?? null);
    return Inference.detect(columnValues, String(headerName), options);
  });
}

function wrap(
  columns: ColumnSchema[],
  rows: Any[][],
  sheetName: string,
  options: BuilderOptions,
): HolySheetSchema {
  const sheet: SheetSchema = { name: sheetName, columns, rows };
  if (options.theme !== undefined) sheet.theme = options.theme;
  if (options.totals && typeof options.totals === "object") sheet.totals = options.totals;
  if (options.frozenRows !== undefined) sheet.frozenRows = Math.trunc(Number(options.frozenRows));
  if (options.frozenCols !== undefined) sheet.frozenCols = Math.trunc(Number(options.frozenCols));
  return { sheets: [sheet] };
}
