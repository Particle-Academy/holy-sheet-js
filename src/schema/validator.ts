import { SchemaException } from "../exceptions";
import { isNumericString, isPlainObject, numericStringToNumber, typeOf } from "../util";
import { CellAddress } from "../workbook/cell-address";
import { ChartRange } from "../workbook/chart-range";
import { ColumnWidths } from "./column-widths";
import { CHART_TYPES } from "./normalizer";
import { Repairer } from "./repairer";
import type { RepairResult, ValidationError } from "./types";

const VALID_THEMES = ["default", "minimal", "plain", "business"];
const ALLOWED_TYPES = [
  "auto",
  "string",
  "number",
  "integer",
  "boolean",
  "date",
  "datetime",
  "currency",
  "percent",
  "formula",
];

type Any = any;

/** Hand-rolled schema validator. Mirrors PHP `Schema\Validator`. */
export class Validator {
  validate(schema: Any): ValidationError[] {
    const errors: ValidationError[] = [];

    if (schema == null || !Array.isArray(schema.sheets)) {
      errors.push(
        error(
          "sheets",
          "array",
          typeOf(schema?.sheets ?? null),
          schema?.sheets ?? null,
          'Top-level "sheets" must be an array of sheet definitions.',
        ),
      );
      return errors;
    }

    if (schema.sheets.length === 0) {
      errors.push(
        error("sheets", "non-empty array", "empty array", [], "A workbook must contain at least one sheet."),
      );
      return errors;
    }

    schema.sheets.forEach((sheet: Any, i: number) => {
      errors.push(...this.validateSheet(sheet, `sheets[${i}]`));
    });

    return errors;
  }

  assert(schema: Any): void {
    const errors = this.validate(schema);
    if (errors.length > 0) throw SchemaException.fromErrors(errors);
  }

  validateAndRepair(schema: Any): RepairResult {
    const [repairedSchema, repairs] = new Repairer().repair(schema);
    const errors = this.validate(repairedSchema);
    return { schema: repairedSchema, errors, repairs };
  }

  private validateSheet(sheet: Any, path: string): ValidationError[] {
    const errors: ValidationError[] = [];

    if (!isPlainObject(sheet)) {
      return [
        error(
          path,
          "object",
          typeOf(sheet),
          sheet,
          'Each sheet must be an object with at least a "name" key.',
        ),
      ];
    }

    if (typeof sheet.name !== "string" || sheet.name.trim() === "") {
      errors.push(
        error(
          `${path}.name`,
          "non-empty string",
          typeOf(sheet.name ?? null),
          sheet.name ?? null,
          "Sheet name is required and visible in Excel's tab strip.",
        ),
      );
    }

    const hasColumnsRows = sheet.rows != null || sheet.columns != null;
    const hasCells = sheet.cells != null;

    if (!hasColumnsRows && !hasCells) {
      errors.push(
        error(
          path,
          "object with rows OR cells",
          "object without either",
          sheet,
          "A sheet needs either {columns, rows} (row-oriented) or {cells} (sparse A1-keyed) data.",
        ),
      );
    }

    if (sheet.columns != null) {
      if (!Array.isArray(sheet.columns)) {
        errors.push(
          error(
            `${path}.columns`,
            "array",
            typeOf(sheet.columns),
            sheet.columns,
            "Columns must be an array of column definitions.",
          ),
        );
      } else {
        sheet.columns.forEach((col: Any, j: number) => {
          errors.push(...this.validateColumn(col, `${path}.columns[${j}]`));
        });
      }
    }

    if (sheet.rows != null) {
      if (!Array.isArray(sheet.rows)) {
        errors.push(
          error(`${path}.rows`, "array", typeOf(sheet.rows), sheet.rows, "Rows must be an array of arrays."),
        );
      } else {
        sheet.rows.forEach((row: Any, j: number) => {
          if (!Array.isArray(row)) {
            errors.push(
              error(
                `${path}.rows[${j}]`,
                "array",
                typeOf(row),
                row,
                "Each row is an array of cell values, in column order.",
              ),
            );
            return;
          }

          // An ARRAY as a cell value is the holy-sheet#9 shape: a row nested one
          // level too deep. It must be caught HERE and not left to the writer,
          // because `validate()` returning no errors for a schema `normalize()`
          // throws on tells a caller the input is fine and moves the failure to
          // where it is hardest to attribute. An agent that validates before
          // writing — which is the whole point of `Agent.validate()` — was being
          // told to go ahead.
          //
          // The test is `Array.isArray`, NOT "is an object": the object cell
          // form {value, format} has always been legal, and rejecting any
          // object here would trade this bug for a far worse one — every
          // formatted cell in every workbook. (PHP distinguishes the two as
          // `array_is_list` vs an associative array; in JS they are already
          // different types.) There is a test holding that line.
          row.forEach((cell: Any, k: number) => {
            if (Array.isArray(cell)) {
              errors.push(
                error(
                  `${path}.rows[${j}][${k}]`,
                  "a cell value or a cell object",
                  "array",
                  cell,
                  'This cell is an array, which usually means the row is nested one level too deep: `rows: [[["x"]]]` makes cell A1 an array, where `rows: [["x"]]` makes it "x". For a cell with formatting use the object form: {value: "x", format: {...}}.',
                ),
              );
            }
          });
        });
      }
    }

    if (sheet.cells != null) {
      // `[]` is an empty map in PHP terms: PHP's describe() reports a sheet with
      // no cells that way, so rejecting it broke describe() -> write().
      const emptyList = Array.isArray(sheet.cells) && sheet.cells.length === 0;
      if (!isPlainObject(sheet.cells) && !emptyList) {
        errors.push(
          error(
            `${path}.cells`,
            "object keyed by A1 address",
            typeOf(sheet.cells),
            sheet.cells,
            'Cells must be an object/map keyed by A1 references like "A1", "B2".',
          ),
        );
      }
    }

    if (sheet.theme != null && !VALID_THEMES.includes(sheet.theme)) {
      errors.push(
        error(
          `${path}.theme`,
          "one of: default, minimal, plain, business",
          "unknown",
          sheet.theme,
          "Pick a built-in theme or omit for default.",
        ),
      );
    }

    if (sheet.columnWidths != null) {
      if (!isPlainObject(sheet.columnWidths) && !Array.isArray(sheet.columnWidths)) {
        errors.push(
          error(
            `${path}.columnWidths`,
            "object keyed by 0-based column index",
            typeOf(sheet.columnWidths),
            sheet.columnWidths,
            'Column widths map a 0-based column index to pixels: {"0": 120, "1": 80}.',
          ),
        );
      } else {
        for (const [key, px] of Object.entries(sheet.columnWidths as Record<string, unknown>)) {
          if (ColumnWidths.index(key) === null) {
            errors.push(
              error(
                `${path}.columnWidths.${key}`,
                `a 0-based column index from 0 to ${ColumnWidths.MAX_INDEX}`,
                typeOf(key),
                key,
                'Keys are 0-based column indexes: "0" is column A, "1" is column B. Use the index, not the letter.',
              ),
            );
          } else if (ColumnWidths.width(px) === null) {
            errors.push(
              error(
                `${path}.columnWidths.${key}`,
                "a non-negative number of pixels",
                typeOf(px),
                px,
                "A width is a number of pixels, like 120.",
              ),
            );
          }
        }
      }
    }

    // Mirrors PHP Validator, path and message for path and message.
    if (sheet.rowHeights != null) {
      if (!isPlainObject(sheet.rowHeights) && !Array.isArray(sheet.rowHeights)) {
        errors.push(
          error(`${path}.rowHeights`, "object keyed by 1-based row number", typeOf(sheet.rowHeights), sheet.rowHeights,
            'Row heights map a row NUMBER to pixels: {"1": 32, "4": 28}.'),
        );
      } else {
        for (const [row, px] of Object.entries(sheet.rowHeights as Record<string, unknown>)) {
          const n = typeof px === "number" ? px : isNumericString(px) ? numericStringToNumber(px as string) : null;
          if (!/^[1-9]\d{0,6}$/.test(row)) {
            errors.push(error(`${path}.rowHeights.${row}`, "a 1-based row number", typeOf(row), row,
              'Keys are row numbers as Excel shows them: "1" is the first row.'));
          } else if (n === null || !Number.isFinite(n) || n < 0) {
            errors.push(error(`${path}.rowHeights.${row}`, "a non-negative number of pixels", typeOf(px), px,
              "A height is a number of pixels, like 28 (20 is a default row)."));
          }
        }
      }
    }

    if (sheet.charts != null) {
      errors.push(...validateCharts(sheet.charts, path, typeof sheet.name === "string" ? sheet.name : ""));
    }

    if (sheet.print != null) {
      const print = sheet.print as Record<string, unknown>;
      if (!isPlainObject(print) && !Array.isArray(print)) {
        errors.push(error(`${path}.print`, "object {orientation?, fitToWidth?, footer?}", typeOf(print), print,
          'Print setup: {"orientation": "landscape", "fitToWidth": true, "footer": {"right": "Page {page} of {pages}"}}.'));
      } else {
        if (print.orientation != null && print.orientation !== "portrait" && print.orientation !== "landscape") {
          errors.push(error(`${path}.print.orientation`, "one of: portrait, landscape", typeOf(print.orientation), print.orientation,
            "Use landscape for a sheet wider than it is tall."));
        }
        if (print.fitToWidth != null && typeof print.fitToWidth !== "boolean") {
          errors.push(error(`${path}.print.fitToWidth`, "boolean", typeOf(print.fitToWidth), print.fitToWidth,
            "true scales the sheet to one page wide when printed."));
        }
        if (print.fitToPage != null && typeof print.fitToPage !== "boolean") {
          errors.push(error(`${path}.print.fitToPage`, "boolean", typeOf(print.fitToPage), print.fitToPage,
            "true scales the sheet onto ONE page, wide and tall: a one-page summary with a chart prints whole."));
        }
        if (print.footer != null) {
          if (!isPlainObject(print.footer) && !Array.isArray(print.footer)) {
            errors.push(error(`${path}.print.footer`, "object {left?, center?, right?}", typeOf(print.footer), print.footer,
              "A footer has up to three slots of text."));
          } else {
            const footer = print.footer as Record<string, unknown>;
            for (const slot of ["left", "center", "right"]) {
              if (footer[slot] != null && typeof footer[slot] !== "string") {
                errors.push(error(`${path}.print.footer.${slot}`, "string", typeOf(footer[slot]), footer[slot],
                  "Footer text; {page} and {pages} print the page number and count."));
              }
            }
          }
        }
      }
    }

    return errors;
  }

  private validateColumn(col: Any, path: string): ValidationError[] {
    if (!isPlainObject(col)) {
      return [
        error(path, "object", typeOf(col), col, 'Each column is an object with at least a "header" key.'),
      ];
    }

    const errors: ValidationError[] = [];
    if (typeof col.header !== "string") {
      errors.push(
        error(
          `${path}.header`,
          "string",
          typeOf(col.header ?? null),
          col.header ?? null,
          "Column header is the visible label in row 1.",
        ),
      );
    }

    if (col.type !== undefined && !ALLOWED_TYPES.includes(col.type)) {
      errors.push(
        error(
          `${path}.type`,
          "one of: " + ALLOWED_TYPES.join(", "),
          "unknown",
          col.type,
          'Pick a supported type or omit for "auto" (inferred per cell).',
        ),
      );
    }

    // Until holy-sheet#8 the writer ignored `width` entirely, so an invalid one
    // and a valid one were equally silent. Now that it is honoured, a width that
    // cannot be used has to say so — the original defect's real cost was that an
    // agent composes against the schema doc, gets no error, and never learns the
    // field was discarded.
    if (col.width !== undefined && ColumnWidths.width(col.width) === null) {
      errors.push(
        error(
          `${path}.width`,
          "a non-negative number of pixels",
          typeOf(col.width),
          col.width,
          "A column width is a number of pixels, like 120. The sheet-level columnWidths map takes precedence over it.",
        ),
      );
    }

    return errors;
  }
}

/** Mirrors PHP Validator::validatePoints. */
function validatePoints(points: unknown, at: string, values: ChartRange | null): ValidationError[] {
  if (!Array.isArray(points)) {
    return [
      error(at, "a list of points", typeOf(points), points,
        'Colour single points: [{"index": 2, "color": "#C2410C"}], index 0-based within values.'),
    ];
  }
  const count = values === null ? 0 : values.addresses().length;
  const errors: ValidationError[] = [];
  points.forEach((p: unknown, k: number) => {
    const rec = isPlainObject(p) || Array.isArray(p) ? (p as Record<string, unknown>) : null;
    const index = rec ? (rec.index ?? null) : null;
    const color = rec ? (rec.color ?? null) : null;
    if (typeof index !== "number" || !Number.isInteger(index) || index < 0 || index >= count) {
      errors.push(error(`${at}[${k}].index`, "an index within values (0 to " + Math.max(0, count - 1) + ")", typeOf(index), index,
        'The 0-based position of the point in its values range: "B5:B7" has points 0, 1 and 2.'));
    }
    if (typeof color !== "string" || !/^#[0-9A-Fa-f]{6}$/.test(color)) {
      errors.push(error(`${at}[${k}].color`, "#RRGGBB", typeOf(color), color, 'A hex colour, like "#C2410C".'));
    }
  });
  return errors;
}

/** Mirrors PHP Validator::validateCharts, path, expectation and hint for each. */
function validateCharts(charts: unknown, path: string, sheetName: string): ValidationError[] {
  if (!Array.isArray(charts)) {
    return [
      error(`${path}.charts`, "a list of charts", typeOf(charts), charts,
        'charts is a list: [{"type": "column", "categories": "A5:A7", "series": [{"name": "Revenue", "values": "B5:B7"}]}].'),
    ];
  }
  const errors: ValidationError[] = [];
  charts.forEach((chart: unknown, i: number) => {
    const at = `${path}.charts[${i}]`;
    if (!isPlainObject(chart) && !Array.isArray(chart)) {
      errors.push(error(at, "object", typeOf(chart), chart, "A chart is an object with a type and series."));
      return;
    }
    const c = chart as Record<string, unknown>;
    if (!(CHART_TYPES as readonly unknown[]).includes(c.type)) {
      errors.push(error(`${at}.type`, "one of: " + CHART_TYPES.join(", "), typeOf(c.type ?? null), c.type ?? null,
        "column compares values across categories; bar does the same horizontally (long labels); line shows change over time; pie shows parts of one whole."));
    }
    const series = c.series;
    if (!Array.isArray(series) || series.length === 0) {
      errors.push(error(`${at}.series`, "a non-empty list of series", typeOf(series ?? null), series ?? null,
        'Each series names the range it plots: [{"name": "Revenue", "values": "B5:B7"}].'));
    } else {
      series.forEach((s: unknown, j: number) => {
        const isObj = isPlainObject(s) || Array.isArray(s);
        const values = isObj ? (s as Record<string, unknown>).values ?? null : s;
        if (!isObj || ChartRange.parse(values, sheetName) === null) {
          errors.push(error(`${at}.series[${j}].values`, 'a cell range like "B5:B7"', typeOf(values), values,
            "A range of cells: \"B5:B7\" on this sheet, or \"'Other sheet'!B5:B7\"."));
        } else {
          const color = (s as Record<string, unknown>).color;
          if (color != null && (typeof color !== "string" || !/^#[0-9A-Fa-f]{6}$/.test(color))) {
            errors.push(error(`${at}.series[${j}].color`, "#RRGGBB", typeOf(color), color, 'A hex colour, like "#1F3A5F".'));
          }
        }
        if (isObj && (s as Record<string, unknown>).points != null) {
          errors.push(...validatePoints((s as Record<string, unknown>).points, `${at}.series[${j}].points`, ChartRange.parse(values, sheetName)));
        }
      });
    }
    if (c.categories != null && ChartRange.parse(c.categories, sheetName) === null) {
      errors.push(error(`${at}.categories`, 'a cell range like "A5:A7"', typeOf(c.categories), c.categories,
        'The labels along the axis, as a range: "A5:A7".'));
    }
    if (c.position != null) {
      const p = c.position as Record<string, unknown>;
      const isObj = isPlainObject(p) || Array.isArray(p);
      const a = isObj && typeof p.from === "string" ? CellAddress.parse(p.from) : null;
      const b = isObj && typeof p.to === "string" ? CellAddress.parse(p.to) : null;
      if (a === null || b === null || b[0] <= a[0] || b[1] <= a[1]) {
        errors.push(error(`${at}.position`, '{"from": top-left cell, "to": bottom-right cell}', typeOf(p), p,
          'Two cells, the second below and to the right of the first: {"from": "E2", "to": "L18"} is 7 columns by 16 rows.'));
      }
    }
    for (const key of ["title", "format"]) {
      if (c[key] != null && typeof c[key] !== "string") {
        errors.push(error(`${at}.${key}`, "string", typeOf(c[key]), c[key],
          key === "format" ? 'An Excel number format for the axis and labels, like "$#,##0" or "0.0%".' : "The chart title."));
      }
    }
    for (const key of ["dataLabels", "legend"]) {
      if (c[key] != null && typeof c[key] !== "boolean") {
        errors.push(error(`${at}.${key}`, "boolean", typeOf(c[key]), c[key],
          key === "dataLabels" ? "true prints each value on the chart." : "Defaults to on for several series or a pie."));
      }
    }
  });
  return errors;
}

function error(
  path: string,
  expected: string,
  got: string,
  value: unknown,
  hint: string,
): ValidationError {
  return { path, expected, got, value, hint };
}
