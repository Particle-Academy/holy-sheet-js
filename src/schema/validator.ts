import { SchemaException } from "../exceptions";
import { isPlainObject, typeOf } from "../util";
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
          }
        });
      }
    }

    if (sheet.cells != null) {
      if (!isPlainObject(sheet.cells)) {
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

    return errors;
  }
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
