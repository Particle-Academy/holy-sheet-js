import { isNumericString, isPlainObject, numericStringToNumber } from "../util";
import { ColumnWidths } from "./column-widths";
import { DV_OPERATORS, DV_STYLES, DV_TYPES } from "./data-validations";

const VALID_THEMES = ["default", "minimal", "plain", "business"];
const ISO_DATE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})?)?$/;
const NUMERIC_TYPES = ["number", "integer", "currency", "percent"];

type Any = any;

/**
 * Conservative, high-confidence schema repairs. Mirrors PHP `Schema\Repairer`.
 * Returns [repairedSchema, repairs].
 */
export class Repairer {
  private repairs: string[] = [];

  repair(schema: Any): [Record<string, unknown>, string[]] {
    this.repairs = [];
    let s = clone(schema);
    s = this.repairTopLevel(s);

    if (Array.isArray(s.sheets)) {
      s.sheets = s.sheets.map((sheet: Any, i: number) => this.repairSheet(sheet, `sheets[${i}]`));
    }

    return [s, this.repairs];
  }

  private repairTopLevel(schema: Any): Any {
    if (schema.sheets == null && schema.sheet != null) {
      const value = schema.sheet;
      delete schema.sheet;
      schema.sheets = Array.isArray(value) ? value : [value];
      this.repairs.push("renamed top-level 'sheet' → 'sheets'");
    }
    return schema;
  }

  private repairSheet(sheet: Any, path: string): Any {
    if (!isPlainObject(sheet)) return sheet;

    // 'row' → 'rows'
    if (sheet.rows == null && sheet.row != null) {
      sheet.rows = sheet.row;
      delete sheet.row;
      this.repairs.push(`renamed '${path}.row' → '${path}.rows'`);
    }

    // Object-as-list: rows is {0:[],1:[],2:[]} → indexed array
    if (sheet.rows != null && isPlainObject(sheet.rows)) {
      const keys = Object.keys(sheet.rows);
      const allIntegerKeys = keys.every((k) => /^\d+$/.test(k));
      if (allIntegerKeys) {
        sheet.rows = Object.values(sheet.rows);
        this.repairs.push(`converted '${path}.rows' from integer-keyed object to indexed list`);
      }
    }

    // Unknown theme → 'default'
    if (sheet.theme != null && !VALID_THEMES.includes(sheet.theme)) {
      const original = sheet.theme;
      sheet.theme = "default";
      this.repairs.push(`changed '${path}.theme' from '${original}' to 'default' (unknown theme)`);
    }

    // Column widths: a letter key becomes its index; an entry that is still not
    // a column index and a width is dropped, and said (PHP 2.3.4).
    if (sheet.columnWidths != null && (isPlainObject(sheet.columnWidths) || Array.isArray(sheet.columnWidths))) {
      const widths: Record<string, unknown> = {};
      let changed = false;
      for (const [key, px] of Object.entries(sheet.columnWidths as Record<string, unknown>)) {
        let index = ColumnWidths.index(key);
        if (index === null) {
          const letters = ColumnWidths.fromLetters(key);
          if (letters !== null) {
            index = letters;
            this.repairs.push(`converted '${path}.columnWidths.${key}' to column index ${index}`);
            changed = true;
          }
        }
        if (index === null || ColumnWidths.width(px) === null) {
          this.repairs.push(`dropped '${path}.columnWidths.${key}' (not a column index and a width)`);
          changed = true;
          continue;
        }
        widths[String(index)] = px;
      }
      if (changed) {
        sheet.columnWidths = widths;
      }
    }

    // Trim whitespace in sparse-cell A1 addresses
    if (sheet.cells != null && isPlainObject(sheet.cells)) {
      const cleaned: Record<string, unknown> = {};
      let changed = false;
      for (const [addr, data] of Object.entries(sheet.cells)) {
        const trimmed = addr.trim();
        if (trimmed !== addr) changed = true;
        cleaned[trimmed] = data;
      }
      if (changed) {
        sheet.cells = cleaned;
        this.repairs.push(`trimmed whitespace from cell addresses in '${path}.cells'`);
      }
    }

    // Data validation enums in the wrong case. Only an exact match ignoring
    // case is repaired; 'dropdown' is not 'list', and saying so is the
    // validator's job.
    if (Array.isArray(sheet.columns)) {
      sheet.columns.forEach((col: Any, colIdx: number) => {
        if (isPlainObject(col) && isPlainObject(col.validation)) {
          this.repairValidation(col.validation, `${path}.columns[${colIdx}].validation`);
        }
      });
    }
    if (Array.isArray(sheet.dataValidations)) {
      sheet.dataValidations.forEach((dv: Any, k: number) => {
        if (isPlainObject(dv)) this.repairValidation(dv, `${path}.dataValidations[${k}]`);
      });
    }

    // Column-driven repairs
    if (sheet.columns != null && sheet.rows != null && Array.isArray(sheet.columns) && Array.isArray(sheet.rows)) {
      this.repairColumnTypeInference(sheet, path);
      this.repairStringifiedNumerics(sheet);
    }

    return sheet;
  }

  private repairValidation(dv: Any, path: string): void {
    this.repairCase(dv, "type", DV_TYPES, `${path}.type`);
    this.repairCase(dv, "operator", DV_OPERATORS, `${path}.operator`);
    if (isPlainObject(dv.error)) this.repairCase(dv.error, "style", DV_STYLES, `${path}.error.style`);
  }

  private repairCase(box: Any, key: string, allowed: string[], path: string): void {
    const value = box[key];
    if (typeof value !== "string" || allowed.includes(value)) return;
    for (const candidate of allowed) {
      if (candidate.toLowerCase() === value.toLowerCase()) {
        box[key] = candidate;
        this.repairs.push(`changed '${path}' from '${value}' to '${candidate}'`);
        break;
      }
    }
  }

  private repairColumnTypeInference(sheet: Any, path: string): void {
    sheet.columns.forEach((col: Any, colIdx: number) => {
      if (!isPlainObject(col)) return;
      if (col.type !== undefined && col.type !== "auto") return;

      const values: unknown[] = [];
      for (const row of sheet.rows as Any[]) {
        if (Array.isArray(row) && colIdx < row.length) values.push(row[colIdx]);
      }
      const nonNull = values.filter((v) => v !== null && v !== undefined);
      if (nonNull.length === 0) return;

      const allDateLike = nonNull.every((v) => typeof v === "string" && ISO_DATE.test(v));
      if (allDateLike) {
        const type = String(nonNull[0]).includes("T") ? "datetime" : "date";
        sheet.columns[colIdx].type = type;
        this.repairs.push(`inferred '${path}.columns[${colIdx}].type' = '${type}' from row values`);
      }
    });
  }

  private repairStringifiedNumerics(sheet: Any): void {
    sheet.columns.forEach((col: Any, colIdx: number) => {
      if (!isPlainObject(col)) return;
      const type = col.type ?? "auto";
      if (!NUMERIC_TYPES.includes(type)) return;

      (sheet.rows as Any[]).forEach((row) => {
        if (!Array.isArray(row)) return;
        const value = row[colIdx];
        if (typeof value === "string" && isNumericString(value)) {
          row[colIdx] = numericStringToNumber(value);
        }
      });
    });
  }
}

function clone<T>(v: T): T {
  return typeof structuredClone === "function"
    ? structuredClone(v)
    : (JSON.parse(JSON.stringify(v)) as T);
}
