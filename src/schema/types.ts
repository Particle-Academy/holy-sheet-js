/**
 * Public input schema + result types. Mirrors the PHP package's array shapes.
 * Inputs are intentionally permissive (agents emit loose JSON) — the Validator
 * is the gate, not the type system.
 */

export type ThemeKey = "default" | "minimal" | "plain" | "business";

export type ColumnType =
  | "auto"
  | "string"
  | "number"
  | "integer"
  | "boolean"
  | "date"
  | "datetime"
  | "currency"
  | "percent"
  | "formula";

export interface ColumnSchema {
  header: string;
  type?: ColumnType;
  currency?: string;
  decimals?: number;
  format?: string;
  width?: number;
}

export type DisplayFormat =
  | "auto"
  | "text"
  | "number"
  | "date"
  | "datetime"
  | "percentage"
  | "currency";

export interface CellFormatInput {
  bold?: boolean;
  italic?: boolean;
  textAlign?: "left" | "center" | "right";
  displayFormat?: DisplayFormat;
  decimals?: number;
  color?: string;
  backgroundColor?: string;
  fontSize?: number;
  borderTop?: string | null;
  borderRight?: string | null;
  borderBottom?: string | null;
  borderLeft?: string | null;
  currency?: string;
}

export interface CommentInput {
  text: string;
  author?: string;
  color?: string;
}

export type CellPrimitive = string | number | boolean | null;

export interface CellData {
  value?: CellPrimitive;
  formula?: string;
  computedValue?: CellPrimitive;
  format?: CellFormatInput;
  comment?: CommentInput;
}

export type RowCell = CellPrimitive | CellData;

export type AggOp = "sum" | "avg" | "count" | "min" | "max";

export interface SheetSchema {
  name: string;
  // row-oriented mode
  columns?: ColumnSchema[];
  rows?: RowCell[][];
  theme?: ThemeKey;
  totals?: Record<string, AggOp | string>;
  // sparse mode
  cells?: Record<string, CellData>;
  // shared
  mergedRegions?: { start: string; end: string }[];
  columnWidths?: Record<number | string, number>;
  frozenRows?: number;
  frozenCols?: number;
}

export interface HolySheetSchema {
  sheets: SheetSchema[];
  meta?: { creator?: string; created?: string };
}

export interface ValidationError {
  path: string;
  expected: string;
  got: string;
  value: unknown;
  hint: string;
}

export interface FormulaProblem {
  sheet: string;
  address: string;
  formula: string;
  error: string;
  hint: string;
}

export interface WriteResult {
  path: string;
  bytes: number;
  sheets: number;
}

export interface RepairResult {
  schema: Record<string, unknown>;
  errors: ValidationError[];
  repairs: string[];
}

/** Options passed through ArrayBuilder / CsvBuilder. */
export interface BuilderOptions {
  theme?: ThemeKey;
  currency?: string;
  totals?: Record<string, AggOp | string>;
  frozenRows?: number;
  frozenCols?: number;
  sheetName?: string;
  // CSV-specific
  delimiter?: string;
  enclosure?: string;
}
