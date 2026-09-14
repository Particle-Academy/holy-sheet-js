import type { CellFormatInput, CellPrimitive, CommentInput, HolySheetSchema, SheetSchema } from "../schema/types";

/**
 * One op from `Agent.diff`, applied by `Agent.reduce`. Mirrors the PHP
 * `HolySheet\Ops\SheetOpSchema` variants: same `type`, same field names.
 *
 * `set_cell`, `set_range` and `set_workbook` are fancy-sheets' `SheetOp` shapes;
 * `set_workbook.data` here is a Holy Sheet schema, not fancy-sheets'
 * `WorkbookData`.
 */
export type SheetOp =
  | SetCellOp
  | SetRangeOp
  | SetWorkbookOp
  | ClearCellOp
  | InsertRowsOp
  | DeleteRowsOp
  | InsertColumnsOp
  | DeleteColumnsOp
  | AddSheetOp
  | RemoveSheetOp
  | RenameSheetOp
  | MoveSheetOp
  | ReplaceSheetOp
  | SetMergedRegionsOp
  | SetColumnWidthsOp
  | SetFrozenOp
  | SetMetaOp;

export type SheetOpType = SheetOp["type"];

/** Omitting `formula` or `computedValue` clears it; omitting `format` or `comment` keeps it; null clears it. */
export interface SetCellOp {
  type: "set_cell";
  sheet: string;
  address: string;
  value?: CellPrimitive;
  formula?: string;
  computedValue?: CellPrimitive;
  format?: CellFormatInput | null;
  comment?: CommentInput | null;
}

/** Values row-major from `start`, each written as a `set_cell` with no formula. `end` is accepted and not read. */
export interface SetRangeOp {
  type: "set_range";
  sheet: string;
  start: string;
  end?: string;
  values: CellPrimitive[][];
}

export interface SetWorkbookOp {
  type: "set_workbook";
  data: HolySheetSchema | Record<string, unknown>;
}

export interface ClearCellOp {
  type: "clear_cell";
  sheet: string;
  address: string;
}

/** Insert rows before 1-based row `at`. */
export interface InsertRowsOp {
  type: "insert_rows";
  sheet: string;
  at: number;
  count: number;
}

/** Delete rows starting at 1-based row `at`. */
export interface DeleteRowsOp {
  type: "delete_rows";
  sheet: string;
  at: number;
  count: number;
}

/** Insert columns before 1-based column `at` (A = 1). */
export interface InsertColumnsOp {
  type: "insert_columns";
  sheet: string;
  at: number;
  count: number;
}

/** Delete columns starting at 1-based column `at` (A = 1). */
export interface DeleteColumnsOp {
  type: "delete_columns";
  sheet: string;
  at: number;
  count: number;
}

/** Insert a sheet at a 0-based position. */
export interface AddSheetOp {
  type: "add_sheet";
  index: number;
  sheet: SheetSchema | Record<string, unknown>;
}

export interface RemoveSheetOp {
  type: "remove_sheet";
  sheet: string;
}

export interface RenameSheetOp {
  type: "rename_sheet";
  sheet: string;
  name: string;
}

/** Move a sheet to a 0-based position. */
export interface MoveSheetOp {
  type: "move_sheet";
  sheet: string;
  toIndex: number;
}

export interface ReplaceSheetOp {
  type: "replace_sheet";
  sheet: string;
  data: SheetSchema | Record<string, unknown>;
}

/** An empty list removes the key. */
export interface SetMergedRegionsOp {
  type: "set_merged_regions";
  sheet: string;
  mergedRegions: { start: string; end: string }[];
}

/** 0-based column index to pixels. An empty map removes the key. */
export interface SetColumnWidthsOp {
  type: "set_column_widths";
  sheet: string;
  columnWidths: Record<number | string, number>;
}

/** 0 removes the key. */
export interface SetFrozenOp {
  type: "set_frozen";
  sheet: string;
  rows: number;
  cols: number;
}

/** null removes the meta. */
export interface SetMetaOp {
  type: "set_meta";
  meta: Record<string, unknown> | null;
}
