import {
  type Any,
  type Obj,
  asciiUpper,
  entriesOf,
  get,
  has,
  isArr,
  isEmptyArr,
  isIndexKey,
  letter,
  parseAddress,
  phpInt,
  phpInteger,
  phpString,
  phpTrim,
  valuesOf,
} from "./php";
import { SheetOpSchema } from "./sheet-op-schema";

/**
 * Apply `SheetOpSchema` ops to a Holy Sheet schema, returning a new schema.
 * Mirrors PHP `HolySheet\Ops\SheetReducer`, op for op.
 *
 * Pure: the input is never modified. An op naming a sheet or cell that is not
 * there is skipped, as fancy-sheets' `reduceWorkbook` skips an unknown sheet, so
 * a replayed history degrades rather than throws. Unchanged parts of the input
 * are shared with the result rather than copied, as in any immutable reducer:
 * nothing here mutates either, but a caller that mutates a result in place
 * reaches into the input too.
 *
 * `set_cell`, `set_range` and `set_workbook` keep fancy-sheets' shapes and
 * semantics — a `set_cell` without `formula` clears the formula and keeps the
 * format and comment; a null write to an absent cell does nothing — so the same
 * ops can drive a live `useSheetSync` session. Everything else is holy-sheet's:
 * the structure fancy-sheets has no op for (sheets, rows, columns, merges, widths,
 * panes, meta) and the cell parts it does not carry (`computedValue`, `format`,
 * `comment`).
 *
 * The row and column ops move cells, merged regions and column widths. They do
 * not rewrite formula text: a formula that should follow an inserted row is a
 * change to that cell, and a diff records it as one.
 */
export const SheetReducer = {
  /** Sheet keys the granular ops address. A sheet with any other key is authored form. */
  CELL_FORM_KEYS: ["name", "cells", "mergedRegions", "columnWidths", "frozenRows", "frozenCols"] as readonly string[],

  applyAll(schema: Any, ops: readonly Any[]): Any {
    for (const op of ops) {
      schema = SheetReducer.apply(schema, op);
    }
    return schema;
  },

  apply(schema: Any, op: Any): Any {
    const type = get(op, "type");

    // A string naming an op type, compared strictly (PHP 2.3.2, where a loose
    // `switch` let `type: true` remove a sheet). This port's switch was strict.
    if (typeof type !== "string" || !(SheetOpSchema.TYPES as readonly string[]).includes(type)) {
      return schema;
    }

    if (type === "set_workbook") {
      return isArr(get(op, "data")) ? op.data : schema;
    }

    if (type === "set_meta") {
      const out: Obj = { ...schema };
      if (get(op, "meta") === null) {
        delete out["meta"];
      } else {
        out["meta"] = op.meta;
      }
      return out;
    }

    const sheets: Any[] = isArr(get(schema, "sheets")) ? valuesOf(schema.sheets) : [];

    if (type === "add_sheet") {
      if (!isArr(get(op, "sheet"))) {
        return schema;
      }
      const index = has(op, "index") ? phpInteger(op.index) : sheets.length;
      if (index === null) {
        return schema;
      }
      sheets.splice(Math.max(0, Math.min(sheets.length, index)), 0, op.sheet);
      return { ...schema, sheets };
    }

    const at = find(sheets, phpString(get(op, "sheet") ?? ""));

    if (at === null) {
      return schema;
    }

    // Positions and counts are ints or digit strings (PHP 2.3.3). A present
    // value that is neither skips the op: `(int)` read junk as 0, which moved a
    // sheet to the front, inserted one there, or unfroze panes.
    for (const field of POSITION_FIELDS) {
      if (has(op, field) && phpInteger(op[field]) === null) {
        return schema;
      }
    }

    switch (type) {
      case "remove_sheet":
        sheets.splice(at, 1);
        break;

      case "rename_sheet":
        sheets[at] = { ...sheets[at], name: phpString(get(op, "name") ?? sheets[at].name) };
        break;

      case "move_sheet": {
        const [moved] = sheets.splice(at, 1);
        const to = Math.max(0, Math.min(sheets.length, phpInteger(get(op, "toIndex") ?? at) ?? at));
        sheets.splice(to, 0, moved);
        break;
      }

      case "replace_sheet":
        if (isArr(get(op, "data"))) {
          sheets[at] = op.data;
        }
        break;

      case "set_merged_regions":
        sheets[at] = setOrUnset(sheets[at], "mergedRegions", get(op, "mergedRegions") ?? [], EMPTY);
        break;

      case "set_column_widths":
        sheets[at] = setOrUnset(sheets[at], "columnWidths", get(op, "columnWidths") ?? [], EMPTY);
        break;

      case "set_frozen":
        sheets[at] = setOrUnset(sheets[at], "frozenRows", phpInteger(get(op, "rows") ?? 0) ?? 0, 0);
        sheets[at] = setOrUnset(sheets[at], "frozenCols", phpInteger(get(op, "cols") ?? 0) ?? 0, 0);
        break;

      case "set_cell":
        sheets[at] = setCell(sheets[at], op);
        break;

      case "set_range":
        sheets[at] = setRange(sheets[at], op);
        break;

      case "clear_cell": {
        const address = asciiUpper(phpTrim(phpString(get(op, "address") ?? "")));
        const cells = sheets[at].cells;
        // PHP's `unset($sheet['cells'][$address])` creates nothing when the key is absent.
        if (isArr(cells) && !Array.isArray(cells) && has(cells, address)) {
          const copy: Obj = { ...cells };
          delete copy[address];
          sheets[at] = { ...sheets[at], cells: copy };
        }
        break;
      }

      case "insert_rows":
        sheets[at] = shiftRows(sheets[at], position(op, "at"), Math.max(0, position(op, "count")));
        break;

      case "delete_rows":
        sheets[at] = shiftRows(sheets[at], position(op, "at"), -Math.max(0, position(op, "count")));
        break;

      case "insert_columns":
        sheets[at] = shiftColumns(sheets[at], position(op, "at"), Math.max(0, position(op, "count")));
        break;

      case "delete_columns":
        sheets[at] = shiftColumns(sheets[at], position(op, "at"), -Math.max(0, position(op, "count")));
        break;

      default:
        return schema;
    }

    return { ...schema, sheets };
  },
};

/** Op fields that hold a position or a count. */
const POSITION_FIELDS = ["index", "toIndex", "rows", "cols", "at", "count"] as const;

/** PHP `self::integer($op[$key] ?? 0) ?? 0`, for a field the guard in apply() has checked. */
function position(op: Obj, key: string): number {
  return phpInteger(get(op, key) ?? 0) ?? 0;
}

/** The `[]` a map or list key is removed at. */
const EMPTY = Symbol("empty");

function find(sheets: Any[], name: string): number | null {
  for (let i = 0; i < sheets.length; i++) {
    if (get(sheets[i], "name") === name) {
      return i;
    }
  }
  return null;
}

function setOrUnset(sheet: Obj, key: string, value: Any, empty: typeof EMPTY | 0): Obj {
  const out: Obj = { ...sheet };
  if (value === null || (empty === EMPTY ? isEmptyArr(value) : value === empty)) {
    delete out[key];
  } else {
    out[key] = value;
  }
  return out;
}

/**
 * fancy-sheets' `set_cell`, plus the parts of a cell it does not carry.
 *
 * - `value` is written; `formula` and `computedValue` are REPLACED — absent
 *   in the op means absent in the cell, as fancy-sheets clears `formula`.
 * - `format` and `comment` are KEPT when the op omits them, replaced when it
 *   carries one, and removed when it carries null.
 * - A null write to an absent cell, carrying nothing else, does nothing.
 */
function setCell(sheet: Obj, op: Obj): Obj {
  // Trimmed as it is validated (PHP 2.3.2): `" a1 "` is A1, not a key of its own.
  const address = asciiUpper(phpTrim(phpString(get(op, "address") ?? "")));

  if (parseAddress(address) === null) {
    return sheet;
  }

  const cells: Obj = isArr(get(sheet, "cells")) ? { ...sheet.cells } : {};
  const existing = get(cells, address);
  const value = get(op, "value");

  const carries = (key: string): boolean => has(op, key) && op[key] !== null;

  if (
    existing === null &&
    value === null &&
    !carries("formula") &&
    !carries("computedValue") &&
    !carries("format") &&
    !carries("comment")
  ) {
    return sheet;
  }

  // `value` is optional in a CellData ({"formula": "SUM(A1:A3)"} is a whole
  // cell), so an op without one writes a cell without one.
  const cell: Obj = has(op, "value") ? { value } : {};

  for (const key of ["formula", "computedValue"]) {
    if (carries(key)) {
      cell[key] = op[key];
    }
  }

  for (const key of ["format", "comment"]) {
    if (has(op, key)) {
      if (op[key] !== null) {
        cell[key] = op[key];
      }
    } else if (isArr(existing) && has(existing, key)) {
      cell[key] = existing[key];
    }
  }

  if (Object.keys(cell).length === 0) {
    delete cells[address];
  } else {
    cells[address] = cell;
  }

  return { ...sheet, cells };
}

/**
 * fancy-sheets' `set_range`: values row-major from `start`, each written as a
 * `set_cell` with no formula. `end` is accepted and, as there, not read.
 */
function setRange(sheet: Obj, op: Obj): Obj {
  const start = parseAddress(phpString(get(op, "start") ?? ""));

  if (start === null || !isArr(get(op, "values"))) {
    return sheet;
  }

  const [col0, row0] = start;
  const rows = valuesOf(op.values);

  for (let r = 0; r < rows.length; r++) {
    const row = valuesOf(isArr(rows[r]) ? rows[r] : []);
    for (let c = 0; c < row.length; c++) {
      sheet = setCell(sheet, { address: letter(col0 + c) + String(row0 + r), value: row[c] === undefined ? null : row[c] });
    }
  }

  return sheet;
}

/** Insert (`delta` > 0) or delete (`delta` < 0) rows at 1-based row `at`. */
function shiftRows(sheet: Obj, at: number, delta: number): Obj {
  if (at < 1 || delta === 0) {
    return sheet;
  }

  sheet = remapCells(sheet, (col, row) => {
    if (row < at) {
      return [col, row];
    }
    if (delta < 0 && row < at - delta) {
      return null;
    }
    return [col, row + delta];
  });

  return remapMerges(sheet, "row", at, delta);
}

/** Insert or delete columns at 1-based column `at` (A = 1). */
function shiftColumns(sheet: Obj, at: number, delta: number): Obj {
  if (at < 1 || delta === 0) {
    return sheet;
  }

  sheet = remapCells(sheet, (col, row) => {
    const number = col + 1;
    if (number < at) {
      return [col, row];
    }
    if (delta < 0 && number < at - delta) {
      return null;
    }
    return [col + delta, row];
  });

  if (isArr(get(sheet, "columnWidths"))) {
    const widths = new Map<number, Any>();
    for (const [index, width] of entriesOf(sheet.columnWidths)) {
      // A key that is not a column index is dropped (PHP 2.3.2), not read as
      // column 0: `(int) "abc"` overwrote column A's width.
      if (!isIndexKey(index)) {
        continue;
      }
      const i = phpInt(index);
      const number = i + 1;
      if (number < at) {
        widths.set(i, width);
      } else if (delta > 0 || number >= at - delta) {
        widths.set(i + delta, width);
      }
    }
    // ksort
    const sorted: Obj = {};
    for (const [index, width] of [...widths.entries()].sort((x, y) => x[0] - y[0])) {
      sorted[index] = width;
    }
    sheet = setOrUnset(sheet, "columnWidths", sorted, EMPTY);
  }

  return remapMerges(sheet, "col", at, delta);
}

/**
 * Move a 1-based span [start, end] for an insert or delete at `at`. Null when a
 * delete removes the whole span; a span the delete cuts into shrinks.
 */
function shiftSpan(start: number, end: number, at: number, delta: number): [number, number] | null {
  if (delta > 0) {
    return [start >= at ? start + delta : start, end >= at ? end + delta : end];
  }

  const last = at - delta - 1; // the last deleted index
  const newStart = start < at ? start : start > last ? start + delta : at;
  const newEnd = end < at ? end : end > last ? end + delta : at - 1;

  return newEnd < newStart ? null : [newStart, newEnd];
}

function remapCells(sheet: Obj, move: (col: number, row: number) => [number, number] | null): Obj {
  if (!isArr(get(sheet, "cells"))) {
    return sheet;
  }

  const placed: [number, number, string, Any][] = [];

  for (const [address, cell] of entriesOf(sheet.cells)) {
    const parsed = parseAddress(address);

    if (parsed === null) {
      continue;
    }

    const to = move(parsed[0], parsed[1]);

    if (to !== null) {
      placed.push([to[1], to[0], letter(to[0]) + String(to[1]), cell]);
    }
  }

  // Row-major, the order describe() reads cells in. Array#sort is stable, as PHP 8's usort is.
  placed.sort((x, y) => x[0] - y[0] || x[1] - y[1]);

  const cells: Obj = {};
  for (const [, , address, cell] of placed) {
    cells[address] = cell;
  }

  return { ...sheet, cells };
}

function remapMerges(sheet: Obj, axis: "row" | "col", at: number, delta: number): Obj {
  if (!isArr(get(sheet, "mergedRegions"))) {
    return sheet;
  }

  const regions: Any[] = [];

  for (const region of valuesOf(sheet.mergedRegions)) {
    const start = parseAddress(phpString(get(region, "start") ?? ""));
    const end = parseAddress(phpString(get(region, "end") ?? ""));

    if (start === null || end === null) {
      regions.push(region);
      continue;
    }

    if (axis === "row") {
      const moved = shiftSpan(start[1], end[1], at, delta);
      if (moved === null) {
        continue;
      }
      regions.push({ start: letter(start[0]) + String(moved[0]), end: letter(end[0]) + String(moved[1]) });
    } else {
      const moved = shiftSpan(start[0] + 1, end[0] + 1, at, delta);
      if (moved === null) {
        continue;
      }
      regions.push({ start: letter(moved[0] - 1) + String(start[1]), end: letter(moved[1] - 1) + String(end[1]) });
    }
  }

  return setOrUnset(sheet, "mergedRegions", regions, EMPTY);
}
