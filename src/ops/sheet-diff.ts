import { Agent } from "../agent";
import { type Any, type Obj, entriesOf, get, has, isArr, isEmptyArr, parseAddress, phpString, valuesOf } from "./php";
import { SheetReducer } from "./sheet-reducer";
import type { SheetOp } from "./types";

/**
 * The op list that turns one Holy Sheet schema into another. Mirrors PHP
 * `HolySheet\Ops\SheetDiff`: the same algorithm, so the same inputs give the
 * same ops in the same order in both runtimes (pinned against the PHP package
 * by `tests/sheet-ops-parity.test.ts`).
 *
 * ## Two guarantees, and where each applies
 *
 * 1. **Same workbook, no ops.** When `a` and `b` write the same workbook —
 *    compared as `read(toBytes(...))`, so a columns/rows sheet and the cell
 *    map it becomes are the same, and the timestamp the writer stamps on a
 *    schema that names none is not a change — the diff is `[]`. This is what
 *    makes `diff(s, read(toBytes(s)))` `[]`: saving without a change records
 *    nothing.
 * 2. **Otherwise, exact.** `reduce(a, diff(a, b))` equals `b`, key order
 *    aside. The ops are computed on the schemas as given and VERIFIED by
 *    replaying them through `SheetReducer`; a sheet whose granular ops do not
 *    reproduce it is replaced whole, and a workbook that still does not match
 *    is replaced whole. Correct first, small second.
 *
 * ## Small edits stay small
 *
 * Rows and columns are ALIGNED before cells are compared (a longest common
 * subsequence over each row's, then each column's, contents), so inserting a row
 * above a hundred others is one `insert_rows` plus the new row's cells, not a
 * hundred rewritten cells. One changed cell is one `set_cell`.
 *
 * Granular ops apply to sheets in CELL form — the form `read()` returns. A sheet
 * authored as columns/rows/theme/totals, on either side, is replaced whole when
 * it changes: those keys expand at write time into cells an op could not
 * address without re-deriving the writer.
 *
 * ## Determinism
 *
 * Alignment ties break toward deleting first, and the alignment is skipped (the
 * changed middle treated as rows changed in place) past `ALIGN_LIMIT` cells of
 * work.
 */
export const SheetDiff = {
  /** Above this many LCS cells (after trimming the common prefix and suffix), rows or columns are not aligned. */
  ALIGN_LIMIT: 250_000,

  diff(a: Any, b: Any): SheetOp[] {
    if (SheetDiff.same(a, b) || SheetDiff.equivalent(a, b)) {
      return [];
    }

    const ops = granular(a, b);

    if (SheetDiff.same(SheetReducer.applyAll(a, ops), b)) {
      return ops;
    }

    return [{ type: "set_workbook", data: b }];
  },

  /**
   * Whether two schemas write the same workbook.
   *
   * Both are written and read back. `meta.created` is compared only when both
   * schemas name one: the writer stamps the current time on a schema that
   * does not, and that timestamp is not an edit.
   */
  equivalent(a: Any, b: Any): boolean {
    let describedA = described(a);
    let describedB = described(b);

    if (!namesCreated(a) || !namesCreated(b)) {
      describedA = withoutCreated(describedA);
      describedB = withoutCreated(describedB);
    }

    return SheetDiff.same(describedA, describedB);
  },

  /**
   * Structural equality with map key order ignored and list order kept.
   *
   * Equality is on values, never on insertion order — a JS object cannot keep
   * the order of integer-like keys anyway (`columnWidths`). It follows PHP's
   * canonical form, where lists and maps are one array type: `{}` equals `[]`,
   * and an object keyed exactly "0".."n-1" equals the list of its values.
   *
   * Two divergences from the PHP `same()` cannot be removed, because the
   * distinctions do not exist in JS values:
   *
   * - PHP encodes the float `1.0` as `1.0` and the int `1` as `1`, so it calls
   *   them different; JS has one number, so here they are the same. A PHP diff
   *   can therefore emit a `set_cell` that rewrites `1` as `1.0` where this one
   *   emits nothing. Parsed JSON gives PHP a float only for a literal written
   *   with a fraction or exponent.
   * - PHP treats an int-keyed array whose keys are 0..n-1 but out of order (for
   *   example `columnWidths` decoded from `{"1":140,"0":120}`) as a map, not a
   *   list, and so as different from the same widths in order. A JS object has
   *   no such order to compare.
   */
  same(a: unknown, b: unknown): boolean {
    return canon(a) === canon(b);
  },

  /**
   * Hunks of a longest-common-subsequence alignment: [start in a, deleted, inserted].
   *
   * Ties break toward deleting first. Past `ALIGN_LIMIT` the changed middle is
   * one hunk.
   */
  hunks(a: readonly string[], b: readonly string[]): [number, number, number][] {
    const n = a.length;
    const m = b.length;
    let prefix = 0;

    while (prefix < n && prefix < m && a[prefix] === b[prefix]) {
      prefix++;
    }

    let suffix = 0;

    while (suffix < n - prefix && suffix < m - prefix && a[n - 1 - suffix] === b[m - 1 - suffix]) {
      suffix++;
    }

    const midA = a.slice(prefix, n - suffix);
    const midB = b.slice(prefix, m - suffix);
    const rows = midA.length;
    const cols = midB.length;

    if (rows === 0 && cols === 0) {
      return [];
    }

    if (rows * cols > SheetDiff.ALIGN_LIMIT) {
      return [[prefix, rows, cols]];
    }

    // lengths[i][j] = LCS of midA[i..] and midB[j..]
    const lengths: Int32Array[] = [];
    for (let i = 0; i <= rows; i++) {
      lengths.push(new Int32Array(cols + 1));
    }

    for (let i = rows - 1; i >= 0; i--) {
      for (let j = cols - 1; j >= 0; j--) {
        lengths[i]![j] =
          midA[i] === midB[j] ? lengths[i + 1]![j + 1]! + 1 : Math.max(lengths[i + 1]![j]!, lengths[i]![j + 1]!);
      }
    }

    const hunks: [number, number, number][] = [];
    let open: [number, number, number] | null = null;
    let i = 0;
    let j = 0;

    while (i < rows || j < cols) {
      if (i < rows && j < cols && midA[i] === midB[j]) {
        if (open !== null) {
          hunks.push(open);
          open = null;
        }
        i++;
        j++;
        continue;
      }

      open ??= [prefix + i, 0, 0];

      if (j >= cols || (i < rows && lengths[i + 1]![j]! >= lengths[i]![j + 1]!)) {
        open[1]++;
        i++;
      } else {
        open[2]++;
        j++;
      }
    }

    if (open !== null) {
      hunks.push(open);
    }

    return hunks;
  },
};

/** PHP `isset($schema['meta']['created'])`. */
function namesCreated(schema: unknown): boolean {
  return get(get(schema, "meta"), "created") !== null;
}

function withoutCreated(schema: Obj): Obj {
  const out: Obj = { ...schema };
  if (isArr(out["meta"]) && !Array.isArray(out["meta"])) {
    const meta: Obj = { ...out["meta"] };
    delete meta["created"];
    out["meta"] = meta;
  }
  if (has(out, "meta") && isEmptyArr(out["meta"])) {
    delete out["meta"];
  }
  return out;
}

function described(schema: Any): Obj {
  return Agent.read(Agent.toBytes(schema));
}

function sheetName(sheet: unknown): string {
  return phpString(get(sheet, "name") ?? "");
}

function granular(a: Any, b: Any): SheetOp[] {
  const sheetsA: Any[] = isArr(get(a, "sheets")) ? valuesOf(a.sheets) : [];
  const sheetsB: Any[] = isArr(get(b, "sheets")) ? valuesOf(b.sheets) : [];
  const namesA = sheetsA.map(sheetName);
  const namesB = sheetsB.map(sheetName);

  // Sheets are addressed by name. Two with one name cannot be told apart.
  if (new Set(namesA).size !== namesA.length || new Set(namesB).size !== namesB.length) {
    return [{ type: "set_workbook", data: b }];
  }

  const ops: SheetOp[] = [];

  if (!SheetDiff.same(get(a, "meta"), get(b, "meta"))) {
    ops.push({ type: "set_meta", meta: get(b, "meta") });
  }

  const removed = namesA.filter((name) => !namesB.includes(name));
  const added = namesB.filter((name) => !namesA.includes(name));
  const renames = new Map<string, string>();
  const renamedTo = (name: string): boolean => [...renames.values()].includes(name);

  // A sheet whose contents are unchanged under a new name is a rename.
  for (const old of removed) {
    for (const name of added) {
      if (renamedTo(name)) {
        continue;
      }
      if (
        SheetDiff.same(withoutName(sheetsA[namesA.indexOf(old)]), withoutName(sheetsB[namesB.indexOf(name)]))
      ) {
        renames.set(old, name);
        break;
      }
    }
  }

  // One sheet gone and one arrived is a rename, possibly with edits.
  let unmatchedRemoved = removed.filter((name) => !renames.has(name));
  let unmatchedAdded = added.filter((name) => !renamedTo(name));

  if (unmatchedRemoved.length === 1 && unmatchedAdded.length === 1) {
    renames.set(unmatchedRemoved[0]!, unmatchedAdded[0]!);
    unmatchedRemoved = [];
    unmatchedAdded = [];
  }

  for (const name of unmatchedRemoved) {
    ops.push({ type: "remove_sheet", sheet: name });
  }

  for (const name of namesA) {
    if (renames.has(name)) {
      ops.push({ type: "rename_sheet", sheet: name, name: renames.get(name)! });
    }
  }

  namesB.forEach((name, index) => {
    if (unmatchedAdded.includes(name)) {
      ops.push({ type: "add_sheet", index, sheet: sheetsB[index] });
    }
  });

  // Put the sheets in b's order.
  let state = SheetReducer.applyAll(a, ops);
  namesB.forEach((name, index) => {
    const current = new Map(entriesOf(get(state, "sheets") ?? []).map(([key, sheet]) => [key, sheetName(sheet)]));
    if ((current.get(String(index)) ?? null) !== name) {
      const op: SheetOp = { type: "move_sheet", sheet: name, toIndex: index };
      ops.push(op);
      state = SheetReducer.apply(state, op);
    }
  });

  sheetsB.forEach((target, index) => {
    const current = state.sheets[index];

    if (SheetDiff.same(current, target)) {
      return;
    }

    for (const op of sheetOps(current, target)) {
      ops.push(op);
    }
  });

  return ops;
}

function sheetOps(sheet: Obj, target: Obj): SheetOp[] {
  const name = phpString(target["name"]);
  const replace: SheetOp[] = [{ type: "replace_sheet", sheet: name, data: target }];

  if (!isCellForm(sheet) || !isCellForm(target)) {
    return replace;
  }

  const ops: SheetOp[] = [];
  let work = sheet;

  for (const axis of ["row", "col"] as const) {
    const structural = structuralOps(name, work, target, axis);
    for (const op of structural) {
      ops.push(op);
      work = applyToSheet(work, op);
    }
  }

  const cellsW: Obj = isArr(get(work, "cells")) ? work["cells"] : {};
  const cellsT: Obj = isArr(get(target, "cells")) ? target["cells"] : {};
  const addresses = [...new Set([...entriesOf(cellsW), ...entriesOf(cellsT)].map(([address]) => address))];

  for (const address of rowMajor(addresses)) {
    if (!has(cellsT, address)) {
      ops.push({ type: "clear_cell", sheet: name, address });
      continue;
    }

    const was = get(cellsW, address);

    if (was !== null && SheetDiff.same(was, cellsT[address])) {
      continue;
    }

    ops.push(setCellOp(name, address, isArr(was) ? was : null, isArr(cellsT[address]) ? cellsT[address] : {}));
  }

  if (!SheetDiff.same(get(work, "mergedRegions") ?? [], get(target, "mergedRegions") ?? [])) {
    ops.push({ type: "set_merged_regions", sheet: name, mergedRegions: get(target, "mergedRegions") ?? [] });
  }

  if (!SheetDiff.same(get(work, "columnWidths") ?? [], get(target, "columnWidths") ?? [])) {
    // PHP emits its empty array here, which JSON-encodes as `[]`; `{}` is the
    // same value (both reducers read either as empty) in the map shape the op
    // schema declares.
    ops.push({ type: "set_column_widths", sheet: name, columnWidths: get(target, "columnWidths") ?? {} });
  }

  if (
    !identical(get(work, "frozenRows") ?? 0, get(target, "frozenRows") ?? 0) ||
    !identical(get(work, "frozenCols") ?? 0, get(target, "frozenCols") ?? 0)
  ) {
    ops.push({
      type: "set_frozen",
      sheet: name,
      rows: get(target, "frozenRows") ?? 0,
      cols: get(target, "frozenCols") ?? 0,
    });
  }

  // Verified per sheet, so one sheet the granular ops cannot reproduce is
  // replaced without costing the rest of the workbook its small diff.
  let result = sheet;
  for (const op of ops) {
    result = applyToSheet(result, op);
  }

  return SheetDiff.same(result, target) ? ops : replace;
}

/** PHP `===`: primitives by value and type; arrays structurally. */
function identical(x: unknown, y: unknown): boolean {
  return isArr(x) || isArr(y) ? SheetDiff.same(x, y) : x === y;
}

function setCellOp(sheet: string, address: string, was: Obj | null, cell: Obj): SheetOp {
  const op: Obj = { type: "set_cell", sheet, address };

  if (has(cell, "value")) {
    op["value"] = cell["value"];
  }

  for (const key of ["formula", "computedValue"]) {
    if (has(cell, key)) {
      op[key] = cell[key];
    }
  }

  for (const key of ["format", "comment"]) {
    if (!SheetDiff.same(get(was, key), get(cell, key))) {
      op[key] = get(cell, key);
    }
  }

  return op as SheetOp;
}

/**
 * Row (or column) inserts and deletes, found by aligning contents.
 *
 * Emitted from the bottom (or right) up, so each op's position is still the
 * position in the sheet as it was.
 */
function structuralOps(name: string, sheet: Obj, target: Obj, axis: "row" | "col"): SheetOp[] {
  const sheetLines = lines(sheet, axis);
  const targetLines = lines(target, axis);
  const ops: SheetOp[] = [];

  for (const [start, deleted, inserted] of SheetDiff.hunks(sheetLines, targetLines).reverse()) {
    const kept = Math.min(deleted, inserted);
    const at = start + kept + 1;

    if (deleted > inserted) {
      ops.push({ type: axis === "row" ? "delete_rows" : "delete_columns", sheet: name, at, count: deleted - inserted });
    } else if (inserted > deleted && at <= sheetLines.length) {
      // Past the last row there is nothing to move down.
      ops.push({ type: axis === "row" ? "insert_rows" : "insert_columns", sheet: name, at, count: inserted - deleted });
    }
  }

  return ops;
}

/** Each row's (or column's) contents, as a comparable string, index 0 = row 1. */
function lines(sheet: Obj, axis: "row" | "col"): string[] {
  const grouped = new Map<number, Map<number, Any>>();
  let max = 0;

  for (const [address, cell] of entriesOf(isArr(get(sheet, "cells")) ? sheet["cells"] : {})) {
    const parsed = parseAddress(address);
    if (parsed === null) {
      continue;
    }
    const [col, row] = parsed;
    const line = axis === "row" ? row : col + 1;
    const position = axis === "row" ? col : row;
    let group = grouped.get(line);
    if (group === undefined) {
      group = new Map();
      grouped.set(line, group);
    }
    group.set(position, cell);
    max = Math.max(max, line);
  }

  const out: string[] = [];
  for (let i = 1; i <= max; i++) {
    const positions = [...(grouped.get(i) ?? new Map<number, Any>()).entries()].sort((x, y) => x[0] - y[0]);
    out.push(canonEntries(positions.map(([position, cell]) => [String(position), cell])));
  }

  return out;
}

function applyToSheet(sheet: Obj, op: SheetOp): Obj {
  return SheetReducer.apply({ sheets: [sheet] }, op).sheets[0];
}

function isCellForm(sheet: Obj): boolean {
  return entriesOf(sheet).every(([key]) => SheetReducer.CELL_FORM_KEYS.includes(key));
}

function withoutName(sheet: Obj): Obj {
  const out: Obj = isArr(sheet) ? { ...sheet } : {};
  delete out["name"];
  return out;
}

function rowMajor(addresses: readonly string[]): string[] {
  const parsed: [number, number, string][] = [];

  for (const address of addresses) {
    const cell = parseAddress(address);
    if (cell !== null) {
      parsed.push([cell[1], cell[0], address]);
    }
  }

  parsed.sort((x, y) => x[0] - y[0] || x[1] - y[1]);

  return parsed.map(([, , address]) => address);
}

/**
 * PHP's canonical JSON, for equality only: map keys sorted, list order kept, and
 * an array keyed exactly 0..n-1 (every empty one included) written as a list.
 * The key order and escaping differ from PHP's bytes; which values compare
 * equal does not (see `same()` for the two exceptions).
 */
function canon(value: unknown): string {
  if (Array.isArray(value)) {
    return "[" + value.map(canon).join(",") + "]";
  }
  if (isArr(value)) {
    return canonEntries(entriesOf(value));
  }
  return JSON.stringify(value) ?? "null";
}

function canonEntries(entries: [string, unknown][]): string {
  if (entries.every(([key], i) => key === String(i))) {
    return "[" + entries.map(([, item]) => canon(item)).join(",") + "]";
  }
  const sorted = entries.slice().sort((x, y) => (x[0] < y[0] ? -1 : x[0] > y[0] ? 1 : 0));
  return "{" + sorted.map(([key, item]) => JSON.stringify(key) + ":" + canon(item)).join(",") + "}";
}
