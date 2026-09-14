import { describe, it, expect } from "vitest";
import { Agent, SheetDiff, SheetOpSchema } from "../src";
import { type Any, EDITS, hsWorkbook, randomEdits } from "./sheet-ops-fixtures";

/*
 * Agent.diff / Agent.reduce / Agent.opSchema / Agent.equivalent (holy-sheet #7).
 * Ported from PHP tests/Unit/SheetOpsTest.php, case for case.
 *
 * What a version history built on these needs, pinned:
 *
 * 1. Round trip: reduce(a, diff(a, b)) equals b.
 * 2. Small edits stay small: one cell is one set_cell, an inserted row is one
 *    insert_rows plus its cells — asserted as the exact op TYPES, not just
 *    "fewer ops than a replace".
 * 3. A save without a change records nothing: diff(s, read(toBytes(s))) is [].
 * 4. set_cell / set_range / set_workbook behave as fancy-sheets' reducer does.
 */

function deepFreeze<T>(value: T): T {
  if (typeof value === "object" && value !== null && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const key of Object.keys(value)) deepFreeze((value as Any)[key]);
  }
  return value;
}

describe("Agent.diff reproduces the target exactly, with the smallest ops", () => {
  for (const [name, [edit, types]] of Object.entries(EDITS)) {
    it(name, () => {
      const a = hsWorkbook();
      const b = edit(hsWorkbook());

      const ops = Agent.diff(a, b);

      expect(ops.map((op) => op.type)).toEqual(types);
      expect(ops.every((op) => (SheetOpSchema.TYPES as readonly string[]).includes(op.type))).toBe(true);
      expect(SheetDiff.same(Agent.reduce(a, ops), b)).toBe(true);

      // And back: the reverse diff is what a version history stores.
      const reverse = Agent.diff(b, a);
      expect(SheetDiff.same(Agent.reduce(b, reverse), a)).toBe(true);
    });
  }
});

describe("Agent.diff", () => {
  it("records nothing for a save without a change", () => {
    const described = hsWorkbook();
    const readBack = Agent.read(Agent.toBytes(described));

    expect(Agent.diff(described, readBack)).toEqual([]);

    // An AUTHORED schema, with no meta: the writer expands columns/rows into
    // cells and stamps a creation time, and neither is an edit.
    const authored = {
      sheets: [
        {
          name: "Deals",
          columns: [{ header: "Name" }, { header: "Value", type: "currency" }],
          rows: [
            ["Acme", 1200],
            ["Globex", 800],
          ],
          totals: { Value: "sum" },
        },
      ],
    };
    const authoredBack = Agent.read(Agent.toBytes(authored));

    expect(Agent.diff(authored, authoredBack)).toEqual([]);
    expect(Agent.equivalent(authored, authoredBack)).toBe(true);
  });

  it("falls back to replacing the workbook when sheets cannot be told apart", () => {
    const a = {
      sheets: [
        { name: "Same", cells: { A1: { value: 1 } } },
        { name: "Same", cells: { A1: { value: 2 } } },
      ],
    };
    const b = { sheets: [{ name: "Same", cells: { A1: { value: 3 } } }] };

    const ops = Agent.diff(a, b);

    expect(ops).toStrictEqual([{ type: "set_workbook", data: b }]);
    expect(Agent.reduce(a, ops)).toStrictEqual(b);
  });

  it("keeps the round trip over a seeded run of random cell-form edits, without falling back", () => {
    const pairs = randomEdits(20260915, 60);
    expect(pairs).toHaveLength(60);

    pairs.forEach(([a, b], run) => {
      const ops = Agent.diff(a, b);

      expect(SheetDiff.same(Agent.reduce(a, ops), b), `run ${run}`).toBe(true);
      expect(
        ops.map((op) => op.type).filter((type) => type === "set_workbook" || type === "replace_sheet"),
        `run ${run} fell back`,
      ).toEqual([]);
    });
  });
});

describe("Agent.reduce", () => {
  it("set_cell, set_range and set_workbook behave as fancy-sheets reduceWorkbook does", () => {
    // Frozen, so any write to the input throws rather than passing unnoticed.
    const w = deepFreeze(hsWorkbook());

    // No formula in the op clears the formula; the format stays.
    const cleared = Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "B5", value: 42 });
    expect(cleared.sheets[0].cells.B5).toStrictEqual({ value: 42 });

    const kept = Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "A1", value: "Area" });
    expect(kept.sheets[0].cells.A1).toStrictEqual({ value: "Area", format: { bold: true } });

    // A null write to an absent cell does nothing.
    expect(Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "Z99", value: null })).toStrictEqual(w);

    // set_range writes values row-major from start.
    const ranged = Agent.reduce(w, {
      type: "set_range",
      sheet: "Q3",
      start: "C2",
      values: [
        [1, 2],
        [3, 4],
      ],
    });
    expect(ranged.sheets[0].cells.D3).toStrictEqual({ value: 4 });

    // An unknown sheet is skipped, and the input is never modified.
    expect(Agent.reduce(w, { type: "set_cell", sheet: "Nope", address: "A1", value: 1 })).toStrictEqual(w);
    expect(w).toStrictEqual(hsWorkbook());

    expect(Agent.reduce(w, { type: "set_workbook", data: { sheets: [] } })).toStrictEqual({ sheets: [] });
  });

  it("follows fancy-sheets' set_cell for the parts of a cell it does not carry", () => {
    const w = deepFreeze(hsWorkbook());
    const cell = (schema: Any, sheet: number, address: string) => schema.sheets[sheet].cells[address];

    // null clears a format or comment; omitting it keeps it.
    expect(cell(Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "A1", value: "Area", format: null }), 0, "A1")).toStrictEqual({ value: "Area" });
    expect(cell(Agent.reduce(w, { type: "set_cell", sheet: "Notes", address: "A1", value: "Done" }), 1, "A1")).toStrictEqual({
      value: "Done",
      comment: { text: "Signed off", author: "CFO" },
    });
    expect(cell(Agent.reduce(w, { type: "set_cell", sheet: "Notes", address: "A1", value: "Done", comment: null }), 1, "A1")).toStrictEqual({ value: "Done" });

    // An op without value writes a cell without one.
    expect(cell(Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "C1", formula: "SUM(B2:B4)" }), 0, "C1")).toStrictEqual({ formula: "SUM(B2:B4)" });

    // A null write to an absent cell carrying a format is not a no-op.
    expect(cell(Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "C9", value: null, format: { bold: true } }), 0, "C9")).toStrictEqual({
      value: null,
      format: { bold: true },
    });

    // A cell the op leaves empty is removed, not kept as {}.
    const emptied = Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "A1", format: null });
    expect(Object.keys(emptied.sheets[0].cells)).not.toContain("A1");

    // The address is upper-cased, and a malformed one is skipped.
    expect(cell(Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "c3", value: 7 }), 0, "C3")).toStrictEqual({ value: 7 });
    expect(Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "3C", value: 7 })).toStrictEqual(w);
  });

  it("moves cells, merges and widths on row and column inserts and deletes", () => {
    const w = deepFreeze({
      sheets: [
        {
          name: "S",
          cells: { A1: { value: 1 }, A2: { value: 2 }, C3: { value: 3 } },
          mergedRegions: [{ start: "A2", end: "C4" }],
          columnWidths: { 0: 10, 2: 30 },
        },
      ],
    });

    const rows = Agent.reduce(w, { type: "insert_rows", sheet: "S", at: 2, count: 2 });
    expect(Object.keys(rows.sheets[0].cells)).toEqual(["A1", "A4", "C5"]);
    expect(rows.sheets[0].mergedRegions).toStrictEqual([{ start: "A4", end: "C6" }]);

    const deleted = Agent.reduce(w, { type: "delete_rows", sheet: "S", at: 2, count: 1 });
    expect(Object.keys(deleted.sheets[0].cells)).toEqual(["A1", "C2"]);
    expect(deleted.sheets[0].mergedRegions).toStrictEqual([{ start: "A2", end: "C3" }]);

    const cols = Agent.reduce(w, { type: "delete_columns", sheet: "S", at: 1, count: 1 });
    expect(Object.keys(cols.sheets[0].cells)).toEqual(["B3"]);
    expect(cols.sheets[0].columnWidths).toStrictEqual({ 1: 30 });
    expect(cols.sheets[0].mergedRegions).toStrictEqual([{ start: "A2", end: "B4" }]);
  });

  it("drops a merge a delete removes whole, and the key with it, and sorts cells row-major", () => {
    const w = deepFreeze({
      sheets: [
        {
          name: "S",
          cells: { C1: { value: "c" }, A2: { value: "a" }, B1: { value: "b" } },
          mergedRegions: [{ start: "B1", end: "B2" }],
          columnWidths: { 1: 50 },
        },
      ],
    });

    const gone = Agent.reduce(w, { type: "delete_columns", sheet: "S", at: 2, count: 1 });
    expect(gone.sheets[0]).toStrictEqual({ name: "S", cells: { B1: { value: "c" }, A2: { value: "a" } } });
    expect(Object.keys(gone.sheets[0].cells)).toEqual(["B1", "A2"]);

    const inserted = Agent.reduce(w, { type: "insert_columns", sheet: "S", at: 1, count: 1 });
    expect(Object.keys(inserted.sheets[0].cells)).toEqual(["C1", "D1", "B2"]);
    expect(inserted.sheets[0].columnWidths).toStrictEqual({ 2: 50 });
    expect(inserted.sheets[0].mergedRegions).toStrictEqual([{ start: "C1", end: "C2" }]);
  });

  it("applies the sheet, pane and meta ops, removing a key at its empty value", () => {
    const w = deepFreeze(hsWorkbook());
    const names = (schema: Any) => schema.sheets.map((s: Any) => s.name);

    expect(names(Agent.reduce(w, { type: "add_sheet", index: 0, sheet: { name: "New", cells: {} } }))).toEqual(["New", "Q3", "Notes"]);
    expect(names(Agent.reduce(w, { type: "add_sheet", index: 99, sheet: { name: "New", cells: {} } }))).toEqual(["Q3", "Notes", "New"]);
    expect(names(Agent.reduce(w, { type: "remove_sheet", sheet: "Q3" }))).toEqual(["Notes"]);
    expect(names(Agent.reduce(w, { type: "rename_sheet", sheet: "Q3", name: "Q4" }))).toEqual(["Q4", "Notes"]);
    expect(names(Agent.reduce(w, { type: "move_sheet", sheet: "Q3", toIndex: 1 }))).toEqual(["Notes", "Q3"]);
    expect(Agent.reduce(w, { type: "replace_sheet", sheet: "Notes", data: { name: "Notes" } }).sheets[1]).toStrictEqual({ name: "Notes" });

    const bare = Agent.reduce(w, [
      { type: "set_merged_regions", sheet: "Q3", mergedRegions: [] },
      { type: "set_column_widths", sheet: "Q3", columnWidths: {} },
      { type: "set_frozen", sheet: "Q3", rows: 0, cols: 2 },
      { type: "set_meta", meta: null },
    ]);
    expect(Object.keys(bare.sheets[0])).toEqual(["name", "cells", "frozenCols"]);
    expect(bare.sheets[0].frozenCols).toBe(2);
    expect(bare).not.toHaveProperty("meta");

    // An unknown sheet or type is skipped.
    expect(Agent.reduce(w, { type: "remove_sheet", sheet: "Nope" })).toStrictEqual(w);
    expect(Agent.reduce(w, { type: "nope", sheet: "Q3" } as Any)).toStrictEqual(w);
  });
});

describe("Agent.opSchema", () => {
  it("publishes one schema variant per op type, and diff only emits those types", () => {
    const schema = Agent.opSchema() as Any;

    expect(schema).toStrictEqual(SheetOpSchema.jsonSchema());
    expect(schema.oneOf.map((v: Any) => v.properties.type.const)).toEqual([...SheetOpSchema.TYPES]);
  });
});

describe("SheetDiff", () => {
  it("aligns rows by content, breaking ties toward deleting first", () => {
    expect(SheetDiff.hunks(["a", "b", "c"], ["a", "x", "b", "c"])).toEqual([[1, 0, 1]]);
    expect(SheetDiff.hunks(["a", "b", "c"], ["a", "c"])).toEqual([[1, 1, 0]]);
    expect(SheetDiff.hunks(["a", "b"], ["a", "z"])).toEqual([[1, 1, 1]]);
    expect(SheetDiff.hunks(["a"], ["a"])).toEqual([]);

    // A tie that changes WHICH line is kept: deleting first keeps the later one.
    expect(SheetDiff.hunks(["x", "y"], ["y", "x"])).toEqual([
      [0, 1, 0],
      [2, 0, 1],
    ]);
    expect(SheetDiff.hunks(["a", "x", "y", "b"], ["a", "y", "x", "b"])).toEqual([
      [1, 1, 0],
      [3, 0, 1],
    ]);
  });

  it("diffs swapped rows as PHP does, through the delete-first alignment", () => {
    const a = { sheets: [{ name: "S", cells: { A1: { value: "x" }, A2: { value: "y" }, A3: { value: "z" } } }] };
    const b = { sheets: [{ name: "S", cells: { A1: { value: "y" }, A2: { value: "x" }, A3: { value: "z" } } }] };

    const ops = Agent.diff(a, b);

    expect(ops).toStrictEqual([
      { type: "insert_rows", sheet: "S", at: 3, count: 1 },
      { type: "delete_rows", sheet: "S", at: 1, count: 1 },
      { type: "set_cell", sheet: "S", address: "A2", value: "x" },
    ]);
    expect(SheetDiff.same(Agent.reduce(a, ops), b)).toBe(true);
  });

  it("compares values, not key order, and keeps list order", () => {
    expect(SheetDiff.same({ a: 1, b: [1, 2] }, { b: [1, 2], a: 1 })).toBe(true);
    expect(SheetDiff.same({ b: [1, 2] }, { b: [2, 1] })).toBe(false);
    // PHP's one array type: {} is [], and a map keyed 0..n-1 is that list.
    expect(SheetDiff.same({}, [])).toBe(true);
    expect(SheetDiff.same({ 0: 120, 1: 140 }, [120, 140])).toBe(true);
    expect(SheetDiff.same({ 1: 140, 0: 120 }, { 0: 120, 1: 140 })).toBe(true);
    expect(SheetDiff.same({ 0: 120, 2: 140 }, [120, 140])).toBe(false);
    expect(SheetDiff.same({ value: null }, {})).toBe(false);
    expect(SheetDiff.same("1", 1)).toBe(false);
  });
});
