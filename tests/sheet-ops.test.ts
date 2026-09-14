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

/*
 * PHP holy-sheet 2.3.2's five op fixes and 2.3.3's sixth, ported from the
 * `it(...)` blocks PHP added at the bottom of SheetOpsTest.php.
 */
describe("PHP 2.3.2 and 2.3.3 op fixes", () => {
  function nest(leaf: unknown, levels: number, map: boolean): unknown {
    let value = leaf;
    for (let i = 0; i < levels; i++) value = map ? { k: value } : [value];
    return value;
  }

  it("publishes a set_column_widths schema that accepts the list PHP encodes widths 0..n-1 as", () => {
    // Found by the Python port: 2.3.1 allowed only an EMPTY list, but PHP
    // encodes widths for columns A, B, C as `[120, 80, 140]`, and a schema read
    // from PHP's JSON carries them into this port's diff that way.
    const a = hsWorkbook();
    const b = hsWorkbook();
    b.sheets[0].columnWidths = [120, 80, 140];

    const ops = Agent.diff(a, b) as Any[];
    expect(JSON.stringify(ops[0].columnWidths)).toBe("[120,80,140]");

    const variant = (Agent.opSchema() as Any).oneOf.find((v: Any) => v.properties.type.const === "set_column_widths");
    const widths = variant.properties.columnWidths;

    expect(widths).not.toHaveProperty("maxItems");
    expect(widths).toStrictEqual({
      type: ["object", "array"],
      items: { type: "number", minimum: 0 },
      additionalProperties: { type: "number", minimum: 0 },
    });
  });

  it("ignores an op whose type is not a string, instead of matching a case loosely", () => {
    const w = deepFreeze(hsWorkbook());

    expect(Agent.reduce(w, { type: true, sheet: "Q3" } as Any)).toStrictEqual(w);
    expect(Agent.reduce(w, { type: 0, sheet: "Q3" } as Any)).toStrictEqual(w);
    expect(Agent.reduce(w, { type: ["remove_sheet"], sheet: "Q3" } as Any)).toStrictEqual(w);
  });

  it("trims an address, so a padded one reaches the cell it names", () => {
    const w = deepFreeze(hsWorkbook());

    const set = Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: " b3 ", value: 1 });
    expect(Object.keys(set.sheets[0].cells)).not.toContain(" B3 ");
    expect(set.sheets[0].cells.B3).toStrictEqual({ value: 1 });

    const cleared = Agent.reduce(w, { type: "clear_cell", sheet: "Q3", address: " a2 " });
    expect(cleared.sheets[0].cells).not.toHaveProperty("A2");

    // PHP trim()'s set: space, tab, LF, CR, NUL and vertical tab. Not NBSP.
    const control = Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "\t\n\r\x00\x0Bc9\x0B", value: 2 });
    expect(control.sheets[0].cells.C9).toStrictEqual({ value: 2 });
    expect(Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "\u00a0c9", value: 2 })).toStrictEqual(w);
  });

  it("drops a column-width key that is not a column index, instead of reading it as column A", () => {
    // A JS object lists "abc" after the integer keys, so reading it as column A
    // would overwrite 10 with 999.
    const w = { sheets: [{ name: "S", cells: {}, columnWidths: { 0: 10, abc: 999, 1: 20 } }] };

    const moved = Agent.reduce(w, { type: "insert_columns", sheet: "S", at: 1, count: 1 });
    expect(moved.sheets[0].columnWidths).toStrictEqual({ 1: 10, 2: 20 });

    // `ctype_digit` keeps a digit string PHP stores as a string key ("007"), and
    // PHP stores "-1" as the int -1, which is kept too. "1.5", "" and "-0" are
    // not indexes.
    const odd = { sheets: [{ name: "S", cells: {}, columnWidths: { "007": 70, "1.5": 15, "": 1, "-0": 2, "-1": 5 } }] };
    const shifted = Agent.reduce(odd, { type: "delete_columns", sheet: "S", at: 1, count: 1 });
    expect(shifted.sheets[0].columnWidths).toStrictEqual({ 6: 70, "-1": 5 });
  });

  it("refuses to compare values JSON cannot hold, instead of calling them the same", () => {
    // PHP: invalid UTF-8, NAN and INF. A JS string cannot hold invalid UTF-8;
    // its equivalent is a lone surrogate, which PHP's own json_decode refuses.
    expect(() => SheetDiff.same("\uD800", "\uD801")).toThrow(TypeError);
    expect(() => SheetDiff.same({ ["\uDC00"]: 1 }, { ["\uDC01"]: 1 })).toThrow(TypeError);
    expect(() => SheetDiff.same(NaN, Infinity)).toThrow(TypeError);
    expect(() => SheetDiff.same({ value: -Infinity }, { value: null })).toThrow(TypeError);
    expect(() => SheetDiff.same(1, NaN)).toThrow(TypeError);
    expect(SheetDiff.same("😀", "😀")).toBe(true); // a surrogate PAIR is fine

    // And so does a diff over cells holding them, where it used to record no change.
    const a = { sheets: [{ name: "S", cells: { A1: { value: NaN } } }] };
    const b = { sheets: [{ name: "S", cells: { A1: { value: Infinity } } }] };
    expect(() => Agent.diff(a, b)).toThrow(TypeError);
  });

  it("compares as deep as PHP's json_encode depth of 4096, and no deeper", () => {
    // The boundaries PHP 8.4 printed for SheetDiff::same: every array is a
    // level, an empty one included.
    for (const map of [false, true]) {
      expect(SheetDiff.same(nest(1, 4096, map), nest(1, 4096, map))).toBe(true);
      expect(() => SheetDiff.same(nest(1, 4097, map), 1)).toThrow(TypeError);
      expect(SheetDiff.same(nest([], 4095, map), nest({}, 4095, map))).toBe(true);
      expect(() => SheetDiff.same(nest([], 4096, map), 1)).toThrow(TypeError);
    }
  });

  it("skips an op whose position or count is not a number, instead of reading it as 0", () => {
    const w = deepFreeze(hsWorkbook());

    // `(int) "last"` is 0: these moved Notes to the front, inserted a sheet
    // there, and unfroze the header row.
    expect(Agent.reduce(w, { type: "move_sheet", sheet: "Notes", toIndex: "last" } as Any)).toStrictEqual(w);
    expect(Agent.reduce(w, { type: "add_sheet", index: "end", sheet: { name: "X", cells: {} } } as Any)).toStrictEqual(w);
    expect(Agent.reduce(w, { type: "set_frozen", sheet: "Q3", rows: "one", cols: 0 } as Any)).toStrictEqual(w);
    expect(Agent.reduce(w, { type: "insert_rows", sheet: "Q3", at: 2, count: "2x" } as Any)).toStrictEqual(w);

    // Present means present: a null position skips the op, and so does a junk
    // one on an op that reads no position.
    expect(Agent.reduce(w, { type: "add_sheet", index: null, sheet: { name: "X", cells: {} } } as Any)).toStrictEqual(w);
    expect(Agent.reduce(w, { type: "set_cell", sheet: "Q3", address: "A1", value: 1, count: "x" } as Any)).toStrictEqual(w);
    expect(Agent.reduce(w, { type: "delete_rows", sheet: "Q3", at: 1.5, count: 1 } as Any)).toStrictEqual(w);

    // Ints, digit strings and absent defaults still work.
    expect(Agent.reduce(w, { type: "move_sheet", sheet: "Notes", toIndex: "0" } as Any).sheets[0].name).toBe("Notes");
    expect(Agent.reduce(w, { type: "add_sheet", sheet: { name: "X", cells: {} } } as Any).sheets[2].name).toBe("X");
    expect(Agent.reduce(w, { type: "move_sheet", sheet: "Q3" } as Any)).toStrictEqual(w);
    const frozen = Agent.reduce(w, { type: "set_frozen", sheet: "Q3", cols: "2" } as Any).sheets[0];
    expect([frozen.frozenRows, frozen.frozenCols]).toEqual([undefined, 2]);
  });
});
