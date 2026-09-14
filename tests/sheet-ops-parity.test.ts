import { describe, it, expect, beforeAll } from "vitest";
import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Agent, SheetDiff, SheetOpSchema, type SheetOp } from "../src";
import { type Any, EDITS, hsWorkbook, randomEdits, seeded } from "./sheet-ops-fixtures";

// Cross-engine OPS parity: for the same two schemas, the PHP holy-sheet's
// Agent::diff (scripts/php-diff.php) and this port's Agent.diff must return the
// SAME ops in the SAME order, and replaying them must give the same schema. A
// version history written by one runtime is replayed by the other, so "both
// round-trip" is not enough: the stored op lists have to be interchangeable.
//
// PHP has one array type for lists and maps, so its JSON writes an empty map as
// `[]` and a map keyed 0..n-1 (columnWidths) as a list. normalize() folds both
// sides into that model before comparing; nothing else is relaxed.
//
// Skips when `php` isn't on PATH LOCALLY; in CI a missing php THROWS instead,
// because a skip there is a green build with zero cross-engine coverage.

const PHP_SCRIPT = join(__dirname, "..", "scripts", "php-diff.php");

function php(args: string[], opts: Parameters<typeof execFileSync>[2] = {}): Buffer {
  return execFileSync("php", args, { shell: true, maxBuffer: 256 * 1024 * 1024, ...opts }) as Buffer;
}

function phpAvailable(): boolean {
  try {
    php(["--version"], { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

function normalize(value: Any): Any {
  if (Array.isArray(value)) return value.map(normalize);
  if (typeof value === "object" && value !== null) {
    const keys = Object.keys(value);
    if (keys.every((key, i) => key === String(i))) return keys.map((key) => normalize(value[key]));
    const out: Record<string, unknown> = {};
    for (const key of keys.sort()) out[key] = normalize(value[key]);
    return out;
  }
  return value;
}

type Nested = { leaf: Any; depth: number; map: boolean };
type Case = { name: string; a?: Any; b?: Any; ops?: Any[]; hunks?: [string[], string[]]; nested?: Nested };

const cases: Case[] = [];

// Every edit PHP's SheetOpsTest diffs, both ways.
for (const [name, [edit]] of Object.entries(EDITS)) {
  cases.push({ name: `edit: ${name}`, a: hsWorkbook(), b: edit(hsWorkbook()) });
  cases.push({ name: `edit reversed: ${name}`, a: edit(hsWorkbook()), b: hsWorkbook() });
}

// The seeded random cell edits, both ways.
randomEdits(20260915, 60).forEach(([a, b], run) => {
  cases.push({ name: `random ${run}`, a, b });
  if (run < 20) cases.push({ name: `random reversed ${run}`, a: b, b: a });
});

// Seeded STRUCTURAL edits on a grid built from a tiny alphabet, so rows and
// columns repeat and the alignment has ties to break: this is where two
// correct LCS implementations can still disagree on WHICH row moved.
function grid(rand: (min: number, max: number) => number): Any {
  const cells: Any = {};
  for (let row = 1; row <= 10; row++) {
    for (let col = 0; col < 5; col++) {
      if (rand(0, 4) > 0) cells[String.fromCharCode(65 + col) + row] = { value: ["x", "y", 1, 2][rand(0, 3)] };
    }
  }
  return {
    sheets: [
      { name: "Grid", cells, mergedRegions: [{ start: "B2", end: "C4" }], columnWidths: { 0: 40, 3: 90 }, frozenRows: 1 },
      { name: "Other", cells: { A1: { value: "keep" } } },
    ],
    meta: { creator: "Parity", created: "2026-09-15T00:00:00Z" },
  };
}

{
  const rand = seeded(7);
  const kinds = ["insert_rows", "delete_rows", "insert_columns", "delete_columns"] as const;
  for (let run = 0; run < 30; run++) {
    const a = grid(rand);
    const ops: SheetOp[] = [];
    for (let k = 0, n = rand(1, 3); k < n; k++) {
      ops.push({ type: kinds[rand(0, 3)]!, sheet: "Grid", at: rand(1, 11), count: rand(1, 3) });
    }
    for (let k = 0, n = rand(0, 3); k < n; k++) {
      const address = String.fromCharCode(65 + rand(0, 5)) + rand(1, 12);
      ops.push(rand(0, 2) === 0 ? { type: "clear_cell", sheet: "Grid", address } : { type: "set_cell", sheet: "Grid", address, value: ["x", "y", 1, 2][rand(0, 3)]! });
    }
    const b = Agent.reduce(a, ops);
    cases.push({ name: `structural ${run}`, a, b });
    cases.push({ name: `structural reversed ${run}`, a: b, b: a });
  }
}

// The other branches of the algorithm.
const two = (first: Any, second: Any) => ({ sheets: [first, second], meta: { creator: "P", created: "2026-09-15T00:00:00Z" } });
cases.push(
  {
    name: "duplicate sheet names fall back to set_workbook",
    a: { sheets: [{ name: "Same", cells: { A1: { value: 1 } } }, { name: "Same", cells: { A1: { value: 2 } } }] },
    b: { sheets: [{ name: "Same", cells: { A1: { value: 3 } } }] },
  },
  { name: "a described save is no change", a: hsWorkbook(), b: Agent.read(Agent.toBytes(hsWorkbook())) },
  {
    name: "an authored save is no change",
    a: { sheets: [{ name: "Deals", columns: [{ header: "Name" }, { header: "Value", type: "currency" }], rows: [["Acme", 1200], ["Globex", 800]], totals: { Value: "sum" } }] },
    b: Agent.read(Agent.toBytes({ sheets: [{ name: "Deals", columns: [{ header: "Name" }, { header: "Value", type: "currency" }], rows: [["Acme", 1200], ["Globex", 800]], totals: { Value: "sum" } }] })),
  },
  {
    name: "renames paired by content, the rest removed and added",
    a: { sheets: [{ name: "A", cells: { A1: { value: 1 } } }, { name: "B", cells: { A1: { value: 2 } } }, { name: "C", cells: { A1: { value: 3 } } }] },
    b: { sheets: [{ name: "D", cells: { A1: { value: 9 } } }, { name: "B2", cells: { A1: { value: 2 } } }, { name: "E", cells: { A1: { value: 8 } } }, { name: "A", cells: { A1: { value: 1 } } }] },
  },
  {
    // clear_cell upper-cases its address, so it cannot remove "b1": that sheet is
    // replaced, and the other sheet keeps its one set_cell.
    name: "a lower-case address cannot be cleared, so only that sheet is replaced",
    a: two({ name: "S", cells: { A1: { value: 2 }, b1: { value: 1 }, B2: { value: 5 } } }, { name: "T", cells: { A1: { value: 1 } } }),
    b: two({ name: "S", cells: { A1: { value: 2 }, B2: { value: 5 } } }, { name: "T", cells: { A1: { value: 4 } } }),
  },
  {
    // A row delete re-keys every cell it moves, so there the lower-case key is no obstacle.
    name: "a lower-case address moved by a row delete",
    a: two({ name: "S", cells: { b3: { value: 1 }, A1: { value: 2 } } }, { name: "T", cells: { A1: { value: 1 } } }),
    b: two({ name: "S", cells: { A1: { value: 3 } } }, { name: "T", cells: { A1: { value: 4 } } }),
  },
  {
    name: "a null formula cannot be written, so the sheet is replaced",
    a: two({ name: "S", cells: { A1: { value: 1 } } }, { name: "T", cells: {} }),
    b: two({ name: "S", cells: { A1: { value: 1, formula: null } } }, { name: "T", cells: { B2: { value: "t" } } }),
  },
  {
    name: "rows appended past the end are cells, not inserts",
    a: two({ name: "S", cells: { A1: { value: 1 }, A2: { value: 2 } } }, { name: "T", cells: {} }),
    b: two({ name: "S", cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, A4: { value: 4 } } }, { name: "T", cells: {} }),
  },
  {
    name: "meta removed and panes cleared",
    a: two({ name: "S", cells: { A1: { value: 1 } }, frozenRows: 2, frozenCols: 1, columnWidths: { 2: 30 } }, { name: "T", cells: {} }),
    b: { sheets: [{ name: "S", cells: { A1: { value: 1 } } }, { name: "T", cells: {} }] },
  },
  // PHP 2.3.2 and 2.3.3.
  {
    name: "widths keyed 0..n-1 arrive as a list",
    a: hsWorkbook(),
    b: (() => {
      const w = hsWorkbook();
      w.sheets[0].columnWidths = [120, 80, 140];
      return w;
    })(),
  },
  {
    // clear_cell trims, so it cannot remove the key " A1 " either.
    name: "a padded cell key",
    a: two({ name: "S", cells: { " A1 ": { value: 1 }, B2: { value: 5 } } }, { name: "T", cells: { A1: { value: 1 } } }),
    b: two({ name: "S", cells: { B2: { value: 5 } } }, { name: "T", cells: { A1: { value: 2 } } }),
  },
  {
    name: "a column inserted beside a width key that is not an index",
    a: { sheets: [{ name: "S", cells: { A1: { value: "x" }, B1: { value: "y" } }, columnWidths: { 0: 10, abc: 999, 1: 20 } }] },
    b: { sheets: [{ name: "S", cells: { B1: { value: "x" }, C1: { value: "y" } }, columnWidths: { 1: 10, 2: 20 } }] },
  },
);

// The alignment on its own: short lines from a two-letter alphabet, so nearly
// every step is a tie and the delete-first rule decides the hunks.
{
  const rand = seeded(11);
  for (let run = 0; run < 60; run++) {
    const line = () => Array.from({ length: rand(0, 7) }, () => (rand(0, 1) ? "p" : "q"));
    cases.push({ name: `hunks ${run}`, hunks: [line(), line()] });
  }
}

// Rows moved, not inserted or deleted: ties decide which rows the diff keeps.
{
  const rand = seeded(13);
  for (let run = 0; run < 20; run++) {
    const rows = Array.from({ length: 6 }, (_, i) => ({ value: ["x", "y", "z"][rand(0, 2)], mark: i }));
    const moved = rows.slice();
    const [row] = moved.splice(rand(0, 5), 1);
    moved.splice(rand(0, 5), 0, row!);
    const sheet = (list: typeof rows) => ({
      name: "S",
      cells: Object.fromEntries(list.flatMap((r, i) => [[`A${i + 1}`, { value: r.value }], [`B${i + 1}`, { value: rand(0, 1) }]])),
    });
    cases.push({ name: `rows moved ${run}`, a: { sheets: [sheet(rows)] }, b: { sheets: [sheet(moved)] } });
  }
}

// The reducer on its own, with ops no diff would emit.
cases.push(
  { name: "reduce: set_cell semantics", a: hsWorkbook(), ops: [
    { type: "set_cell", sheet: "Q3", address: "B5", value: 42 },
    { type: "set_cell", sheet: "Q3", address: "a1", value: "Area", format: null },
    { type: "set_cell", sheet: "Notes", address: "A1", value: "Done" },
    { type: "set_cell", sheet: "Q3", address: "C1", formula: "SUM(B2:B4)", computedValue: 3640400.75 },
    { type: "set_cell", sheet: "Q3", address: "Z99", value: null },
    { type: "set_cell", sheet: "Q3", address: "C9", value: null, format: { bold: true } },
    { type: "set_cell", sheet: "Q3", address: "B1", format: null },
    { type: "set_cell", sheet: "Q3", address: "3C", value: 1 },
    { type: "set_cell", sheet: "Nope", address: "A1", value: 1 },
  ] },
  { name: "reduce: set_range and set_workbook", a: hsWorkbook(), ops: [
    { type: "set_range", sheet: "Q3", start: "C2", end: "D3", values: [[1, 2], [3, null]] },
    { type: "set_range", sheet: "Q3", start: "b1", values: [[null, "kept format"]] },
  ] },
  { name: "reduce: shifts", a: { sheets: [{ name: "S", cells: { A1: { value: 1 }, A2: { value: 2 }, C3: { value: 3 }, E5: { value: 5 } }, mergedRegions: [{ start: "A2", end: "C4" }, { start: "D1", end: "E2" }, { start: "B6", end: "B6" }], columnWidths: { 0: 10, 2: 30, 4: 50 } }] }, ops: [
    { type: "insert_rows", sheet: "S", at: 2, count: 2 },
    { type: "delete_rows", sheet: "S", at: 3, count: 3 },
    { type: "insert_columns", sheet: "S", at: 2, count: 1 },
    { type: "delete_columns", sheet: "S", at: 1, count: 2 },
    { type: "delete_rows", sheet: "S", at: 0, count: 1 },
    { type: "insert_columns", sheet: "S", at: 1, count: 0 },
  ] },
  { name: "reduce: a delete that swallows merges and widths", a: { sheets: [{ name: "S", cells: { C1: { value: "c" }, A2: { value: "a" }, B1: { value: "b" } }, mergedRegions: [{ start: "B1", end: "B2" }], columnWidths: { 1: 50 } }] }, ops: [
    { type: "delete_columns", sheet: "S", at: 2, count: 1 },
  ] },
  { name: "reduce: sheets, panes, meta", a: hsWorkbook(), ops: [
    { type: "add_sheet", index: 99, sheet: { name: "New", cells: {} } },
    { type: "move_sheet", sheet: "New", toIndex: 0 },
    { type: "rename_sheet", sheet: "Notes", name: "Sign-off" },
    { type: "replace_sheet", sheet: "Sign-off", data: { name: "Sign-off", cells: { A1: { value: "x" } } } },
    { type: "set_merged_regions", sheet: "Q3", mergedRegions: [] },
    { type: "set_column_widths", sheet: "Q3", columnWidths: { 3: 70 } },
    { type: "set_frozen", sheet: "Q3", rows: 0, cols: 2 },
    { type: "clear_cell", sheet: "Q3", address: "b2" },
    { type: "remove_sheet", sheet: "New" },
    { type: "set_meta", meta: null },
    { type: "nope", sheet: "Q3" },
  ] },
  // PHP 2.3.2: types, addresses and width keys.
  { name: "reduce: a type that is not an op type is skipped", a: hsWorkbook(), ops: [
    { type: true, sheet: "Q3" },
    { type: 0, sheet: "Q3" },
    { type: null, sheet: "Notes" },
    { type: ["remove_sheet"], sheet: "Notes" },
    { type: "REMOVE_SHEET", sheet: "Notes" },
  ] },
  { name: "reduce: a padded address is trimmed", a: hsWorkbook(), ops: [
    { type: "set_cell", sheet: "Q3", address: " b3 ", value: 1 },
    { type: "clear_cell", sheet: "Q3", address: "\t a2\n" },
    { type: "set_cell", sheet: "Q3", address: "\x00c1\x0B", value: 2 },
    { type: "set_cell", sheet: "Q3", address: "\u00a0c2", value: 3 },
    { type: "clear_cell", sheet: "Q3", address: "\u00a0A4" },
    { type: "set_range", sheet: "Q3", start: " d1 ", values: [[4]] },
  ] },
  { name: "reduce: width keys that are not indexes are dropped", a: { sheets: [{ name: "S", cells: { B1: { value: 1 } }, columnWidths: { 0: 10, abc: 999, 1: 20, "1.5": 15, "007": 70, "-1": 5, "-0": 6, "": 7, " 2": 8 } }] }, ops: [
    { type: "insert_columns", sheet: "S", at: 1, count: 1 },
    { type: "delete_columns", sheet: "S", at: 2, count: 1 },
  ] },
  // PHP 2.3.3: positions and counts.
  { name: "reduce: a position or count that is not a number skips the op", a: hsWorkbook(), ops: [
    { type: "move_sheet", sheet: "Notes", toIndex: "last" },
    { type: "add_sheet", index: "end", sheet: { name: "X", cells: {} } },
    { type: "add_sheet", index: null, sheet: { name: "Y", cells: {} } },
    { type: "add_sheet", index: -1.5, sheet: { name: "Z", cells: {} } },
    { type: "set_frozen", sheet: "Q3", rows: "one", cols: 0 },
    { type: "set_frozen", sheet: "Q3", rows: " 1", cols: 0 },
    { type: "set_frozen", sheet: "Q3", rows: 0, cols: null },
    { type: "insert_rows", sheet: "Q3", at: 2, count: "2x" },
    { type: "insert_rows", sheet: "Q3", at: "", count: 1 },
    { type: "delete_columns", sheet: "Q3", at: 1.5, count: 1 },
    { type: "delete_rows", sheet: "Q3", at: "-1", count: 1 },
    { type: "set_cell", sheet: "Q3", address: "A1", value: "skipped", count: "x" },
    { type: "remove_sheet", sheet: "Notes", at: null },
    { type: "rename_sheet", sheet: "Notes", name: "N", toIndex: true },
  ] },
  { name: "reduce: ints, digit strings and absent positions still apply", a: hsWorkbook(), ops: [
    { type: "move_sheet", sheet: "Notes", toIndex: "0" },
    { type: "add_sheet", sheet: { name: "X", cells: {} } },
    { type: "add_sheet", index: "01", sheet: { name: "Y", cells: {} } },
    { type: "add_sheet", index: -3, sheet: { name: "Z", cells: {} } },
    { type: "insert_rows", sheet: "Q3", at: "2", count: "007" },
    { type: "set_frozen", sheet: "Q3", cols: "2" },
    { type: "move_sheet", sheet: "X" },
    { type: "move_sheet", sheet: "Y", toIndex: 99 },
    { type: "delete_rows", sheet: "Q3", at: 3 },
    { type: "delete_columns", sheet: "Q3", at: "1", count: 1 },
  ] },
);

// Where json_encode's depth runs out in SheetDiff::same. PHP builds the nesting
// itself, so none of it crosses JSON.
for (const leaf of [1, [], "x"]) {
  for (const map of [false, true]) {
    for (const depth of [4095, 4096, 4097]) {
      cases.push({ name: `same: ${JSON.stringify(leaf)} in ${depth} ${map ? "maps" : "lists"}`, nested: { leaf, depth, map } });
    }
  }
}

const HAS_PHP = phpAvailable();

if (process.env.CI && !HAS_PHP) {
  throw new Error(
    "php is not on PATH. This suite is the cross-engine ops parity guarantee; " +
      "skipping it in CI would report success with no coverage. Install PHP, " +
      "or set HOLY_SHEET_PHP_SRC and ensure `php` resolves.",
  );
}

describe.skipIf(!HAS_PHP)("cross-engine ops parity (PHP vs TS)", () => {
  let results: Any[];

  beforeAll(() => {
    const dir = mkdtempSync(join(tmpdir(), "holy-sheet-ops-parity-"));
    const file = join(dir, "cases.json");
    writeFileSync(file, JSON.stringify(cases.map(({ a, b, ops, hunks, nested }) => (nested ? { nested } : hunks ? { hunks } : ops ? { a, ops } : { a, b }))));
    results = JSON.parse(php([PHP_SCRIPT, file]).toString("utf8"));
  }, 120_000);

  it("answers every case", () => {
    expect(results).toHaveLength(cases.length);
  });

  cases.forEach((testCase, index) => {
    it(testCase.name, () => {
      const phpResult = results[index];

      if (testCase.nested !== undefined) {
        // PHP throws JsonException past its depth; this port must throw there too, and only there.
        const { leaf, depth, map } = testCase.nested;
        let value: Any = leaf;
        for (let i = 0; i < depth; i++) value = map ? { k: value } : [value];
        let ts: Any;
        try {
          ts = { same: SheetDiff.same(value, value) };
        } catch (error) {
          ts = { error: String(error) };
        }
        expect("error" in ts, `PHP ${JSON.stringify(phpResult)} / TS ${JSON.stringify(ts)}`).toBe("error" in phpResult);
        if (!("error" in ts)) expect(ts.same).toBe(phpResult.same);
        return;
      }

      expect(phpResult.error, "PHP threw").toBeUndefined();

      if (testCase.hunks !== undefined) {
        expect(SheetDiff.hunks(...testCase.hunks)).toEqual(phpResult.hunks);
        return;
      }

      const tsOps = testCase.ops ?? JSON.parse(JSON.stringify(Agent.diff(testCase.a, testCase.b)));

      // The same ops, in the same order, with the same fields in the same order.
      expect(normalize(tsOps)).toEqual(normalize(phpResult.ops));
      expect(tsOps.map((op: Any) => Object.keys(op))).toEqual(phpResult.ops.map((op: Any) => Object.keys(op)));

      // And replaying them gives the same schema.
      const tsReduced = JSON.parse(JSON.stringify(Agent.reduce(testCase.a, tsOps)));
      expect(normalize(tsReduced)).toEqual(normalize(phpResult.reduced));

      if (testCase.b !== undefined) {
        // No ops means the two write the same workbook, not that they are the same JSON.
        expect(tsOps.length > 0 ? SheetDiff.same(tsReduced, testCase.b) : Agent.equivalent(tsReduced, testCase.b)).toBe(true);
      }
    });
  });

  // Parity between two engines that both fell back would pass too. The seeded
  // structural cases exist to compare ALIGNMENTS, so make sure they still do.
  it("compares alignments on the structural cases, not fallbacks", () => {
    const types = cases
      .map((testCase, index) => [testCase.name, results[index].ops ?? []] as const)
      .filter(([name]) => name.startsWith("structural"))
      .flatMap(([, ops]) => ops.map((op: Any) => op.type as string));

    for (const type of ["insert_rows", "delete_rows", "insert_columns", "delete_columns"]) {
      expect(types, type).toContain(type);
    }
    expect(types).not.toContain("replace_sheet");
    expect(types).not.toContain("set_workbook");

    const lowerCase = cases.findIndex((c) => c.name.startsWith("a lower-case address cannot be cleared"));
    expect(results[lowerCase].ops.map((op: Any) => op.type)).toEqual(["replace_sheet", "set_cell"]);
  });

  // The depth cases would pass with both engines throwing everywhere, or nowhere.
  it("finds json_encode's depth limit inside the nested cases, not outside them", () => {
    const outcomes = cases
      .map((testCase, index) => [testCase.nested, results[index]] as const)
      .filter(([nested]) => nested !== undefined)
      .map(([nested, result]) => `${JSON.stringify(nested!.leaf)} ${nested!.depth}: ${"error" in result ? "throws" : "ok"}`);

    expect(outcomes).toContain("1 4096: ok");
    expect(outcomes).toContain("1 4097: throws");
    expect(outcomes).toContain("[] 4095: ok");
    expect(outcomes).toContain("[] 4096: throws");
  });

  it("publishes the same op schema, byte for byte", () => {
    expect(JSON.stringify(Agent.opSchema())).toBe(php([PHP_SCRIPT, "--op-schema"]).toString("utf8"));
    expect(SheetOpSchema.TYPES.length).toBe(17);
  });
});

if (!HAS_PHP) {
  // eslint-disable-next-line no-console
  console.warn("[sheet-ops-parity] php not found on PATH — cross-engine ops parity tests skipped.");
}
