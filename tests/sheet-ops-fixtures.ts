/**
 * The workbook and the edits PHP's tests/Unit/SheetOpsTest.php diffs, shared by
 * the unit port (sheet-ops.test.ts) and the PHP parity suite
 * (sheet-ops-parity.test.ts) so both exercise exactly the same cases.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type Any = any;

export function hsWorkbook(): Any {
  return {
    sheets: [
      {
        name: "Q3",
        cells: {
          A1: { value: "Region", format: { bold: true } },
          B1: { value: "Revenue", format: { bold: true } },
          A2: { value: "North" },
          B2: { value: 1250000.5, format: { displayFormat: "currency", decimals: 2 } },
          A3: { value: "South" },
          B3: { value: 980400.25 },
          A4: { value: "West" },
          B4: { value: 1410000 },
          A5: { value: "Total", format: { bold: true } },
          B5: { value: null, formula: "SUM(B2:B4)" },
        },
        mergedRegions: [{ start: "A7", end: "B7" }],
        columnWidths: { 0: 120, 1: 140 },
        frozenRows: 1,
      },
      { name: "Notes", cells: { A1: { value: "Checked by finance", comment: { text: "Signed off", author: "CFO" } } } },
    ],
    meta: { creator: "MOIC", created: "2026-09-14T00:00:00Z" },
  };
}

export function hsSet(w: Any, sheet: string, address: string, cell: Any): Any {
  for (const s of w.sheets) {
    if (s.name === sheet) {
      s.cells[address] = cell;
    }
  }
  return w;
}

function renumberRows(cells: Any, map: (column: string, row: number) => string | null): Any {
  const out: Any = {};
  for (const [address, cell] of Object.entries(cells)) {
    const m = /^([A-Z]+)(\d+)$/.exec(address)!;
    const to = map(m[1]!, Number(m[2]));
    if (to !== null) out[to] = cell;
  }
  return out;
}

/** [edit, the exact op types PHP's diff gives for it]. */
export const EDITS: Record<string, [(w: Any) => Any, string[]]> = {
  "one value": [(w) => hsSet(w, "Q3", "B3", { value: 990000 }), ["set_cell"]],
  "one format": [
    (w) => hsSet(w, "Q3", "B2", { value: 1250000.5, format: { displayFormat: "currency", decimals: 0 } }),
    ["set_cell"],
  ],
  "a comment removed": [(w) => hsSet(w, "Notes", "A1", { value: "Checked by finance" }), ["set_cell"]],
  "a formula": [(w) => hsSet(w, "Q3", "B5", { value: null, formula: "SUM(B2:B3)" }), ["set_cell"]],
  "a cell cleared": [
    (w) => {
      delete w.sheets[0].cells.A4;
      return w;
    },
    ["clear_cell"],
  ],
  "a row inserted": [
    (w) => {
      const cells = renumberRows(w.sheets[0].cells, (c, row) => c + (row >= 3 ? row + 1 : row));
      cells.A3 = { value: "East" };
      cells.B3 = { value: 505000 };
      cells.B6 = { value: null, formula: "SUM(B2:B5)" };
      w.sheets[0].cells = cells;
      w.sheets[0].mergedRegions = [{ start: "A8", end: "B8" }];
      return w;
    },
    ["insert_rows", "set_cell", "set_cell", "set_cell"],
  ],
  "a row deleted": [
    (w) => {
      const cells = renumberRows(w.sheets[0].cells, (c, row) => (row === 3 ? null : c + (row > 3 ? row - 1 : row)));
      cells.B4 = { value: null, formula: "SUM(B2:B3)" };
      w.sheets[0].cells = cells;
      w.sheets[0].mergedRegions = [{ start: "A6", end: "B6" }];
      return w;
    },
    ["delete_rows", "set_cell"],
  ],
  "a column inserted": [
    (w) => {
      const cells: Any = {};
      for (const [address, cell] of Object.entries(w.sheets[0].cells)) {
        cells[address.replaceAll("B", "C")] = cell;
      }
      cells.B1 = { value: "Units", format: { bold: true } };
      cells.B2 = { value: 12 };
      w.sheets[0].cells = cells;
      w.sheets[0].mergedRegions = [{ start: "A7", end: "C7" }];
      w.sheets[0].columnWidths = { 0: 120, 1: 80, 2: 140 };
      return w;
    },
    ["insert_columns", "set_cell", "set_cell", "set_column_widths"],
  ],
  "a column deleted": [
    (w) => {
      w.sheets[0].cells = Object.fromEntries(Object.entries(w.sheets[0].cells).filter(([a]) => a[0] !== "B"));
      w.sheets[0].mergedRegions = [{ start: "A7", end: "A7" }];
      w.sheets[0].columnWidths = { 0: 120 };
      return w;
    },
    ["delete_columns"],
  ],
  "a sheet added": [
    (w) => {
      w.sheets.splice(1, 0, { name: "Q4", cells: { A1: { value: "Region" } } });
      return w;
    },
    ["add_sheet"],
  ],
  "a sheet removed": [
    (w) => {
      w.sheets.pop();
      return w;
    },
    ["remove_sheet"],
  ],
  "a sheet renamed": [
    (w) => {
      w.sheets[1].name = "Sign-off";
      return w;
    },
    ["rename_sheet"],
  ],
  "a sheet renamed and edited": [
    (w) => {
      w.sheets[1].name = "Sign-off";
      w.sheets[1].cells.A2 = { value: "Approved" };
      return w;
    },
    ["rename_sheet", "set_cell"],
  ],
  "sheets reordered": [
    (w) => {
      w.sheets.reverse();
      return w;
    },
    ["move_sheet"],
  ],
  "merges, widths, panes": [
    (w) => {
      delete w.sheets[0].mergedRegions;
      w.sheets[0].columnWidths = { 0: 200, 1: 140 };
      w.sheets[0].frozenCols = 1;
      return w;
    },
    ["set_merged_regions", "set_column_widths", "set_frozen"],
  ],
  meta: [
    (w) => {
      w.meta.creator = "Compass";
      return w;
    },
    ["set_meta"],
  ],
  "an authored sheet changed": [
    (w) => {
      w.sheets[1] = { name: "Notes", columns: [{ header: "Item" }, { header: "Owner" }], rows: [["Budget", "CFO"]] };
      return w;
    },
    ["replace_sheet"],
  ],
};

/** mulberry32: a small deterministic PRNG. Any fixed-seed generator will do; PHP's mt_rand sequence is not needed. */
export function seeded(seed: number): (min: number, max: number) => number {
  let state = seed >>> 0;
  return (min, max) => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    const unit = ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    return min + Math.floor(unit * (max - min + 1));
  };
}

/**
 * The seeded random cell-form edits of PHP's "keeps the round trip over a
 * seeded run" test: 1-5 edits per run to A1:D9 of the first sheet, each a value,
 * a formatted string, a formula, or a removal.
 */
export function randomEdits(seed: number, runs: number): [Any, Any][] {
  const rand = seeded(seed);
  const pairs: [Any, Any][] = [];

  for (let run = 0; run < runs; run++) {
    const a = hsWorkbook();
    const b = hsWorkbook();
    let cells = b.sheets[0].cells;

    for (let k = 0, edits = rand(1, 5); k < edits; k++) {
      const address = String.fromCharCode(65 + rand(0, 3)) + rand(1, 9);
      switch (rand(0, 3)) {
        case 0:
          cells[address] = { value: rand(1, 999) };
          break;
        case 1:
          cells[address] = { value: "x" + rand(1, 9), format: { italic: true } };
          break;
        case 2:
          cells[address] = { value: null, formula: "SUM(B2:B" + rand(3, 6) + ")" };
          break;
        default:
          cells = Object.fromEntries(Object.entries(cells).filter(([key]) => key !== address));
          b.sheets[0].cells = cells;
      }
    }

    pairs.push([a, b]);
  }

  return pairs;
}
