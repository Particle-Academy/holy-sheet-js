import { ReadLimitExceeded } from "../exceptions";
import { DEFAULT_READ_LIMITS, readAddress } from "./read-limits";

export interface RowsOptions {
  /** `"computed"` (default): a formula gives its cached computedValue. `"formula"`: the formula, with its leading `=`. */
  formulas?: "computed" | "formula";
  /** The most positions the view may hold, padding included. Default: the `maxCells` read limit. */
  maxCells?: number;
}

type Any = any; // eslint-disable-line @typescript-eslint/no-explicit-any

/**
 * A read sheet as rows: a list of rows, each a list of values in column order,
 * row 1 and column A at index 0. Mirrors PHP `Reader\RowsView`.
 *
 * A blank row in the middle is `[]`, so indexes are row numbers minus one;
 * trailing blank rows are not padded, and a row stops at its last value. A
 * merged region's value is at its top-left cell only. The view is sized before
 * it is built and refuses past `maxCells` positions.
 */
export function rowsView(sheet: Record<string, unknown>, options: RowsOptions = {}): unknown[][] {
  for (const key of Object.keys(options)) {
    if (key !== "formulas" && key !== "maxCells") throw new Error(`[holy-sheet] unknown rows option: ${key}`);
  }
  const mode = options.formulas ?? "computed";
  if (mode !== "computed" && mode !== "formula") {
    throw new Error('[holy-sheet] rows option formulas must be "computed" or "formula"');
  }
  const maxCells = options.maxCells ?? DEFAULT_READ_LIMITS.maxCells;
  if (!Number.isSafeInteger(maxCells) || maxCells < 1) {
    throw new Error("[holy-sheet] rows option maxCells must be a positive integer");
  }

  const covered = coveredRegions(sheet["mergedRegions"]);
  const values = new Map<number, Map<number, unknown>>();
  const cells = sheet["cells"];
  if (cells !== null && typeof cells === "object" && !Array.isArray(cells)) {
    for (const [address, cell] of Object.entries(cells as Record<string, unknown>)) {
      const at = readAddress(address);
      if (at === null) continue;
      const [col, row] = at;
      if (row < 1 || isCovered(covered, col, row)) continue;
      const value = cellValue(cell, mode);
      if (value === null || value === undefined) continue;
      let byCol = values.get(row);
      if (!byCol) values.set(row, (byCol = new Map()));
      byCol.set(col, value);
    }
  }
  if (values.size === 0) return [];

  let size = 0;
  for (const byCol of values.values()) size += maxKey(byCol) + 1;
  if (size > maxCells) throw new ReadLimitExceeded("maxCells", maxCells, size, "the rows view would hold");

  const out: unknown[][] = [];
  const last = maxKey(values);
  for (let r = 1; r <= last; r++) {
    const byCol = values.get(r);
    if (!byCol) {
      out.push([]);
      continue;
    }
    const line: unknown[] = new Array(maxKey(byCol) + 1).fill(null);
    for (const [c, v] of byCol) line[c] = v;
    out.push(line);
  }
  return out;
}

function cellValue(cell: unknown, mode: "computed" | "formula"): unknown {
  if (cell === null || typeof cell !== "object" || Array.isArray(cell)) {
    // A bare scalar; a bare "=..." string is a formula, as the writer reads it.
    if (typeof cell === "string" && cell.length > 1 && cell[0] === "=") return mode === "formula" ? cell : null;
    return cell ?? null;
  }
  const c = cell as Any;
  const formula = c.formula;
  if (typeof formula === "string" && formula !== "") {
    if (mode === "formula") return formula[0] === "=" ? formula : "=" + formula;
    return c.computedValue ?? null;
  }
  return c.value ?? null;
}

function coveredRegions(regions: unknown): Array<[number, number, number, number]> {
  const out: Array<[number, number, number, number]> = [];
  for (const m of Array.isArray(regions) ? regions : []) {
    const a = m && typeof m.start === "string" ? readAddress(m.start) : null;
    const b = m && typeof m.end === "string" ? readAddress(m.end) : null;
    if (a && b) out.push([Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1])]);
  }
  return out;
}

function isCovered(covered: Array<[number, number, number, number]>, col: number, row: number): boolean {
  for (const [c1, r1, c2, r2] of covered) {
    if (col >= c1 && col <= c2 && row >= r1 && row <= r2 && (col !== c1 || row !== r1)) return true;
  }
  return false;
}

/** The largest key, without spreading a million keys into one call. */
function maxKey(map: Map<number, unknown>): number {
  let max = -1;
  for (const k of map.keys()) if (k > max) max = k;
  return max;
}
