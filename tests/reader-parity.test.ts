import { describe, it, expect, beforeAll } from "vitest";
import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Agent } from "../src";

// Cross-engine READER parity: for the same .xlsx file, the PHP holy-sheet
// reader (scripts/php-describe.php) and this TS port's reader should recover the
// SAME schema. We generate each file with the TS writer (Agent.toBytes), hand it
// to the PHP reader (capturing its JSON), and structurally compare to the TS
// reader's output.
//
// PHP/JSON ambiguity: PHP json_encode serializes an empty associative array as
// `[]` not `{}`, and may key-order differently. So we do a STRUCTURAL deep
// compare with a normalize() that treats empty [] and empty {} as equal, sorts
// object keys, and compares numbers by value. (Cell map key order in particular
// differs: PHP preserves insertion order, JS too, but both engines emit cells
// row-major from the same parse, so they line up — normalize() sorts anyway to
// be safe.)
//
// Skips automatically when `php` isn't on PATH.

const PHP_SCRIPT = join(__dirname, "..", "scripts", "php-describe.php");

function php(args: string[], opts: Parameters<typeof execFileSync>[2] = {}): Buffer {
  return execFileSync("php", args, { shell: true, ...opts }) as Buffer;
}

function phpAvailable(): boolean {
  try {
    php(["--version"], { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

const META = { creator: "Parity", created: "2024-01-01T00:00:00Z" };

const SCHEMAS: Record<string, unknown> = {
  sparse: {
    meta: META,
    sheets: [
      {
        name: "Data",
        cells: {
          A1: { value: "Region" },
          A2: { value: "North" },
          B1: { value: "Revenue" },
          B2: { value: 12000 },
          B3: { value: 9800.5 },
          B4: { formula: "SUM(B2:B3)" },
          C1: { value: true },
        },
      },
    ],
  },
  dateColumn: {
    meta: META,
    sheets: [
      {
        name: "Dates",
        cells: {
          A1: { value: "When" },
          A2: { value: "2024-03-15", format: { displayFormat: "date" } },
          A3: { value: "2025-12-31", format: { displayFormat: "date" } },
        },
      },
    ],
  },
  comments: {
    meta: META,
    sheets: [
      {
        name: "Annotated",
        cells: {
          A1: { value: "x", comment: { text: "a note", author: "Ada" } },
          B1: { value: "y", comment: { text: "another", author: "Linus" } },
        },
      },
    ],
  },
  merges: {
    meta: META,
    sheets: [
      {
        name: "Merged",
        cells: {
          A1: { value: "Header", format: { textAlign: "center" } },
          A2: { value: 1 },
          B2: { value: 2 },
        },
        mergedRegions: [{ start: "A1", end: "B1" }],
      },
    ],
  },
  frozen: {
    meta: META,
    sheets: [
      {
        name: "Frozen",
        cells: {
          A1: { value: "H1" },
          B1: { value: "H2" },
          A2: { value: "r1c1" },
          B2: { value: "r1c2" },
        },
        frozenRows: 1,
        frozenCols: 1,
      },
    ],
  },
  multiSheet: {
    meta: META,
    sheets: [
      { name: "One", cells: { A1: { value: 1 } } },
      { name: "Two", cells: { A1: { value: 2 }, A2: { value: "two" } } },
    ],
  },
};

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const isPlainObject = (v: unknown): v is Record<string, unknown> =>
  typeof v === "object" && v !== null && !Array.isArray(v);

const isEmpty = (v: unknown): boolean =>
  (Array.isArray(v) && v.length === 0) || (isPlainObject(v) && Object.keys(v).length === 0);

/**
 * Normalize a parsed-JSON tree for cross-engine structural comparison:
 *   - empty [] and empty {} collapse to the same sentinel,
 *   - object keys are sorted (so cell-map / format key order is irrelevant),
 *   - numbers compare by value (already the case once parsed).
 */
function normalize(value: Any): Any {
  if (isEmpty(value)) return "∅empty";
  if (Array.isArray(value)) return value.map(normalize);
  if (isPlainObject(value)) {
    const out: Record<string, unknown> = {};
    for (const key of Object.keys(value).sort()) {
      out[key] = normalize(value[key]);
    }
    return out;
  }
  return value;
}

const HAS_PHP = phpAvailable();

describe.skipIf(!HAS_PHP)("cross-engine reader parity (PHP vs TS)", () => {
  let dir: string;
  beforeAll(() => {
    dir = mkdtempSync(join(tmpdir(), "holy-sheet-reader-parity-"));
  });

  for (const [name, schema] of Object.entries(SCHEMAS)) {
    it(`readers agree on workbook content: ${name}`, () => {
      const bytes = Agent.toBytes(schema);
      const xlsxFile = join(dir, `${name}.xlsx`);
      writeFileSync(xlsxFile, bytes);

      const phpJson = php([PHP_SCRIPT, xlsxFile]).toString("utf8");
      const phpSchema = JSON.parse(phpJson);
      const tsSchema = Agent.read(bytes);

      expect(normalize(tsSchema)).toEqual(normalize(phpSchema));
    });
  }
});

if (!HAS_PHP) {
  // eslint-disable-next-line no-console
  console.warn("[reader-parity] php not found on PATH — cross-engine reader parity tests skipped.");
}
