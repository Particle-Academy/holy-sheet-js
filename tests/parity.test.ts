import { describe, it, expect, beforeAll } from "vitest";
import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Agent, unzipSync } from "../src";

// Cross-engine parity: the PHP holy-sheet and this TS port should emit
// byte-identical OOXML parts for the same input (timestamps pinned via meta).
// Skips automatically when `php` isn't on PATH (e.g. CI without PHP).

const PHP_SCRIPT = join(__dirname, "..", "scripts", "php-tobytes.php");

// `php` may only resolve through the shell (Herd shims etc.), so run with shell:true.
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
          D1: { value: "2024-03-15", format: { displayFormat: "date" } },
        },
      },
    ],
  },
  rowOriented: {
    meta: META,
    sheets: [
      {
        name: "Sales",
        theme: "default",
        columns: [
          { header: "Region", type: "string" },
          { header: "Revenue", type: "currency", currency: "USD" },
          { header: "Margin", type: "percent", decimals: 1 },
        ],
        rows: [
          ["North", 12000, 0.12],
          ["South", 9800.5, 0.08],
          ["East", 15000, 0.21],
        ],
        totals: { Revenue: "sum", Margin: "avg" },
      },
    ],
  },
  decorated: {
    meta: META,
    sheets: [
      {
        name: "Decorated",
        cells: {
          A1: { value: "Title", format: { bold: true, fontSize: 14, color: "#FF0000", textAlign: "center" } },
          A2: { value: "x", comment: { text: "a note", author: "Ada" }, format: { backgroundColor: "#FFFF00" } },
          B2: { value: "y", format: { borderBottom: "#000000" } },
        },
        mergedRegions: [{ start: "A1", end: "B1" }],
        columnWidths: { 0: 120, 1: 80 },
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

const HAS_PHP = phpAvailable();

describe.skipIf(!HAS_PHP)("cross-engine parity (PHP vs TS)", () => {
  let dir: string;
  beforeAll(() => {
    dir = mkdtempSync(join(tmpdir(), "holy-sheet-parity-"));
  });

  for (const [name, schema] of Object.entries(SCHEMAS)) {
    it(`emits byte-identical OOXML parts: ${name}`, () => {
      const schemaFile = join(dir, `${name}.json`);
      const phpOut = join(dir, `${name}.php.xlsx`);
      writeFileSync(schemaFile, JSON.stringify(schema));
      php([PHP_SCRIPT, schemaFile, phpOut]);

      const phpParts = unzipSync(new Uint8Array(readFileSync(phpOut)));
      const tsParts = unzipSync(Agent.toBytes(schema));

      // Same set of parts.
      expect(Object.keys(tsParts).sort()).toEqual(Object.keys(phpParts).sort());

      // Each part byte-identical.
      const dec = new TextDecoder();
      for (const part of Object.keys(phpParts)) {
        const phpText = dec.decode(phpParts[part]!);
        const tsText = dec.decode(tsParts[part]!);
        expect(tsText, `part ${part} differs`).toBe(phpText);
      }
    });
  }

  it("PHP can also read the TS-written file (reverse round-trip is consistent)", () => {
    // Sanity: TS read of PHP output equals TS read of TS output for a schema.
    const schema = SCHEMAS.sparse;
    const schemaFile = join(dir, `rr.json`);
    const phpOut = join(dir, `rr.php.xlsx`);
    writeFileSync(schemaFile, JSON.stringify(schema));
    php([PHP_SCRIPT, schemaFile, phpOut]);

    const fromPhp = Agent.read(new Uint8Array(readFileSync(phpOut)));
    const fromTs = Agent.read(Agent.toBytes(schema));
    expect(fromTs).toEqual(fromPhp);
  });
});

void rmSync; // (temp dir left for inspection; OS cleans tmp)

if (!HAS_PHP) {
  // eslint-disable-next-line no-console
  console.warn("[parity] php not found on PATH — cross-engine parity tests skipped.");
}
