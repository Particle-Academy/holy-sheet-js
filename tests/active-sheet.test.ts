import { describe, it, expect, beforeAll } from "vitest";
import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, unlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Agent, HolySheet, MissingPartException, unzipSync } from "../src";

/**
 * CC-fancy-0050: read() names the workbook's active sheet, the writer emits it
 * when a schema names one, and the `formulas` write option turns off the
 * promotion of a bare "=" string. Driven by the PHP repo's
 * tests/fixtures/active-sheet/cases.json, the same bytes and hand-written
 * expectations all three ports pass. Found like the read-option fixtures
 * (HOLY_SHEET_PHP_SRC, then the sibling checkout); missing is an error, never a skip.
 */
const DIR = process.env.HOLY_SHEET_PHP_SRC
  ? join(process.env.HOLY_SHEET_PHP_SRC, "..", "tests", "fixtures", "active-sheet")
  : join(__dirname, "..", "..", "holy-sheet", "tests", "fixtures", "active-sheet");

if (!existsSync(join(DIR, "cases.json"))) {
  throw new Error(`active-sheet fixtures not found at ${DIR}; check out the PHP holy-sheet repo beside this one or set HOLY_SHEET_PHP_SRC.`);
}

type Any = any; // eslint-disable-line @typescript-eslint/no-explicit-any
const CASES: Any = JSON.parse(readFileSync(join(DIR, "cases.json"), "utf8"));
const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(join(DIR, name)));
const META = { creator: "Parity", created: "2024-01-01T00:00:00Z" };
const dec = new TextDecoder();

function textParts(bytes: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) out[name] = dec.decode(data);
  return out;
}

function bookViewsOf(bytes: Uint8Array): string | null {
  const m = /<bookViews>.*?<\/bookViews>/.exec(textParts(bytes)["xl/workbook.xml"]!);
  return m === null ? null : m[0];
}

/** Each expected cell's `value` and `formula` as read, in the expected order. */
function valueAndFormula(cells: Any, expected: Any): Any {
  const out: Any = {};
  for (const address of Object.keys(expected)) {
    const cell = cells[address] ?? {};
    const kept: Any = {};
    for (const key of ["value", "formula"]) if (key in cell) kept[key] = cell[key];
    out[address] = kept;
  }
  return out;
}

describe("active sheet + formulas option (shared cases)", () => {
  it("has cases to run", () => {
    expect(CASES.read.length).toBeGreaterThan(15);
    expect(CASES.write.length).toBeGreaterThan(3);
    expect(CASES.invalid.length).toBeGreaterThan(3);
    expect(CASES.formulas.length).toBe(3);
    expect(CASES.badOptions.length).toBe(3);
  });

  for (const c of CASES.read) {
    it(`reads: ${c.name}`, () => {
      const schema = Agent.read(fixture(c.file), c.options);
      expect(schema["activeSheet"] ?? null).toEqual(c.activeSheet);
    });
  }

  for (const c of CASES.readRaises) {
    it(`raises: ${c.name}`, () => {
      let error: unknown;
      try {
        Agent.read(fixture(c.file), c.options);
      } catch (e) {
        error = e;
      }
      expect(error).toBeInstanceOf(MissingPartException);
      expect((error as MissingPartException).part).toBe(c.part);
    });
  }

  it("describes the active sheet from a path too", async () => {
    expect((await Agent.describe(join(DIR, "active-1.xlsx")))["activeSheet"]).toEqual({ index: 1, name: "Second" });
  });

  it("places activeSheet right after sheets", () => {
    expect(Object.keys(Agent.read(fixture("active-1.xlsx"))).slice(0, 2)).toEqual(["sheets", "activeSheet"]);
  });

  for (const c of CASES.write) {
    it(`writes: ${c.name}`, () => {
      const bytes = Agent.toBytes(c.schema);
      expect(bookViewsOf(bytes)).toBe(c.bookViews);
      expect(Agent.read(bytes)["activeSheet"]).toEqual(c.activeSheet);
    });
  }

  for (const c of CASES.invalid) {
    it(`refuses: ${c.name}`, () => {
      expect(Agent.validate(c.schema)).toEqual(c.errors);
    });
  }

  it("round-trips a read workbook with its active sheet", () => {
    const described = Agent.read(fixture("active-2.xlsx"));
    expect(Agent.validate(described)).toEqual([]);
    expect(Agent.read(Agent.toBytes(described))["activeSheet"]).toEqual({ index: 2, name: "Third" });
  });

  for (const c of CASES.formulas) {
    it(`formulas: ${c.name}`, () => {
      const read: Any = Agent.read(Agent.toBytes(CASES.formulaSchema, c.options));
      expect(valueAndFormula(read.sheets[0].cells, c.cells)).toEqual(c.cells);
      expect(valueAndFormula(read.sheets[1].cells, c.rows)).toEqual(c.rows);
    });
  }

  it("writes byte-identical parts under promote and under no option", () => {
    const plain = textParts(Agent.toBytes(CASES.formulaSchema));
    const promote = textParts(Agent.toBytes(CASES.formulaSchema, { formulas: "promote" }));
    expect(promote).toEqual(plain);
  });

  it("takes the formulas option on write() and on the HolySheet instance", async () => {
    const path = join(mkdtempSync(join(tmpdir(), "hs-formulas-")), "out.xlsx");
    try {
      await Agent.write(CASES.formulaSchema, path, { formulas: "explicit" });
      expect(((await Agent.describe(path)) as Any).sheets[0].cells.A1.value).toBe("=SUM(B1:B2)");
      await new HolySheet().write(CASES.formulaSchema, path, { formulas: "explicit" });
      expect(((await Agent.describe(path)) as Any).sheets[0].cells.A1.value).toBe("=SUM(B1:B2)");
    } finally {
      if (existsSync(path)) unlinkSync(path);
    }
    const bytes = new HolySheet().toBytes(CASES.formulaSchema, { formulas: "explicit" });
    expect((Agent.read(bytes) as Any).sheets[0].cells.A1.value).toBe("=SUM(B1:B2)");
  });

  for (const c of CASES.badOptions) {
    it(`refuses the bad write option: ${c.name}`, () => {
      expect(() => Agent.toBytes(CASES.formulaSchema, c.options)).toThrow(c.message);
    });
  }
});

// Part parity with the PHP reference for every shared write and formulas case,
// the options included. Skips without php LOCALLY; in CI a missing php throws,
// as tests/parity.test.ts does.
const PHP_SCRIPT = join(__dirname, "..", "scripts", "php-tobytes.php");

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

const HAS_PHP = phpAvailable();
if (process.env.CI && !HAS_PHP) {
  throw new Error("php is not on PATH. The active-sheet parity cases are cross-engine coverage; skipping them in CI would report success with none.");
}

describe.skipIf(!HAS_PHP)("active sheet + formulas option: PHP part parity", () => {
  let dir: string;
  beforeAll(() => {
    dir = mkdtempSync(join(tmpdir(), "holy-sheet-active-parity-"));
  });

  function phpBytes(name: string, schema: unknown, options: unknown): Uint8Array {
    const schemaFile = join(dir, `${name}.json`);
    const optionsFile = join(dir, `${name}.options.json`);
    const out = join(dir, `${name}.php.xlsx`);
    writeFileSync(schemaFile, JSON.stringify(schema));
    writeFileSync(optionsFile, JSON.stringify(options));
    php([PHP_SCRIPT, schemaFile, out, optionsFile]);
    return new Uint8Array(readFileSync(out));
  }

  const cases: Array<[string, Any, Any]> = [
    ...CASES.write.map((c: Any, i: number): [string, Any, Any] => [`write ${i}: ${c.name}`, { ...c.schema, meta: META }, {}]),
    ...CASES.formulas.map((c: Any, i: number): [string, Any, Any] => [`formulas ${i}: ${c.name}`, CASES.formulaSchema, c.options]),
  ];

  for (const [i, [name, schema, options]] of cases.entries()) {
    it(`emits byte-identical OOXML parts: ${name}`, () => {
      const phpParts = textParts(phpBytes(`case${i}`, schema, options));
      const tsParts = textParts(Agent.toBytes(schema, options));
      expect(Object.keys(tsParts).sort()).toEqual(Object.keys(phpParts).sort());
      for (const part of Object.keys(phpParts)) expect(tsParts[part], `part ${part} differs`).toBe(phpParts[part]);
      // A php process per case; under a full parallel run that outlasts the 5s default.
    }, 60_000);
  }
});
