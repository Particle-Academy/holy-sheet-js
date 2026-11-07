import { describe, it, expect, beforeAll } from "vitest";
import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Csv, OptionsInvalid } from "../src";

/**
 * CC-fancy-0012: this port's CSV writer against PHP's `Exchange\Csv::toCsv()`,
 * call for call, text and report. The inputs live in the PHP repo
 * (tests/fixtures/csv/tables.json) so the Python port runs the same ones.
 * Skips locally without php; in CI a missing php THROWS, as in parity.test.ts.
 */
const PHP_SCRIPT = join(__dirname, "..", "scripts", "php-csv.php");
const TABLES = process.env.HOLY_SHEET_PHP_SRC
  ? join(process.env.HOLY_SHEET_PHP_SRC, "..", "tests", "fixtures", "csv", "tables.json")
  : join(__dirname, "..", "..", "holy-sheet", "tests", "fixtures", "csv", "tables.json");
if (!existsSync(TABLES)) throw new Error(`CSV parity inputs not found at ${TABLES}; check out the PHP holy-sheet repo beside this one.`);

type Any = any; // eslint-disable-line @typescript-eslint/no-explicit-any
const CALLS: Any[] = JSON.parse(readFileSync(TABLES, "utf8")).calls;

function phpAvailable(): boolean {
  try {
    execFileSync("php", ["--version"], { shell: true, stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}
const HAS_PHP = phpAvailable();
if (process.env.CI && !HAS_PHP) throw new Error("php is not on PATH; the CSV parity suite cannot run, and skipping it in CI would report success with no coverage.");

function ours(call: Any): Any {
  try {
    return Csv.toCsv(call.table, call.options);
  } catch (e) {
    if (e instanceof OptionsInvalid) return { error: "OptionsInvalid", message: e.message };
    throw e;
  }
}

describe.skipIf(!HAS_PHP)("CSV parity (PHP vs TS)", () => {
  let php: Any[] = [];
  beforeAll(() => {
    const dir = mkdtempSync(join(tmpdir(), "holy-sheet-csv-"));
    const file = join(dir, "calls.json");
    writeFileSync(file, JSON.stringify(CALLS));
    php = JSON.parse(execFileSync("php", [PHP_SCRIPT, file], { shell: true }).toString("utf8"));
  });

  it("compared every call", () => {
    expect(CALLS.length).toBeGreaterThan(12);
    expect(php).toHaveLength(CALLS.length);
  });

  CALLS.forEach((call, i) => {
    it(`matches PHP: ${call.name}`, () => {
      expect(ours(call)).toEqual(php[i]);
    });
  });
});

describe("CSV writer (pinned bytes, no PHP needed)", () => {
  it("quotes RFC 4180 and ends every line CRLF", () => {
    const { text } = Csv.toCsv({ rows: [{ cells: [{ t: "string", v: 'a,"b"' }, { t: "number", v: 1.5 }] }, { cells: [{ t: "bool", v: true }, { t: "empty" }] }] });
    expect(text).toBe('"a,""b""",1.5\r\nTRUE,\r\n');
  });

  it("escapes a leading formula character by default, and reports where", () => {
    const out = Csv.toCsv({ rows: [{ cells: [{ t: "string", v: "=SUM(A1)" }, { t: "number", v: -1 }] }] });
    expect(out.text).toBe("'=SUM(A1),-1\r\n");
    expect(out.report.approximated).toEqual([{ kind: "csv-escaped", count: 1, where: ["/rows/0/cells/0"] }]);
  });

  it("leaves it alone only when the host says so", () => {
    expect(Csv.toCsv({ rows: [{ cells: [{ t: "string", v: "=1" }] }] }, { escapeFormulas: false }).text).toBe("=1\r\n");
  });

  it("refuses an option it does not know", () => {
    expect(() => Csv.toCsv({ rows: [] }, { delimiter: ";" } as Any)).toThrow(OptionsInvalid);
  });
});
