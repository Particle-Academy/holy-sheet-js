import { describe, expect, it } from "vitest";
import { execFileSync } from "node:child_process";
import { join } from "node:path";
import { Agent } from "../src";
import { odsFixture, odsFixturePath } from "./ods-fixtures";

// Cross-engine ODS READER parity: for each ODS fixture in the PHP repo, the PHP
// reader (scripts/php-describe.php) and this port's reader must recover the SAME
// schema. edge.ods exists for this suite: it holds what LibreOffice rewrites on
// save, so it is the fixture where two readers written from one design are most
// likely to part ways.
//
// Unlike reader-parity.test.ts this compares key ORDER too, except where PHP's
// JSON cannot express it: an empty cell map is `[]` in PHP and `{}` here.
//
// Skips when `php` is not on PATH, and throws instead in CI, for the reason
// reader-parity.test.ts gives.

const PHP_SCRIPT = join(__dirname, "..", "scripts", "php-describe.php");
const FIXTURES = ["workbook.ods", "native.ods", "edge.ods", "workbook.xlsx"];

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

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/** PHP encodes an empty associative array as []; make both sides say {}. */
function emptyCellsAsObject(schema: Any): Any {
  for (const sheet of schema.sheets ?? []) {
    if (Array.isArray(sheet.cells) && sheet.cells.length === 0) sheet.cells = {};
  }
  return schema;
}

const HAS_PHP = phpAvailable();

if (process.env.CI && !HAS_PHP) {
  throw new Error("php is not on PATH. This suite is the cross-engine ODS reader parity guarantee; skipping it in CI would report success with no coverage.");
}

describe.skipIf(!HAS_PHP)("cross-engine ODS reader parity (PHP vs TS)", () => {
  for (const name of FIXTURES) {
    it(`readers agree on ${name}, key order included`, () => {
      const phpSchema = emptyCellsAsObject(JSON.parse(php([PHP_SCRIPT, `"${odsFixturePath(name)}"`]).toString("utf8")));
      const tsSchema = Agent.read(odsFixture(name));

      expect(JSON.stringify(tsSchema, null, 1)).toBe(JSON.stringify(phpSchema, null, 1));
    });
  }
});
