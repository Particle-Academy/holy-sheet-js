import { describe, expect, it } from "vitest";
import { Agent, unzipSync } from "../src";
import { Normalizer } from "../src/schema/normalizer";

/**
 * An explicit `{value: <string>}` cell is a STRING cell, written exactly as
 * given.
 *
 * The object form is how a caller says "I mean this literally" (it is already
 * the documented escape hatch for a leading "="). Numeric-string coercion ran
 * on it anyway, so `{value: "+1"}` was written as the number 1 and
 * `{value: "007"}` as 7. Bare strings keep their promotion rules unchanged.
 *
 * Mirrors PHP `tests/Unit/ExplicitStringValueTest.php` and Python
 * `tests/test_explicit_string_value.py`; keep the three in step.
 */
type Any = Record<string, any>;

const NUMERIC_LOOKING = ["+1", "007", "1e5", "-0", "1.50", " 12 ", "0.0", "-3"];

const cellsOf = (schema: Any) => new Normalizer().normalize(schema).sheets[0]!.cells;

function written(schema: Any, ref: string): { xml: string; read: Any } {
  const bytes = Agent.toBytes(schema);
  const sheet = new TextDecoder().decode(unzipSync(bytes)["xl/worksheets/sheet1.xml"]!);
  const m = new RegExp(`<c r="${ref}"[ >].*?(?:/>|</c>)`).exec(sheet);
  return { xml: m ? m[0] : `(no ${ref})`, read: Agent.read(bytes) as Any };
}

describe("an explicit {value: string} cell", () => {
  it.each(NUMERIC_LOOKING)("is a string cell in a cells map: %j", (s) => {
    const schema = { sheets: [{ name: "S", cells: { A1: { value: s } } }] };

    expect(cellsOf(schema).get("A1")!.value).toBe(s);

    const { xml, read } = written(schema, "A1");
    expect(xml).toContain('t="inlineStr"');
    expect(xml).toContain(`<t xml:space="preserve">${s}</t>`);
    expect(read.sheets[0].cells.A1.value).toBe(s);
  });

  it.each(NUMERIC_LOOKING)("is a string cell in a row: %j", (s) => {
    const schema = { sheets: [{ name: "S", columns: [{ header: "Code" }], rows: [[{ value: s }]] }] };

    expect(cellsOf(schema).get("A2")!.value).toBe(s);

    const { xml, read } = written(schema, "A2");
    expect(xml).toContain('t="inlineStr"');
    expect(read.sheets[0].cells.A2.value).toBe(s);
  });

  it("stays a string under a numeric column type and a number format", () => {
    const cells = cellsOf({
      sheets: [{
        name: "S",
        columns: [{ header: "N", type: "number" }, { header: "C", type: "currency" }],
        rows: [[{ value: "+1" }, { value: "007", format: { displayFormat: "currency" } }]],
      }],
    });

    expect(cells.get("A2")!.value).toBe("+1");
    expect(cells.get("B2")!.value).toBe("007");
  });

  it("does not promote a leading-= string to a formula", () => {
    const cell = cellsOf({ sheets: [{ name: "S", cells: { A1: { value: "=1+1" } } }] }).get("A1")!;

    expect(cell.formula).toBeNull();
    expect(cell.value).toBe("=1+1");
  });

  it("leaves bare strings, explicit numbers and explicit dates as they were", () => {
    const cells = cellsOf({
      sheets: [{
        name: "S",
        cells: {
          A1: "+1",
          A2: "=SUM(B1:B2)",
          A3: { value: 1 },
          A4: { value: "2024-03-15", format: { displayFormat: "date" } },
        },
      }],
    });

    expect(cells.get("A1")!.value).toBe(1);
    expect(cells.get("A2")!.formula).toBe("SUM(B1:B2)");
    expect(cells.get("A3")!.value).toBe(1);
    expect(cells.get("A4")!.value).toBe(45366);

    const rows = cellsOf({ sheets: [{ name: "S", columns: [{ header: "N" }], rows: [["+1"]] }] });
    expect(rows.get("A2")!.value).toBe(1);
  });
});
