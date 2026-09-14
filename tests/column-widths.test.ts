/**
 * `columnWidths` entries that are not a column index and a width.
 *
 * Before 2.4.3 the three writers disagreed: this one turned the key "abc" into
 * NaN with parseInt and wrote a NaN column, PHP overwrote column A, and Python
 * raised. One rule now, mirroring PHP holy-sheet 2.3.4's tests.
 */
import { describe, expect, it } from "vitest";
import { Agent, SchemaException } from "../src";
import { Normalizer } from "../src/schema/normalizer";

const sheet = (widths: unknown) => ({
  sheets: [{ name: "S", cells: { A1: { value: "x" } }, columnWidths: widths }],
});

describe("columnWidths", () => {
  it("never writes a NaN column or overwrites a column for a key that is not an index", () => {
    const workbook = new Normalizer().normalize(sheet({ 0: 120, abc: 999, 1: 80 }) as never);
    const widths = (workbook.sheets[0] as unknown as { columnWidths: Map<number, number> }).columnWidths;

    expect([...widths.entries()]).toEqual([
      [0, 120],
      [1, 80],
    ]);
  });

  it("reports every bad entry by path instead of writing it", () => {
    const errors = Agent.validate(sheet({ abc: 50, "-1": 50, "16384": 50, B: 50, 0: "wide", 1: -5 }) as never);

    // JavaScript orders integer-like keys first, so compare as sets.
    expect(errors.map((e) => e.path).sort()).toEqual(
      [
        "sheets[0].columnWidths.abc",
        "sheets[0].columnWidths.-1",
        "sheets[0].columnWidths.16384",
        "sheets[0].columnWidths.B",
        "sheets[0].columnWidths.0",
        "sheets[0].columnWidths.1",
      ].sort(),
    );

    expect(() => Agent.toBytes(sheet({ abc: 999 }) as never)).toThrow(SchemaException);
  });

  it("accepts indexes as digit strings, widths as numbers or digit strings, and a list", () => {
    expect(Agent.validate(sheet({ 0: 120, 1: "80", 2: 140.5, 16383: 10 }) as never)).toEqual([]);
    expect(Agent.validate(sheet([120, 80]) as never)).toEqual([]);
  });

  it("repairs a letter key to its index and drops what it cannot repair", () => {
    const result = Agent.validateAndRepair(sheet({ B: 90, abc: 5, 0: "wide", 2: 60 }) as never);

    expect(result.errors).toEqual([]);
    expect((result.schema as { sheets: { columnWidths: unknown }[] }).sheets[0]!.columnWidths).toEqual({ 1: 90, 2: 60 });
    expect(result.repairs).toContain("converted 'sheets[0].columnWidths.B' to column index 1");
    expect(result.repairs).toContain("dropped 'sheets[0].columnWidths.abc' (not a column index and a width)");
    expect(result.repairs).toContain("dropped 'sheets[0].columnWidths.0' (not a column index and a width)");
  });
});
