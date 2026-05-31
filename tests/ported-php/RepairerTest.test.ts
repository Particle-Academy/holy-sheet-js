import { describe, it, expect } from "vitest";
import { Agent } from "../../src";

// Ported from PHP tests/Unit/RepairerTest.php.

describe("repairer (ported PHP RepairerTest)", () => {
  it("renames singular sheet to sheets", () => {
    const result = Agent.validateAndRepair({ sheet: [{ name: "A", rows: [] }] });
    expect(result.schema).toHaveProperty("sheets");
    expect(result.schema).not.toHaveProperty("sheet");
    expect(result.repairs).not.toHaveLength(0);
  });

  it("renames row to rows", () => {
    const result = Agent.validateAndRepair({ sheets: [{ name: "A", row: [["x"]] }] });
    const sheet = (result.schema as any).sheets[0];
    expect(sheet).toHaveProperty("rows");
    expect(sheet).not.toHaveProperty("row");
  });

  it("converts object-keyed rows to indexed list", () => {
    const result = Agent.validateAndRepair({
      sheets: [{ name: "A", rows: { "0": ["a"], "1": ["b"] } }],
    });
    const rows = (result.schema as any).sheets[0].rows;
    expect(rows).toEqual([["a"], ["b"]]);
  });

  it("coerces stringified numerics in number columns", () => {
    const result = Agent.validateAndRepair({
      sheets: [
        {
          name: "A",
          columns: [{ header: "X", type: "number" }],
          rows: [["1.5"], ["2"]],
        },
      ],
    });
    const rows = (result.schema as any).sheets[0].rows;
    expect(rows[0][0]).toBe(1.5);
    expect(Number.isInteger(rows[1][0])).toBe(true);
  });

  it("replaces unknown theme with default", () => {
    const result = Agent.validateAndRepair({
      sheets: [{ name: "A", theme: "wonkyland", rows: [] }],
    });
    expect((result.schema as any).sheets[0].theme).toBe("default");
  });

  it("passes valid schemas through unchanged", () => {
    const valid = { sheets: [{ name: "A", rows: [["x"]] }] };
    const result = Agent.validateAndRepair(structuredClone(valid));
    expect(result.schema).toEqual(valid);
    expect(result.repairs).toEqual([]);
  });

  it("does not invent missing required fields", () => {
    const result = Agent.validateAndRepair({ sheets: [{ rows: [] }] });
    expect(result.errors).not.toHaveLength(0);
  });
});
