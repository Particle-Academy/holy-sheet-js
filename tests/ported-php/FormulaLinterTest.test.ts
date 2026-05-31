import { describe, it, expect } from "vitest";
import { Agent } from "../../src";

// Ported from PHP tests/Unit/FormulaLinterTest.php.

describe("formula linter (ported PHP FormulaLinterTest)", () => {
  it("returns no issues for a workbook with valid formulas", () => {
    const schema = {
      sheets: [
        {
          name: "Q4",
          rows: [
            ["Region", "Revenue", "Doubled"],
            ["NA", 100, { formula: "B2*2" }],
            ["EU", 200, { formula: "B3*2" }],
            ["Total", { formula: "SUM(B2:B3)" }, { formula: "SUM(C2:C3)" }],
          ],
        },
      ],
    };
    expect(Agent.lint(schema)).toEqual([]);
  });

  it("catches the header-row off-by-one bug and suggests the correct row", () => {
    const schema = {
      sheets: [
        {
          name: "Q4",
          rows: [
            ["Region", "Annual", "Monthly"],
            ["NA", 12000, { formula: "B1*12" }],
          ],
        },
      ],
    };
    const issues = Agent.lint(schema);
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#VALUE!");
    expect(issues[0].address).toBe("C2");
    expect(issues[0].hint).toContain('B1 = "Annual" (string)');
    expect(issues[0].hint).toContain("Did you mean B2");
  });

  it("catches division by zero", () => {
    const schema = {
      sheets: [{ name: "D", rows: [["x"], [0], [{ formula: "100/A2" }]] }],
    };
    const issues = Agent.lint(schema);
    expect(issues[0].error).toBe("#DIV/0!");
  });

  it("catches circular references", () => {
    const schema = {
      sheets: [
        {
          name: "C",
          rows: [["x"], [{ formula: "A1+1" }], [{ formula: "A2" }]],
        },
      ],
    };
    const issues = Agent.lint(schema);
    expect(issues.length).toBeGreaterThan(0);
  });

  it("detects true circular dependency (A1 = B1, B1 = A1)", () => {
    const schema = {
      sheets: [
        {
          name: "C",
          cells: {
            A1: { formula: "B1" },
            B1: { formula: "A1" },
          },
        },
      ],
    };
    const issues = Agent.lint(schema);
    expect(issues).not.toHaveLength(0);
    expect(issues[0].error).toBe("#CIRC!");
  });

  it("flags unknown function names as #NAME?", () => {
    const schema = {
      sheets: [{ name: "F", rows: [["x"], [{ formula: "BOGUSFN(1,2)" }]] }],
    };
    expect(Agent.lint(schema)[0].error).toBe("#NAME?");
  });

  it("evaluates cross-sheet references", () => {
    const schema = {
      sheets: [
        { name: "Detail", rows: [["x"], [100], [200]] },
        {
          name: "Summary",
          cells: {
            A1: { value: "Total" },
            B1: { formula: "SUM(Detail!A2:A3)" },
          },
        },
      ],
    };
    expect(Agent.lint(schema)).toEqual([]);
  });

  it("handles SUM across a numeric column with no errors", () => {
    const schema = {
      sheets: [
        {
          name: "S",
          rows: [["x"], [10], [20], [30], [{ formula: "SUM(A2:A4)" }], [{ formula: "AVERAGE(A2:A4)" }]],
        },
      ],
    };
    expect(Agent.lint(schema)).toEqual([]);
  });

  it("catches arithmetic on a string in the middle of an expression", () => {
    const schema = {
      sheets: [
        {
          name: "M",
          rows: [
            ["x", "y"],
            [10, "oops"],
            [20, 30],
            [{ formula: "A2+B2" }],
          ],
        },
      ],
    };
    const issues = Agent.lint(schema);
    expect(issues[0].error).toBe("#VALUE!");
    expect(issues[0].hint).toContain('B2 = "oops"');
  });
});
