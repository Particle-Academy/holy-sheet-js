import { describe, expect, it } from "vitest";
import { Agent } from "../src";

/*
 * A cell's `formula` with a leading "=" is read the same way by every entry
 * point. Ported from PHP tests/Unit/LeadingEqualsFormulaTest.php.
 *
 * write() always stripped the "=" while lint() parsed the raw text and reported
 * `#NAME?` "unexpected '='", so an agent following lint -> ops -> reduce ->
 * lint wrote "=A2*2", was told it was broken, and looped.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

function base(): Any {
  return Agent.validateAndRepair({
    sheets: [
      {
        name: "Model",
        columns: [
          { header: "Base", type: "number" },
          { header: "Scenario", type: "number" },
        ],
        rows: [[100, 0]],
      },
    ],
    meta: { creator: "Test", created: "2024-01-01T00:00:00Z" },
  }).schema;
}

function set(formula: string): Any {
  return Agent.reduce(base(), [{ type: "set_cell", sheet: "Model", address: "B2", formula }]);
}

const bytes = (schema: Any): string => Buffer.from(Agent.toBytes(schema)).toString("base64");

describe("a formula with a leading =", () => {
  it("lints a set_cell formula written with a leading = clean, as write reads it", () => {
    expect(Agent.lint(set("=A2*2"))).toEqual([]);
    expect(Agent.lint(set("A2*2"))).toEqual([]);
  });

  it("writes the = form and the canonical form byte for byte the same", () => {
    expect(bytes(set("=A2*2"))).toBe(bytes(set("A2*2")));
  });

  it("reads a repeated leading = the way write does", () => {
    expect(Agent.lint(set("==A2*2"))).toEqual([]);
    expect(bytes(set("==A2*2"))).toBe(bytes(set("A2*2")));
  });

  it("agrees in validate, validateAndRepair and equivalent", () => {
    const a = set("=A2*2");
    const repaired = Agent.validateAndRepair(a);
    expect(Agent.validate(a)).toEqual([]);
    expect(repaired.errors).toEqual([]);
    expect(Agent.lint(repaired.schema)).toEqual([]);
    expect(Agent.equivalent(a, set("A2*2"))).toBe(true);
  });

  it("lints an object cell formula with a leading = in a cells map clean", () => {
    const schema = { sheets: [{ name: "S", cells: { A1: { value: 2 }, B1: { formula: "=A1*2" }, C1: "==A1*3" } }] };
    expect(Agent.lint(schema)).toEqual([]);
  });

  it("still reports a real error in an = formula, with the canonical text", () => {
    const issues = Agent.lint(set("=A2/0"));
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#DIV/0!");
    expect(issues[0].formula).toBe("A2/0");
  });

  it("describes the canonical formula form in opSchema", () => {
    const setCell = (Agent.opSchema() as Any).oneOf.find((v: Any) => v.properties.type.const === "set_cell");
    const formula = setCell.properties.formula;
    expect(formula.type).toBe("string");
    expect(formula.description ?? "").toContain("without the leading '='");
    expect(formula.description ?? "").toContain("a leading '=' is accepted and stripped");
  });
});
