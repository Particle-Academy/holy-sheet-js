import { describe, it, expect } from "vitest";
import { Agent, SchemaException } from "../../src";

// Ported from PHP tests/Unit/ValidatorTest.php.

describe("validator (ported PHP ValidatorTest)", () => {
  it("returns no errors for a valid schema", () => {
    const errors = Agent.validate({
      sheets: [{ name: "OK", columns: [{ header: "A" }], rows: [[1]] }],
    });
    expect(errors).toEqual([]);
  });

  it("flags missing top-level sheets", () => {
    const errors = Agent.validate({});
    expect(errors).toHaveLength(1);
    expect(errors[0].path).toBe("sheets");
    expect(errors[0].expected).toBe("array");
  });

  it("flags an empty sheets array", () => {
    const errors = Agent.validate({ sheets: [] });
    expect(errors[0].path).toBe("sheets");
    expect(errors[0].expected).toContain("non-empty");
  });

  it("flags a sheet missing its name", () => {
    const errors = Agent.validate({ sheets: [{ columns: [{ header: "A" }], rows: [[1]] }] });
    expect(errors[0].path).toBe("sheets[0].name");
  });

  it("flags an unknown column type", () => {
    const errors = Agent.validate({
      sheets: [{ name: "X", columns: [{ header: "A", type: "banana" }], rows: [[1]] }],
    });
    expect(errors).toHaveLength(1);
    expect(errors[0].path).toBe("sheets[0].columns[0].type");
    expect(typeof errors[0].hint).toBe("string");
  });

  it("flags an unknown theme", () => {
    const errors = Agent.validate({
      sheets: [{ name: "X", columns: [{ header: "A" }], rows: [[1]], theme: "neon" }],
    });
    expect(errors[0].path).toBe("sheets[0].theme");
  });

  it("throws SchemaException on assert with structured errors", () => {
    expect(() => Agent.toBytes({})).toThrow(SchemaException);
  });
});
