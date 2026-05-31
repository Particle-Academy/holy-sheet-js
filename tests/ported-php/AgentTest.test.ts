import { describe, it, expect } from "vitest";
import { Agent } from "../../src";

// Ported from PHP tests/Unit/AgentTest.php. describe() is async/Node-only in TS.

describe("agent (ported PHP AgentTest)", () => {
  it("returns the JSON schema as a parsed object from toolDefinition", () => {
    const def = Agent.toolDefinition() as any;
    expect(def).toBeTypeOf("object");
    expect(def).toHaveProperty("$schema");
    expect(def.title).toBe("Holy Sheet workbook schema");
    expect(def.definitions).toHaveProperty("Sheet");
    expect(def.definitions).toHaveProperty("Column");
    expect(def.definitions).toHaveProperty("CellData");
    expect(def.definitions).toHaveProperty("CellFormat");
  });

  it("describe() returns not_found for a missing path", async () => {
    const result = (await Agent.describe("/this/path/does/not/exist.xlsx")) as any;
    expect(result.error).toBe("not_found");
  });

  it("toBytes returns a non-empty xlsx byte string", () => {
    const bytes = Agent.toBytes({ sheets: [{ name: "X", columns: [{ header: "A" }], rows: [[1]] }] });
    expect(bytes.length).toBeGreaterThan(100);
    expect([...bytes.subarray(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04]);
  });
});
