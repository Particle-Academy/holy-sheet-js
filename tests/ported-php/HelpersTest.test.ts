import { describe, it, expect } from "vitest";
import { mkdtempSync, writeFileSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Agent, Inference } from "../../src";

// Ported from PHP tests/Unit/HelpersTest.php. The TS Agent.fromCsv takes CSV
// *content* (isomorphic), not a path; the "from a file path" PHP case is ported
// by reading the file ourselves and passing the string (same inference spec).

describe("helpers (ported PHP HelpersTest)", () => {
  it("infers integer column from header + values", () => {
    const col = Inference.detect([1, 2, 3], "count");
    expect(col.type).toBe("integer");
  });

  it("infers currency from header pattern", () => {
    const col = Inference.detect([1.5, 2.0], "revenue");
    expect(col.type).toBe("currency");
    expect((col as any).currency).toBe("USD");
  });

  it("infers percent only when values are in [0,1]", () => {
    const pct = Inference.detect([0.1, 0.5], "growth_rate");
    const not = Inference.detect([10, 50], "growth_rate");
    expect(pct.type).toBe("percent");
    expect(not.type).toBe("integer");
  });

  it("infers date when values match ISO date", () => {
    const col = Inference.detect(["2024-01-01", "2024-02-01"], "created");
    expect(col.type).toBe("date");
  });

  it("falls back to auto for mixed types", () => {
    const col = Inference.detect([1, "two", true], "mixed");
    expect(col.type).toBe("auto");
  });

  it("builds a schema from rows + headers via fromArray", () => {
    const schema = Agent.fromArray(
      [
        ["NA", 100],
        ["EU", 200],
      ],
      ["Region", "Revenue"],
    );
    expect(schema.sheets[0].name).toBe("Sheet 1");
    expect(schema.sheets[0].columns![1].type).toBe("currency");
    expect(schema.sheets[0].rows).toHaveLength(2);
  });

  it("treats first row as headers when omitted", () => {
    const schema = Agent.fromArray([
      ["Name", "Age"],
      ["Alice", 30],
      ["Bob", 42],
    ]);
    expect(schema.sheets[0].columns![0].header).toBe("Name");
    expect(schema.sheets[0].rows).toHaveLength(2);
  });

  it("builds a schema from CSV string via fromCsv", () => {
    const schema = Agent.fromCsv("Name,Age\nAlice,30\nBob,42");
    expect(schema.sheets[0].columns![0].header).toBe("Name");
    expect(schema.sheets[0].columns![1].type).toBe("integer");
    expect(schema.sheets[0].rows).toEqual([
      ["Alice", 30],
      ["Bob", 42],
    ]);
  });

  it("handles CSV with quoted fields and embedded newlines", () => {
    const csv = 'Name,Note\n"Alice","line 1\nline 2"\n';
    const schema = Agent.fromCsv(csv);
    expect(schema.sheets[0].rows[0][1]).toBe("line 1\nline 2");
  });

  it("builds CSV from a file path (read content ourselves in Node)", () => {
    const dir = mkdtempSync(join(tmpdir(), "holy-csv-"));
    const tmp = join(dir, "cities.csv");
    writeFileSync(tmp, "City,Pop\nNYC,8000000\nLA,4000000\n");
    const schema = Agent.fromCsv(readFileSync(tmp, "utf8"));
    expect(schema.sheets[0].rows).toEqual([
      ["NYC", 8000000],
      ["LA", 4000000],
    ]);
  });
});
