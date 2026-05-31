import { describe, it, expect } from "vitest";
import { mkdtempSync } from "node:fs";
import { existsSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Agent, HolySheet, unzipSync } from "../../src";

// Ported from PHP tests/Unit/WriterTest.php. Where PHP used Agent::write to a
// real path, we do the same (Agent.write is async, Node-only). The PHP
// HolySheet::writeFile static convenience is NOT present in the TS port (the TS
// HolySheet is an instance facade only) — that one assertion is noted below.

const dir = mkdtempSync(join(tmpdir(), "holy-writer-"));

describe("writer (ported PHP WriterTest)", () => {
  it("writes a minimum viable xlsx", async () => {
    const schema = {
      sheets: [
        {
          name: "Sheet 1",
          columns: [{ header: "Name" }, { header: "Age", type: "integer" }],
          rows: [
            ["Alice", 30],
            ["Bob", 42],
          ],
        },
      ],
    };
    const tmp = join(dir, "mvp.xlsx");
    const result = await Agent.write(schema, tmp);

    expect(result).toMatchObject({ path: tmp, sheets: 1 });
    expect(result.bytes).toBeGreaterThan(0);
    expect(existsSync(tmp)).toBe(true);

    // First 4 bytes "PK\x03\x04".
    const head = readFileSync(tmp).subarray(0, 4);
    expect([...head]).toEqual([0x50, 0x4b, 0x03, 0x04]);
  });

  it("writes a workbook with multiple sheets", async () => {
    const schema = {
      sheets: [
        { name: "A", columns: [{ header: "X" }], rows: [[1], [2]] },
        { name: "B", columns: [{ header: "Y" }], rows: [[3], [4]] },
      ],
    };
    const tmp = join(dir, "multi.xlsx");
    const result = await Agent.write(schema, tmp);
    expect(result.sheets).toBe(2);
  });

  it("accepts the fancy-sheets-style sparse cells map", async () => {
    const schema = {
      sheets: [
        {
          name: "Sparse",
          cells: {
            A1: { value: "Header" },
            A2: { value: 100 },
            B2: { value: 200 },
            A3: { formula: "SUM(A2:B2)" },
          },
        },
      ],
    };
    const tmp = join(dir, "sparse.xlsx");
    const result = await Agent.write(schema, tmp);
    expect(result.bytes).toBeGreaterThan(0);
  });

  it("the produced xlsx contains the mandatory parts", () => {
    const bytes = Agent.toBytes({
      sheets: [{ name: "Test", columns: [{ header: "A" }], rows: [[1]] }],
    });
    const files = unzipSync(bytes);
    for (const entry of [
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels",
      "xl/worksheets/sheet1.xml",
      "docProps/core.xml",
      "docProps/app.xml",
    ]) {
      expect(Object.keys(files), `missing required xlsx entry: ${entry}`).toContain(entry);
    }
  });

  // PHP: HolySheet::writeFile static convenience. The TS HolySheet facade is
  // instance-only and has no static writeFile, so we exercise the instance
  // .write() path instead (same effect: a file lands on disk).
  it("HolySheet instance write lands a file on disk", async () => {
    const tmp = join(dir, "facade.xlsx");
    const hs = new HolySheet();
    await hs.write({ sheets: [{ name: "X", columns: [{ header: "A" }], rows: [[1]] }] }, tmp);
    expect(existsSync(tmp)).toBe(true);
  });

  it("the HolySheet singleton mirrors Agent for facade/DI use", async () => {
    const hs = new HolySheet();
    const tmp = join(dir, "singleton.xlsx");
    const result = await hs.write({ sheets: [{ name: "X", columns: [{ header: "A" }], rows: [[1]] }] }, tmp);
    expect(result.sheets).toBe(1);
    expect(hs.validate({ sheets: [] })).not.toEqual([]);
    const bytes = hs.toBytes({ sheets: [{ name: "Y", columns: [{ header: "A" }], rows: [[1]] }] });
    expect([...bytes.subarray(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04]);
    expect(hs.toolDefinition()).toHaveProperty("$schema");
    expect(typeof hs.getVersion()).toBe("string");
  });
});
