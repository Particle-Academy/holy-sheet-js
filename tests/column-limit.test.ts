import { describe, expect, it } from "vitest";
import { Agent, unzipSync, zipSync } from "../src/index";

/**
 * A worksheet's `<col min max>` range is refused past Excel's column limit,
 * 16,384, before the reader walks it. Mirrors PHP tests/Unit/ColumnLimitTest.php.
 */
const enc = new TextEncoder();
const dec = new TextDecoder();

function workbookWithCols(cols: string): Uint8Array {
  const parts = unzipSync(
    Agent.toBytes({ sheets: [{ name: "Sheet1", cells: { A1: { value: 1 } } }] } as never),
  );
  const sheet = dec.decode(parts["xl/worksheets/sheet1.xml"]!);
  expect(sheet).toContain("<sheetData");
  parts["xl/worksheets/sheet1.xml"] = enc.encode(sheet.replace("<sheetData", `<cols>${cols}</cols><sheetData`));
  return zipSync(Object.entries(parts).map(([name, data]) => ({ name, data })));
}

describe("column range limit", () => {
  it("refuses a <col> whose max is past 16,384", () => {
    const bytes = workbookWithCols('<col min="1" max="2000000000" width="10" customWidth="1"/>');
    expect(() => Agent.read(bytes)).toThrow(/16,384/);
  });

  it("refuses a <col> whose min is past 16,384", () => {
    const bytes = workbookWithCols('<col min="16385" max="16385" width="10" customWidth="1"/>');
    expect(() => Agent.read(bytes)).toThrow(/16,384/);
  });

  it("still reads a <col> that ends exactly at column 16,384", () => {
    const bytes = workbookWithCols('<col min="16384" max="16384" width="10" customWidth="1"/>');
    const schema = Agent.read(bytes) as { sheets: { name: string }[] };
    expect(schema.sheets[0]!.name).toBe("Sheet1");
  });
});
