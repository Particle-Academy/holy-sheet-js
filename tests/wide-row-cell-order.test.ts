import { describe, expect, it } from "vitest";
import { Agent, CellAddress, unzipSync } from "../src";

/**
 * Cells within a row are written in ascending column INDEX: A..Z, AA, AB, ...
 * Mirrors the PHP engine's WideRowCellOrderTest.
 *
 * OOXML requires it, and the old lexicographic order (A, AA, AB, ..., B, C)
 * loses data in Google Sheets: on 2026-10-28 a workbook whose row 1 was written
 * A1, AA1..AE1, B1..Z1 came back from a Sheets import/export with only A1.
 */
type Any = Record<string, unknown>;

function wideCells(): Record<string, unknown> {
  const cells: Record<string, unknown> = {};
  for (let i = 31; i >= 0; i--) cells[`${CellAddress.letter(i)}1`] = i;
  // A sparse row, inserted backwards.
  cells.AZ2 = "az";
  cells.B2 = "b";
  cells.AA2 = "aa";
  cells.Z2 = "z";
  return cells;
}

const schema = (): Any => ({ sheets: [{ name: "Wide", cells: wideCells() }] });

function columns(xml: string, row: number): string[] {
  return [...xml.matchAll(new RegExp(`<c r="([A-Z]+)${row}"`, "g"))].map((m) => m[1]!);
}

describe("cell order within a wide row", () => {
  it("writes the cells of a row wider than 26 columns in column-index order", () => {
    const files = unzipSync(Agent.toBytes(schema() as never));
    const xml = new TextDecoder().decode(files["xl/worksheets/sheet1.xml"]!);

    expect(columns(xml, 1)).toEqual(Array.from({ length: 32 }, (_, i) => CellAddress.letter(i)));
    expect(columns(xml, 2)).toEqual(["B", "Z", "AA", "AZ"]);
  });

  it("round-trips every cell of a wide sheet", () => {
    const read = Agent.read(Agent.toBytes(schema() as never)) as Any;
    const sheet = (read.sheets as Any[])[0]!;
    const got: Record<string, unknown> = {};
    for (const [address, cell] of Object.entries(sheet.cells as Any)) {
      got[address] = cell !== null && typeof cell === "object" ? (cell as Any).value : cell;
    }
    expect(got).toEqual(wideCells());
  });
});
