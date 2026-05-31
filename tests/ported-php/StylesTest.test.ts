import { describe, it, expect } from "vitest";
import { Agent, unzipSync } from "../../src";

// Ported from PHP tests/Unit/StylesTest.php. PHP wrote a file then re-opened
// with ZipArchive; we go straight through Agent.toBytes + unzipSync and assert
// the same XML substrings on the same parts.

const dec = new TextDecoder();
function parts(bytes: Uint8Array): Record<string, string> {
  const files = unzipSync(bytes);
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(files)) out[k] = dec.decode(v as Uint8Array);
  return out;
}

describe("styles (ported PHP StylesTest)", () => {
  it("produces a styles.xml part in every workbook", () => {
    const p = parts(Agent.toBytes({ sheets: [{ name: "X", columns: [{ header: "A" }], rows: [[1]] }] }));
    expect(p["xl/styles.xml"]).toBeDefined();
    for (const needle of ["<styleSheet", "<fonts", "<fills", "<borders", "<cellXfs"]) {
      expect(p["xl/styles.xml"], `missing ${needle} in styles.xml`).toContain(needle);
    }
  });

  it("emits bold + colored header style when default theme is on", () => {
    const p = parts(
      Agent.toBytes({
        sheets: [{ name: "X", columns: [{ header: "A" }, { header: "B" }], rows: [[1, 2]], theme: "default" }],
      }),
    );
    expect(p["xl/styles.xml"]).toContain("<b/>");
    expect(/<c r="A1" s="\d+"/.test(p["xl/worksheets/sheet1.xml"]!)).toBe(true);
  });

  it("applies currency number format on a currency column", () => {
    const p = parts(
      Agent.toBytes({
        sheets: [
          {
            name: "X",
            columns: [{ header: "Amount", type: "currency", currency: "USD" }],
            rows: [[1234.56]],
            theme: "plain",
          },
        ],
      }),
    );
    expect(p["xl/styles.xml"]).toContain("&quot;$&quot;#,##0.00");
  });

  it("applies percent format on a percent column", () => {
    const p = parts(
      Agent.toBytes({
        sheets: [
          { name: "X", columns: [{ header: "Rate", type: "percent", decimals: 2 }], rows: [[0.124]], theme: "plain" },
        ],
      }),
    );
    expect(p["xl/styles.xml"]).toContain("0.00%");
  });

  it("converts ISO date strings to Excel serial numbers on date columns", () => {
    const p = parts(
      Agent.toBytes({
        sheets: [{ name: "X", columns: [{ header: "When", type: "date" }], rows: [["2026-05-01"]], theme: "plain" }],
      }),
    );
    // 2026-05-01 → 46143 (serial days since 1899-12-30)
    expect(p["xl/worksheets/sheet1.xml"]).toContain("<v>46143</v>");
  });

  it("emits merged regions, frozen panes, and column widths", () => {
    const p = parts(
      Agent.toBytes({
        sheets: [
          {
            name: "X",
            cells: { A1: { value: "Hello" } },
            mergedRegions: [{ start: "A1", end: "C1" }],
            frozenRows: 1,
            frozenCols: 0,
            columnWidths: { 0: 200 },
          },
        ],
      }),
    );
    const ws = p["xl/worksheets/sheet1.xml"]!;
    expect(ws).toContain('<mergeCell ref="A1:C1"/>');
    expect(ws).toContain("<sheetView");
    expect(ws).toContain('ySplit="1"');
    expect(ws).toContain('<col min="1" max="1"');
  });

  it("writes the symbolic totals row with SUM/AVG formulas", () => {
    const p = parts(
      Agent.toBytes({
        sheets: [
          {
            name: "X",
            columns: [
              { header: "Region" },
              { header: "Revenue", type: "number" },
              { header: "YoY", type: "percent" },
            ],
            rows: [
              ["A", 100, 0.1],
              ["B", 200, 0.2],
              ["C", 300, 0.3],
            ],
            totals: { Revenue: "sum", YoY: "avg" },
            theme: "plain",
          },
        ],
      }),
    );
    const ws = p["xl/worksheets/sheet1.xml"]!;
    expect(ws).toContain("<f>SUM(B2:B4)</f>");
    expect(ws).toContain("<f>AVERAGE(C2:C4)</f>");
  });

  it("writes comments xml + vml drawing when a cell has a comment", () => {
    const p = parts(
      Agent.toBytes({
        sheets: [
          {
            name: "X",
            cells: { A1: { value: "Hello", comment: { text: "Note this cell", author: "Agent" } } },
          },
        ],
      }),
    );
    expect(p["xl/comments1.xml"]).toBeDefined();
    expect(p["xl/drawings/vmlDrawing1.vml"]).toBeDefined();
    expect(p["xl/worksheets/_rels/sheet1.xml.rels"]).toBeDefined();
    expect(p["xl/comments1.xml"]).toContain("Note this cell");
    expect(p["xl/comments1.xml"]).toContain("<author>Agent</author>");
  });

  it("emits cached formula values when computedValue is provided", () => {
    const p = parts(
      Agent.toBytes({
        sheets: [
          {
            name: "X",
            cells: {
              A1: { value: 10 },
              A2: { value: 20 },
              A3: { formula: "SUM(A1:A2)", computedValue: 30 },
            },
          },
        ],
      }),
    );
    expect(p["xl/worksheets/sheet1.xml"]).toContain("<f>SUM(A1:A2)</f><v>30</v>");
  });
});
