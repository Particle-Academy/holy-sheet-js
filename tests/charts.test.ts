import { describe, it, expect } from "vitest";
import { Agent, unzipSync } from "../src";

/**
 * Native charts. Mirrors PHP tests/Unit/ChartsTest.php case for case.
 *
 * A chart is a real Excel chart over CELL RANGES, not a picture: edit a figure
 * and the chart follows. These cases also keep the parity fixture honest, since
 * a byte diff of two engines that both wrote no chart parts would pass.
 */
function chWorkbook(charts: unknown[], extra: Record<string, unknown> = {}): any {
  return {
    sheets: [
      {
        name: "Q3 Revenue",
        cells: {
          A4: "Region", B4: "Revenue", C4: "Growth",
          A5: "North", B5: 1250000.5, C5: 0.184,
          A6: "South", B6: 980400.25, C6: -0.042,
          A7: "EMEA", B7: 2100000, C7: 0.311,
        },
        charts,
        ...extra,
      },
    ],
  };
}

function chParts(schema: unknown): Record<string, string> {
  const dec = new TextDecoder();
  const out: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(unzipSync(Agent.toBytes(schema as never)))) out[name] = dec.decode(bytes);
  return out;
}

const CH_REVENUE = {
  type: "column",
  title: "Q3 revenue by region",
  categories: "A5:A7",
  series: [{ name: "Revenue", values: "B5:B7" }],
  position: { from: "E2", to: "L18" },
};

describe("charts", () => {
  it("writes a chart over live cell ranges, wired into the sheet", () => {
    const parts = chParts(chWorkbook([CH_REVENUE]));
    expect(parts["xl/worksheets/sheet1.xml"]).toContain('<drawing r:id="rId1"/>');
    expect(parts["xl/worksheets/_rels/sheet1.xml.rels"]).toContain('Target="../drawings/drawing1.xml"');
    expect(parts["xl/drawings/_rels/drawing1.xml.rels"]).toContain('Target="../charts/chart1.xml"');
    expect(parts["[Content_Types].xml"]).toContain('PartName="/xl/drawings/drawing1.xml"');
    expect(parts["[Content_Types].xml"]).toContain('PartName="/xl/charts/chart1.xml"');

    const chart = parts["xl/charts/chart1.xml"]!;
    expect(chart).toContain('<c:barChart><c:barDir val="col"/>');
    expect(chart).toContain("<c:f>'Q3 Revenue'!$B$5:$B$7</c:f>");
    expect(chart).toContain("<c:f>'Q3 Revenue'!$A$5:$A$7</c:f>");
    expect(chart).toContain('<c:pt idx="2"><c:v>2100000</c:v></c:pt>');
    expect(chart).toContain('<c:pt idx="0"><c:v>North</c:v></c:pt>');
    expect(chart).toContain("Q3 revenue by region");
    expect(chart).not.toContain("<c:legend>");
  });

  it("anchors the chart between the two cells it is given", () => {
    const drawing = chParts(chWorkbook([CH_REVENUE]))["xl/drawings/drawing1.xml"];
    expect(drawing).toContain(
      "<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>",
    );
    expect(drawing).toContain(
      "<xdr:to><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>17</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>",
    );
  });

  it("draws data labels, a legend for several series, and a value format", () => {
    const chart = chParts(
      chWorkbook([
        {
          ...CH_REVENUE,
          series: [
            { name: "Revenue", values: "B5:B7", color: "#F97316" },
            { name: "Growth", values: "C5:C7" },
          ],
          dataLabels: true,
          format: "$#,##0",
        },
      ]),
    )["xl/charts/chart1.xml"];
    expect(chart).toContain('<c:showVal val="1"/>');
    expect(chart).toContain("<c:legend>");
    expect(chart).toContain('<a:srgbClr val="F97316"/>');
    expect(chart).toContain('formatCode="$#,##0"');
  });

  it("makes line and pie charts, and a horizontal bar", () => {
    const line = chParts(chWorkbook([{ ...CH_REVENUE, type: "line" }]))["xl/charts/chart1.xml"];
    const pie = chParts(chWorkbook([{ ...CH_REVENUE, type: "pie" }]))["xl/charts/chart1.xml"];
    const bar = chParts(chWorkbook([{ ...CH_REVENUE, type: "bar" }]))["xl/charts/chart1.xml"];
    expect(line).toContain("<c:lineChart>");
    expect(pie).toContain("<c:pieChart>");
    expect(pie).not.toContain("<c:valAx>");
    expect(bar).toContain('<c:barDir val="bar"/>');
  });

  it("numbers charts across sheets and keeps comment rels first", () => {
    const schema = chWorkbook([CH_REVENUE, { ...CH_REVENUE, type: "line", position: { from: "E20", to: "L36" } }]);
    schema.sheets[0].cells.A1 = { value: "Title", comment: { text: "note" } };
    schema.sheets.push({ name: "Second", cells: { A1: 1, A2: 2 }, charts: [{ type: "column", series: [{ values: "A1:A2" }] }] });
    const parts = chParts(schema);

    for (const key of ["xl/charts/chart1.xml", "xl/charts/chart2.xml", "xl/charts/chart3.xml", "xl/drawings/drawing1.xml", "xl/drawings/drawing2.xml"]) {
      expect(parts).toHaveProperty([key]);
    }
    expect(parts["xl/worksheets/_rels/sheet1.xml.rels"]).toContain(
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>',
    );
    expect(parts["xl/worksheets/sheet1.xml"]).toContain('<drawing r:id="rId3"/><legacyDrawing r:id="rId1"/>');
    expect(parts["xl/drawings/_rels/drawing2.xml.rels"]).toContain('Target="../charts/chart3.xml"');
  });

  it("adds nothing when there are no charts", () => {
    const parts = chParts(chWorkbook([]));
    expect(Object.keys(parts).filter((p) => p.includes("chart") || p.includes("drawing"))).toEqual([]);
    expect(parts["xl/worksheets/sheet1.xml"]).not.toContain("<drawing");
  });

  it("reads a chart back as it was written", () => {
    const chart = {
      type: "column",
      title: "Q3 revenue by region",
      categories: "A5:A7",
      series: [
        { name: "Revenue", values: "B5:B7", color: "#F97316" },
        { name: "Growth", values: "C5:C7" },
      ],
      position: { from: "E2", to: "L18" },
      dataLabels: true,
      legend: true,
      format: "$#,##0",
    };
    const read = Agent.read(Agent.toBytes(chWorkbook([chart]))) as any;
    // toStrictEqual on a JSON round-trip also pins key ORDER, which PHP's toBe does.
    expect(JSON.stringify(read.sheets[0].charts)).toBe(JSON.stringify([chart]));
  });

  it("refuses a malformed chart with a hint", () => {
    const paths = Agent.validate(
      chWorkbook([
        { type: "donut", series: [{ values: "B5:B7" }] },
        { type: "column", series: [] },
        { type: "column", series: [{ values: "B5-B7" }], categories: "nope", position: { from: "L18", to: "E2" } },
      ]),
    ).map((e) => e.path);
    expect(paths).toContain("sheets[0].charts[0].type");
    expect(paths).toContain("sheets[0].charts[1].series");
    expect(paths).toContain("sheets[0].charts[2].series[0].values");
    expect(paths).toContain("sheets[0].charts[2].categories");
    expect(paths).toContain("sheets[0].charts[2].position");
  });

  it("publishes charts in the schema", () => {
    const sheet = (Agent.toolDefinition() as any).definitions.Sheet.properties;
    expect(sheet).toHaveProperty("charts");
    expect(JSON.stringify(sheet.charts)).toContain("categories");
  });
});
