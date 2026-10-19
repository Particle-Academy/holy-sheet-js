import { describe, expect, it } from "vitest";
import { Agent, unzipSync } from "../src";

/**
 * Print setup and row heights. Mirrors the PHP engine's PrintAndRowHeightsTest;
 * byte parity with PHP is the "printSetup" entry in parity.test.ts.
 */
type Any = Record<string, unknown>;

const schema = (extra: Any = {}): Any => ({ sheets: [{ name: "Q3", cells: { A1: "Title", B2: 42 }, ...extra }] });

function sheetXml(s: Any): string {
  const files = unzipSync(Agent.toBytes(s as never));
  return new TextDecoder().decode(files["xl/worksheets/sheet1.xml"]!);
}

describe("print setup and row heights", () => {
  it("fits to one page wide and picks the orientation", () => {
    const xml = sheetXml(schema({ print: { orientation: "landscape", fitToWidth: true } }));
    expect(xml).toContain('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
    expect(xml).toContain('<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>');
  });

  it("prints a footer with live page codes, escaping a literal ampersand", () => {
    const xml = sheetXml(schema({ print: { footer: { left: "Q3 Review & Forecast", right: "Page {page} of {pages}" } } }));
    expect(xml).toContain(
      "<headerFooter><oddFooter>&amp;LQ3 Review &amp;&amp; Forecast&amp;RPage &amp;P of &amp;N</oddFooter></headerFooter>",
    );
  });

  it("sets row heights in pixels, on rows with and without cells", () => {
    const xml = sheetXml(schema({ rowHeights: { "1": 32, "3": 20 } }));
    expect(xml).toContain('<row r="1" ht="24" customHeight="1">');
    expect(xml).toContain('<row r="3" ht="15" customHeight="1"/>');
  });

  it("adds nothing when neither is asked for", () => {
    const xml = sheetXml(schema());
    expect(xml).not.toContain("sheetPr");
    expect(xml).not.toContain("pageSetup");
    expect(xml).not.toContain("customHeight");
  });

  it("refuses a malformed print or row height", () => {
    const paths = Agent.validate(
      schema({ rowHeights: { "0": 20, "2": -1 }, print: { orientation: "sideways", fitToWidth: "yes", footer: { left: 7 } } }) as never,
    ).map((e: Any) => e.path);
    expect(paths).toEqual(
      expect.arrayContaining([
        "sheets[0].rowHeights.0",
        "sheets[0].rowHeights.2",
        "sheets[0].print.orientation",
        "sheets[0].print.fitToWidth",
        "sheets[0].print.footer.left",
      ]),
    );
  });
});

describe("fit to one page", () => {
  it("fits a short sheet onto one page, so a chart below the table is never cut", () => {
    const xml = sheetXml(schema({ print: { orientation: "portrait", fitToPage: true } }));
    expect(xml).toContain('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
    expect(xml).toContain('<pageSetup orientation="portrait" fitToWidth="1" fitToHeight="1"/>');

    const read = Agent.read(Agent.toBytes(schema({ print: { fitToPage: true, footer: { right: "Page {page} of {pages}" } } }) as never)) as Any;
    expect(JSON.stringify((read.sheets as Any[])[0]!.print)).toBe(JSON.stringify({ fitToPage: true, footer: { right: "Page {page} of {pages}" } }));

    expect(Agent.validate(schema({ print: { fitToPage: 1 } }) as never).map((e) => e.path)).toStrictEqual(["sheets[0].print.fitToPage"]);
    expect(JSON.stringify((Agent.toolDefinition() as Any as { definitions: { Sheet: { properties: Any } } }).definitions.Sheet.properties.print)).toContain("fitToPage");
  });
});
