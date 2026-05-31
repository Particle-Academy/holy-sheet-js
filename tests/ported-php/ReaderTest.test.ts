import { describe, it, expect } from "vitest";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Agent, zipSync } from "../../src";

// Ported from PHP tests/Unit/ReaderTest.php. PHP wrote a temp file then called
// Agent::describe(path). The TS Agent.describe is async + Node-only and behaves
// identically; we use it for the file cases. The shared-strings case builds a
// raw zip via zipSync and feeds the bytes through Agent.read (same reader).

const dir = mkdtempSync(join(tmpdir(), "holy-reader-"));
const enc = new TextEncoder();

async function writeTmp(schema: unknown, file: string): Promise<string> {
  const path = join(dir, file);
  await Agent.write(schema, path);
  return path;
}

describe("reader (ported PHP ReaderTest)", () => {
  it("describes a simple workbook back to schema", async () => {
    const path = await writeTmp(
      {
        sheets: [
          {
            name: "Q1",
            columns: [
              { header: "Region", type: "string" },
              { header: "Revenue", type: "currency", currency: "USD" },
            ],
            rows: [
              ["NA", 100],
              ["EU", 200],
            ],
          },
        ],
      },
      "q1.xlsx",
    );

    const out = (await Agent.describe(path)) as any;
    expect(out).toHaveProperty("sheets");
    expect(out.sheets[0].name).toBe("Q1");
    expect(out.sheets[0].cells).toHaveProperty("A1");
    expect(out.sheets[0].cells.A1.value).toBe("Region");
    expect(out.sheets[0].cells.B2.value).toBe(100);
  });

  it("returns not_found for a missing path", async () => {
    const out = await Agent.describe("/nope/missing.xlsx");
    expect(out).toEqual({ error: "not_found", path: "/nope/missing.xlsx" });
  });

  it("round-trips merged regions", async () => {
    const path = await writeTmp(
      { sheets: [{ name: "M", rows: [["a", "b"]], mergedRegions: [{ start: "A1", end: "B1" }] }] },
      "merge.xlsx",
    );
    const out = (await Agent.describe(path)) as any;
    expect(out.sheets[0].mergedRegions[0]).toEqual({ start: "A1", end: "B1" });
  });

  it("round-trips frozen panes", async () => {
    const path = await writeTmp(
      { sheets: [{ name: "F", rows: [["x"]], frozenRows: 1, frozenCols: 2 }] },
      "frozen.xlsx",
    );
    const out = (await Agent.describe(path)) as any;
    expect(out.sheets[0].frozenRows).toBe(1);
    expect(out.sheets[0].frozenCols).toBe(2);
  });

  it("round-trips formulas with cached values", async () => {
    const path = await writeTmp(
      { sheets: [{ name: "C", rows: [[1, 2, { formula: "A1+B1", computedValue: 3 }]] }] },
      "cached.xlsx",
    );
    const out = (await Agent.describe(path)) as any;
    expect(out.sheets[0].cells.C1.formula).toBe("A1+B1");
    expect(out.sheets[0].cells.C1.computedValue).toBe(3);
  });

  it("round-trips comments", async () => {
    const path = await writeTmp(
      {
        sheets: [{ name: "N", cells: { A1: { value: "hi", comment: { text: "note", author: "me" } } } }],
      },
      "notes.xlsx",
    );
    const out = (await Agent.describe(path)) as any;
    expect(out.sheets[0].cells.A1.comment.text).toBe("note");
  });

  it("resolves shared-string references from xl/sharedStrings.xml", () => {
    const file = (name: string, xml: string) => ({ name, data: enc.encode(xml) });
    const bytes = zipSync([
      file(
        "[Content_Types].xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
      ),
      file(
        "_rels/.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
      ),
      file(
        "xl/_rels/workbook.xml.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
      ),
      file(
        "xl/workbook.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
      ),
      file(
        "xl/sharedStrings.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
<si><t>Revenue Category</t></si>
<si><t>Subscription</t></si>
<si><r><t>Rich </t></r><r><t>String</t></r></si>
</sst>`,
      ),
      file(
        "xl/worksheets/sheet1.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
<row r="2"><c r="A2" t="s"><v>2</v></c></row>
</sheetData>
</worksheet>`,
      ),
    ]);

    const out = Agent.read(bytes) as any;
    expect(out.sheets[0].cells.A1.value).toBe("Revenue Category");
    expect(out.sheets[0].cells.B1.value).toBe("Subscription");
    expect(out.sheets[0].cells.A2.value).toBe("Rich String");
  });
});
