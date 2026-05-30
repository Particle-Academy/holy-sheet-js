import { describe, it, expect } from "vitest";
import { Agent, HolySheet, SchemaException, unzipSync } from "../src";

const dec = new TextDecoder();
const partText = (bytes: Uint8Array, name: string) => dec.decode(unzipSync(bytes)[name]!);

describe("writer", () => {
  it("emits a valid OOXML package with the expected parts", () => {
    const bytes = Agent.toBytes({ sheets: [{ name: "S1", cells: { A1: { value: "hi" } } }] });
    const files = unzipSync(bytes);
    expect(Object.keys(files)).toEqual(
      expect.arrayContaining([
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/styles.xml",
        "xl/worksheets/sheet1.xml",
        "docProps/core.xml",
        "docProps/app.xml",
      ]),
    );
    expect(partText(bytes, "xl/worksheets/sheet1.xml")).toContain("hi");
  });

  it("writes numbers, booleans, formulas, and serial dates", () => {
    const bytes = Agent.toBytes({
      sheets: [
        {
          name: "Data",
          cells: {
            A1: { value: 42 },
            A2: { value: 3.5 },
            A3: { value: true },
            A4: { formula: "SUM(A1:A2)" },
            A5: { value: "2024-01-15", format: { displayFormat: "date" } },
          },
        },
      ],
    });
    const xml = partText(bytes, "xl/worksheets/sheet1.xml");
    expect(xml).toContain('<c r="A1"><v>42</v></c>');
    expect(xml).toContain("<v>3.5</v>");
    expect(xml).toContain('t="b"><v>1</v>');
    expect(xml).toContain("<f>SUM(A1:A2)</f>");
    // 2024-01-15 → Excel serial 45306
    expect(xml).toContain("<v>45306</v>");
  });

  it("dedups styles and applies themes/totals", () => {
    const bytes = Agent.toBytes({
      sheets: [
        {
          name: "Sales",
          columns: [
            { header: "Region", type: "string" },
            { header: "Revenue", type: "currency", currency: "USD" },
          ],
          rows: [
            ["North", 12000],
            ["South", 9800.5],
          ],
          totals: { Revenue: "sum" },
          theme: "default",
        },
      ],
    });
    const sheet = partText(bytes, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain("<f>SUM(B2:B3)</f>");
    // currency numFmt, XML-escaped (matches PHP htmlspecialchars)
    expect(partText(bytes, "xl/styles.xml")).toContain("&quot;$&quot;#,##0.00");
  });
});

describe("validator", () => {
  it("flags missing sheets", () => {
    const errors = Agent.validate({} as never);
    expect(errors).toHaveLength(1);
    expect(errors[0]!.path).toBe("sheets");
  });

  it("flags empty sheets array", () => {
    expect(Agent.validate({ sheets: [] })[0]!.expected).toBe("non-empty array");
  });

  it("flags a sheet without name or data", () => {
    const errors = Agent.validate({ sheets: [{}] as never });
    expect(errors.some((e) => e.path === "sheets[0].name")).toBe(true);
  });

  it("toBytes throws SchemaException on invalid input", () => {
    expect(() => Agent.toBytes({ sheets: [] })).toThrow(SchemaException);
    try {
      Agent.toBytes({ sheets: [] });
    } catch (e) {
      expect((e as SchemaException).getErrors()).toHaveLength(1);
    }
  });
});

describe("repairer", () => {
  it("renames singular sheet → sheets", () => {
    const { schema, repairs } = Agent.validateAndRepair({
      sheet: { name: "S", cells: { A1: { value: 1 } } },
    } as never);
    expect((schema as any).sheets).toHaveLength(1);
    expect(repairs.some((r) => r.includes("'sheet' → 'sheets'"))).toBe(true);
  });

  it("renames row → rows and converts object-as-list", () => {
    const { schema, repairs } = Agent.validateAndRepair({
      sheets: [{ name: "S", columns: [{ header: "A" }], row: { 0: [1], 1: [2] } }],
    } as never);
    expect((schema as any).sheets[0].rows).toEqual([[1], [2]]);
    expect(repairs.length).toBeGreaterThanOrEqual(2);
  });

  it("normalizes unknown theme and coerces stringified numerics", () => {
    const { schema } = Agent.validateAndRepair({
      sheets: [{ name: "S", theme: "neon", columns: [{ header: "N", type: "number" }], rows: [["12.5"]] }],
    } as never);
    expect((schema as any).sheets[0].theme).toBe("default");
    expect((schema as any).sheets[0].rows[0][0]).toBe(12.5);
  });

  it("infers date type from ISO values", () => {
    const { schema } = Agent.validateAndRepair({
      sheets: [{ name: "S", columns: [{ header: "When" }], rows: [["2024-01-01"], ["2024-02-01"]] }],
    } as never);
    expect((schema as any).sheets[0].columns[0].type).toBe("date");
  });
});

describe("formula linter", () => {
  const lint = (cells: Record<string, unknown>) =>
    Agent.lint({ sheets: [{ name: "S", cells }] });

  it("detects #DIV/0!", () => {
    const issues = lint({ A1: { value: 10 }, A2: { formula: "A1/0" } });
    expect(issues[0]!.error).toBe("#DIV/0!");
  });

  it("detects #VALUE! with an off-by-one hint", () => {
    const issues = lint({ A1: { value: "Revenue" }, A2: { value: 100 }, B1: { formula: "A1+1" } });
    expect(issues[0]!.error).toBe("#VALUE!");
    expect(issues[0]!.hint).toContain("string");
  });

  it("detects #NAME? for unknown functions", () => {
    expect(lint({ A1: { formula: "FOOBAR(1)" } })[0]!.error).toBe("#NAME?");
  });

  it("detects circular references", () => {
    const issues = lint({ A1: { formula: "B1" }, B1: { formula: "A1" } });
    expect(issues.some((i) => i.error === "#CIRC!")).toBe(true);
  });

  it("evaluates SUM/AVERAGE/IF cleanly", () => {
    const issues = lint({
      A1: { value: 10 },
      A2: { value: 20 },
      B1: { formula: "SUM(A1:A2)" },
      B2: { formula: "AVERAGE(A1:A2)" },
      B3: { formula: 'IF(A1>5,"big","small")' },
    });
    expect(issues).toHaveLength(0);
  });
});

describe("builders", () => {
  it("fromArray infers column types from headers + values", () => {
    const schema = Agent.fromArray(
      [
        ["North", 12000, 0.12],
        ["South", 9800, 0.08],
      ],
      ["Region", "Revenue", "Growth Rate"],
    );
    const cols = schema.sheets[0]!.columns!;
    expect(cols[0]!.type).toBe("string");
    expect(cols[1]!.type).toBe("currency");
    expect(cols[2]!.type).toBe("percent");
  });

  it("fromCsv parses content and infers headers", () => {
    const schema = Agent.fromCsv("Name,Age\nAlice,30\nBob,42");
    expect(schema.sheets[0]!.columns!.map((c) => c.header)).toEqual(["Name", "Age"]);
    expect(schema.sheets[0]!.rows).toEqual([
      ["Alice", 30],
      ["Bob", 42],
    ]);
  });

  it("fromCsv handles quoted fields with embedded commas + newlines", () => {
    const schema = Agent.fromCsv('A,B\n"x,y","line1\nline2"');
    expect(schema.sheets[0]!.rows![0]).toEqual(["x,y", "line1\nline2"]);
  });
});

describe("round-trip (read)", () => {
  it("toBytes → read recovers cell values", () => {
    const schema = {
      sheets: [
        {
          name: "Data",
          cells: {
            A1: { value: "Region" },
            A2: { value: "North" },
            B1: { value: "Revenue" },
            B2: { value: 12000 },
            B3: { value: 9800.5 },
          },
        },
      ],
    };
    const back = Agent.read(Agent.toBytes(schema)) as any;
    expect(back.sheets[0].name).toBe("Data");
    expect(back.sheets[0].cells.A2.value).toBe("North");
    expect(back.sheets[0].cells.B2.value).toBe(12000);
    expect(back.sheets[0].cells.B3.value).toBe(9800.5);
  });

  it("round-trips a date column to an ISO string", () => {
    const bytes = Agent.toBytes({
      sheets: [{ name: "S", cells: { A1: { value: "2024-03-15", format: { displayFormat: "date" } } } }],
    });
    const back = Agent.read(bytes) as any;
    expect(back.sheets[0].cells.A1.value).toBe("2024-03-15");
  });

  it("round-trips frozen panes + merged regions + comments", () => {
    const bytes = Agent.toBytes({
      sheets: [
        {
          name: "S",
          cells: {
            A1: { value: "x", comment: { text: "note here", author: "Ada" } },
            B1: { value: "y" },
          },
          mergedRegions: [{ start: "A1", end: "B1" }],
          frozenRows: 1,
          frozenCols: 1,
        },
      ],
    });
    const back = Agent.read(bytes) as any;
    expect(back.sheets[0].frozenRows).toBe(1);
    expect(back.sheets[0].frozenCols).toBe(1);
    expect(back.sheets[0].mergedRegions).toEqual([{ start: "A1", end: "B1" }]);
    expect(back.sheets[0].cells.A1.comment.text).toBe("note here");
    expect(back.sheets[0].cells.A1.comment.author).toBe("Ada");
  });
});

describe("surface", () => {
  it("HolySheet instance mirrors Agent", () => {
    const hs = new HolySheet();
    expect(hs.validate({ sheets: [] })).toHaveLength(1);
    expect(typeof hs.getVersion()).toBe("string");
    expect(HolySheet.version()).toMatch(/^\d+\.\d+\.\d+$/);
  });

  it("toolDefinition returns the JSON schema", () => {
    const def = Agent.toolDefinition();
    expect((def as any).required).toContain("sheets");
  });

  it("write() + describe() work on disk (Node)", async () => {
    const os = await import("node:os");
    const path = await import("node:path");
    const fs = await import("node:fs");
    const file = path.join(os.tmpdir(), `holy-sheet-test-${process.pid}.xlsx`);
    const result = await Agent.write({ sheets: [{ name: "S", cells: { A1: { value: 7 } } }] }, file);
    expect(result.sheets).toBe(1);
    expect(result.bytes).toBeGreaterThan(0);
    const described = (await Agent.describe(file)) as any;
    expect(described.sheets[0].cells.A1.value).toBe(7);
    fs.unlinkSync(file);
  });
});
