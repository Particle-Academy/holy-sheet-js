import { describe, it, expect } from "vitest";
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import {
  Agent,
  ArchiveRefused,
  DEFAULT_ARCHIVE_LIMITS,
  DEFAULT_READ_LIMITS,
  ReadLimitExceeded,
  SheetNotFoundException,
  unzipSync,
  zipSync,
} from "../src";

/**
 * CC-fancy-0012 (sheet selection, the rows view) and CC-fancy-0013 (read
 * limits), driven by the PHP repo's tests/fixtures/read-options/cases.json:
 * the same bytes and the same hand-written expectations all three ports pass.
 * Found like the ODS fixtures (HOLY_SHEET_PHP_SRC, then the sibling checkout);
 * missing is an error, never a skip.
 */
const DIR = process.env.HOLY_SHEET_PHP_SRC
  ? join(process.env.HOLY_SHEET_PHP_SRC, "..", "tests", "fixtures", "read-options")
  : join(__dirname, "..", "..", "holy-sheet", "tests", "fixtures", "read-options");

if (!existsSync(join(DIR, "cases.json"))) {
  throw new Error(`read-option fixtures not found at ${DIR}; check out the PHP holy-sheet repo beside this one or set HOLY_SHEET_PHP_SRC.`);
}

type Any = any; // eslint-disable-line @typescript-eslint/no-explicit-any
const CASES: Any = JSON.parse(readFileSync(join(DIR, "cases.json"), "utf8"));
const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(join(DIR, name)));

function outcome(run: () => unknown): Any {
  try {
    return { value: run() };
  } catch (e) {
    // ReadLimitExceeded is an ArchiveRefused too. A refusal naming a limit
    // carries which one, its value and the count; every refusal a message.
    if (e instanceof ArchiveRefused) {
      const named = e.limit !== null ? { limit: e.limit, max: e.max, actual: e.actual } : {};
      return { refused: { reason: e.reason, ...named, message: e.message } };
    }
    if (e instanceof SheetNotFoundException) return { sheetNotFound: { sheet: e.sheet, available: e.available, message: e.message } };
    throw e;
  }
}

describe("read options (shared cases)", () => {
  it("has cases to run", () => {
    expect(CASES.read.length).toBeGreaterThan(20);
    expect(CASES.rows.length).toBeGreaterThan(4);
  });

  for (const c of CASES.read) {
    it(`read: ${c.name}`, () => {
      const got = outcome(() => Agent.read(fixture(c.file), c.options));
      if (c.refused) expect(got.refused).toEqual(c.refused);
      else if (c.sheetNotFound) expect(got.sheetNotFound).toEqual(c.sheetNotFound);
      else expect((got.value as Any).sheets.map((s: Any) => s.name)).toEqual(c.sheets);
    });
  }

  for (const c of CASES.rows) {
    it(`rows: ${c.name}`, () => {
      const got = outcome(() => Agent.rows(c.sheet, c.options));
      if (c.refused) expect(got.refused).toEqual(c.refused);
      else expect(got.value).toEqual(c.rows);
    });
  }

  it("pins the default limits the ports share", () => {
    expect({ ...DEFAULT_READ_LIMITS }).toEqual(CASES.defaultLimits);
    expect({ ...DEFAULT_ARCHIVE_LIMITS }).toEqual(CASES.defaultArchiveLimits);
  });
});

describe("read options (this port)", () => {
  it("is still an ArchiveRefused, so an existing catch keeps catching", () => {
    expect(() => Agent.read(fixture("two-sheets.xlsx"), { limits: { maxRows: 1 } })).toThrow(ArchiveRefused);
  });

  it("reads only the selected sheet, with its cells", () => {
    const schema: Any = Agent.read(fixture("two-sheets.xlsx"), { sheet: "Summary" });
    expect(schema.sheets).toHaveLength(1);
    expect(schema.sheets[0].cells).toEqual({ A1: { value: 1 }, B1: { value: 2 } });
  });

  it("takes the same options from a path", async () => {
    const path = join(DIR, "two-sheets.xlsx");
    const schema: Any = await Agent.describe(path, { sheet: "Data" });
    expect(schema.sheets.map((s: Any) => s.name)).toEqual(["Data"]);
    await expect(Agent.describe(path, { limits: { maxRows: 1 } })).rejects.toThrow(ReadLimitExceeded);
  });

  it("refuses an unknown read option and a sheet that is not a name", () => {
    const bytes = fixture("two-sheets.xlsx");
    expect(() => Agent.read(bytes, { sheets: "Data" } as Any)).toThrow("unknown read option: sheets");
    expect(() => Agent.read(bytes, { sheet: 1 } as Any)).toThrow("sheet must be a sheet name");
    expect(() => Agent.rows({ cells: {} }, { formulas: "cached" } as Any)).toThrow("formulas must be");
    expect(() => Agent.rows({ cells: {} }, { bogus: 1 } as Any)).toThrow("unknown rows option: bogus");
  });

  it("reads a writer round trip into rows", () => {
    const bytes = Agent.toBytes({
      sheets: [{ name: "T", columns: [{ header: "Name" }, { header: "Qty", type: "integer" }], rows: [["a", 1], ["b", 2]] }],
    });
    const sheet: Any = (Agent.read(bytes, { sheet: "T" }) as Any).sheets[0];
    expect(Agent.rows(sheet)).toEqual([["Name", "Qty"], ["a", 1], ["b", 2]]);
  });
});

describe("border styles (CC-fancy-0012)", () => {
  const stylesXml = (format: Any): string => {
    const parts = unzipSync(Agent.toBytes({ sheets: [{ name: "B", cells: { A1: { value: 1, format } } }] }));
    return new TextDecoder().decode(parts["xl/styles.xml"]!);
  };

  it("writes a bare colour as a thin border, exactly as before", () => {
    expect(stylesXml({ borderTop: "#FF0000" })).toContain(
      '<border><left/><right/><top style="thin"><color rgb="FFFF0000"/></top><bottom/><diagonal/></border>',
    );
  });

  for (const style of ["thin", "medium", "thick", "dashed", "dotted", "double"]) {
    it(`writes ${style} per side`, () => {
      expect(stylesXml({ borderBottom: { style, color: "#00AA00" }, borderLeft: "#000000" })).toContain(
        `<border><left style="thin"><color rgb="FF000000"/></left><right/><top/><bottom style="${style}"><color rgb="FF00AA00"/></bottom><diagonal/></border>`,
      );
    });
  }

  it("defaults an object border to thin and black", () => {
    expect(stylesXml({ borderRight: { style: "dashed" } })).toContain('<right style="dashed"><color rgb="FF000000"/></right>');
    expect(stylesXml({ borderRight: {} })).toContain('<right style="thin"><color rgb="FF000000"/></right>');
  });

  it("dedupes a thin object border with the bare colour it means", () => {
    const parts = unzipSync(
      Agent.toBytes({
        sheets: [{ name: "B", cells: { A1: { value: 1, format: { borderTop: "#FF0000" } }, A2: { value: 2, format: { borderTop: { style: "thin", color: "#FF0000" } } } } }],
      }),
    );
    expect(new TextDecoder().decode(parts["xl/styles.xml"]!)).toContain('<borders count="2">');
  });

  it("reads a border style back, and a thin one as the bare colour", () => {
    const schema: Any = Agent.read(
      Agent.toBytes({
        sheets: [
          {
            name: "B",
            cells: {
              A1: { value: 1, format: { borderTop: "#FF0000", borderBottom: { style: "double", color: "#0000FF" } } },
              A2: { value: 2, format: { borderLeft: { style: "thin", color: "#00FF00" } } },
            },
          },
        ],
      }),
    );
    expect(schema.sheets[0].cells.A1.format.borderTop).toBe("#FF0000");
    expect(schema.sheets[0].cells.A1.format.borderBottom).toEqual({ style: "double", color: "#0000FF" });
    expect(schema.sheets[0].cells.A2.format.borderLeft).toBe("#00FF00");
  });

  it("reads a style outside the six as the thin border it always read as", () => {
    const parts = unzipSync(Agent.toBytes({ sheets: [{ name: "B", cells: { A1: { value: 1, format: { borderTop: { style: "dashed", color: "#FF0000" } } } } }] }));
    const styles = new TextDecoder().decode(parts["xl/styles.xml"]!).replace('style="dashed"', 'style="mediumDashDot"');
    parts["xl/styles.xml"] = new TextEncoder().encode(styles);
    const schema: Any = Agent.read(zipSync(Object.entries(parts).map(([name, data]) => ({ name, data }))));
    expect(schema.sheets[0].cells.A1.format.borderTop).toBe("#FF0000");
  });

  it("round trips styled borders as the same workbook", () => {
    const schema = { sheets: [{ name: "B", cells: { A1: { value: 1, format: { borderTop: { style: "medium", color: "#123456" } } } } }] };
    expect(Agent.diff(schema, Agent.read(Agent.toBytes(schema)))).toEqual([]);
  });

  it("validates a border as a colour or {style, color}", () => {
    const errors = (border: unknown): Any[] =>
      Agent.validate({ sheets: [{ name: "B", cells: { A1: { value: 1, format: { borderTop: border } } } }] });
    expect(errors("#000000")).toEqual([]);
    expect(errors({ style: "thick", color: "#00ff00" })).toEqual([]);
    expect(errors({ style: "dotted" })).toEqual([]);
    const bad = errors({ style: "zigzag" });
    expect(bad).toHaveLength(1);
    expect(bad[0].path).toBe("sheets[0].cells.A1.format.borderTop.style");
    expect(bad[0].hint).toContain("thin, medium, thick, dashed, dotted, double");
    expect(errors({ style: "thin", color: "red" })[0].path).toBe("sheets[0].cells.A1.format.borderTop.color");
    expect(errors({ style: "thin", width: 2 })[0].path).toBe("sheets[0].cells.A1.format.borderTop.width");
    expect(errors(5)[0].path).toBe("sheets[0].cells.A1.format.borderTop");
    expect(errors(["thin"])[0].path).toBe("sheets[0].cells.A1.format.borderTop");
  });

  it("validates borders in the row form too", () => {
    const errors: Any[] = Agent.validate({ sheets: [{ name: "B", rows: [[{ value: 1, format: { borderLeft: { style: "wavy" } } }]] }] });
    expect(errors).toHaveLength(1);
    expect(errors[0].path).toBe("sheets[0].rows[0][0].format.borderLeft.style");
  });
});
