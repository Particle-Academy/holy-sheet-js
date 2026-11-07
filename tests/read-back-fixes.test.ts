import { describe, expect, it } from "vitest";
import { Agent, unzipSync, zipSync } from "../src/index";

/**
 * Four read/write defects a consumer (LiftPoint) hit. Mirrors PHP
 * tests/Unit/ReadBackFixesTest.php case for case:
 *
 *  - the 1904 date system (`<workbookPr date1904="1"/>`) was ignored, so every
 *    date in such a workbook read 1462 days early;
 *  - a rich inline string (`<is><r><t>..</t></r></is>`) read as null;
 *  - a date-formatted formula's cached value came back as a raw serial;
 *  - a column's custom `format` code was ignored by the normalizer.
 *
 * Fixtures are built here, from the writer's own output with the parts a
 * producer like Excel writes patched in, so nothing depends on a binary file.
 */
type Parts = Record<string, Uint8Array>;
const enc = new TextEncoder();
const dec = new TextDecoder();

const zipOf = (parts: Parts): Uint8Array => zipSync(Object.entries(parts).map(([name, data]) => ({ name, data })));
const partsOf = (schema: unknown): Parts => unzipSync(Agent.toBytes(schema as never));
const text = (parts: Parts, part: string): string => dec.decode(parts[part]!);
const put = (parts: Parts, part: string, xml: string): void => {
  parts[part] = enc.encode(xml);
};

/** A1 a date, A2 a datetime, A3 a date formula, A4 a plain number formula, A5 an inline string. */
function baseParts(): Parts {
  return partsOf({
    sheets: [
      {
        name: "S",
        cells: {
          A1: { value: "2026-11-07", format: { displayFormat: "date" } },
          A2: { value: "2026-11-07T12:00:00Z", format: { displayFormat: "datetime" } },
          A3: { formula: "A1+1", computedValue: 46334, format: { displayFormat: "date" } },
          A4: { formula: "1+1", computedValue: 2 },
          A5: "plain",
        },
      },
    ],
  });
}

function withDate1904(parts: Parts, flag: string): Parts {
  put(parts, "xl/workbook.xml", text(parts, "xl/workbook.xml").replace("<sheets>", `<workbookPr date1904="${flag}"/><sheets>`));
  return parts;
}

/** The same workbook in the 1904 date system: the flag, and every serial 1462 days smaller. */
function parts1904(flag = "1"): Parts {
  const parts = withDate1904(baseParts(), flag);
  put(
    parts,
    "xl/worksheets/sheet1.xml",
    text(parts, "xl/worksheets/sheet1.xml")
      .replace("<v>46333</v>", "<v>44871</v>")
      .replace("<v>46333.5</v>", "<v>44871.5</v>")
      .replace("<v>46334</v>", "<v>44872</v>"),
  );
  return parts;
}

const cellsOf = (parts: Parts): any => (Agent.read(zipOf(parts)) as any).sheets[0].cells;

describe("the 1904 date system", () => {
  it("fixes the fixture: the 1900-system serials are the ones the 1904 copy rewrites", () => {
    const sheet = text(baseParts(), "xl/worksheets/sheet1.xml");
    expect(sheet).toContain("<v>46333</v>");
    expect(sheet).toContain("<v>46333.5</v>");
    expect(sheet).toContain("<v>46334</v>");
  });

  it.each(["1", "true"])("reads a date from a workbook that declares date1904=%s", (flag) => {
    const cells = cellsOf(parts1904(flag));
    expect(cells.A1.value).toBe("2026-11-07");
    expect(cells.A2.value).toBe("2026-11-07T12:00:00Z");
  });

  it.each(["0", "false", "absent"])("reads the 1900 system when date1904 is %s", (flag) => {
    const cells = cellsOf(flag === "absent" ? baseParts() : withDate1904(baseParts(), flag));
    expect(cells.A1.value).toBe("2026-11-07");
    expect(cells.A2.value).toBe("2026-11-07T12:00:00Z");
  });

  it("converts the cached value of a date formula with the same epoch", () => {
    expect(cellsOf(parts1904()).A3.computedValue).toBe("2026-11-08");
  });

  it("writes a read 1904 workbook back as the same dates in the 1900 system", () => {
    const again = partsOf(Agent.read(zipOf(parts1904())));
    expect(text(again, "xl/workbook.xml")).not.toContain("date1904");
    expect(text(again, "xl/worksheets/sheet1.xml")).toContain('<c r="A1" s="1"><v>46333</v></c>');
  });
});

describe("the cached value of a date-formatted formula", () => {
  it("reads as the ISO date the cell displays, not as a serial", () => {
    const cells = cellsOf(baseParts());
    expect(cells.A3.formula).toBe("A1+1");
    expect(cells.A3.computedValue).toBe("2026-11-08");
  });

  it("reads as an ISO datetime for a datetime format", () => {
    const parts = partsOf({
      sheets: [{ name: "S", cells: { A1: { formula: "NOW()", computedValue: 46333.25, format: { displayFormat: "datetime" } } } }],
    });
    expect(cellsOf(parts).A1.computedValue).toBe("2026-11-07T06:00:00Z");
  });

  it("stays a number when the formula is not date-formatted", () => {
    expect(cellsOf(baseParts()).A4.computedValue).toBe(2);
  });

  it("is written back as the serial, so describe() output round-trips", () => {
    const parts = partsOf({
      sheets: [{ name: "S", cells: { A1: { formula: "A2+1", computedValue: "2026-11-08", format: { displayFormat: "date" } } } }],
    });
    expect(text(parts, "xl/worksheets/sheet1.xml")).toContain('<c r="A1" s="1"><f>A2+1</f><v>46334</v></c>');
  });

  it("survives a read and a write unchanged", () => {
    const again = partsOf(Agent.read(zipOf(baseParts())));
    expect(text(again, "xl/worksheets/sheet1.xml")).toContain("<f>A1+1</f><v>46334</v>");
    expect(cellsOf(again).A3.computedValue).toBe("2026-11-08");
  });
});

describe("a rich inline string", () => {
  it("reads as the text of its runs, as a rich shared string does", () => {
    const parts = baseParts();
    const rich = '<c r="A5" t="inlineStr"><is><r><t xml:space="preserve">Hello </t></r><r><rPr><b/></rPr><t>world</t></r></is></c>';
    const sheet = text(parts, "xl/worksheets/sheet1.xml");
    const patched = sheet.replace(/<c r="A5"[^>]*><is>.*?<\/is><\/c>/, rich);
    expect(patched).not.toBe(sheet);
    put(parts, "xl/worksheets/sheet1.xml", patched);

    expect(cellsOf(parts).A5.value).toBe("Hello world");
  });

  it("still reads a plain inline string", () => {
    expect(cellsOf(baseParts()).A5.value).toBe("plain");
  });
});

describe("a column format code", () => {
  const xmlAttr = (s: string): string =>
    s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  const reEscape = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

  /** [the numFmt id written for code, the numFmt id of the xf cell ref uses]. */
  function numFmtOf(parts: Parts, code: string, ref: string): [string | null, string | null] {
    const styles = text(parts, "xl/styles.xml");
    const fmt = new RegExp(`<numFmt numFmtId="(\\d+)" formatCode="${reEscape(xmlAttr(code))}"/>`).exec(styles);
    const cell = new RegExp(`<c r="${ref}" s="(\\d+)"`).exec(text(parts, "xl/worksheets/sheet1.xml"));
    const xfs = /<cellXfs[^>]*>(.*)<\/cellXfs>/.exec(styles)?.[1] ?? "";
    const ids = [...xfs.matchAll(/<xf [^>]*numFmtId="(\d+)"/g)].map((m) => m[1]!);
    return [fmt?.[1] ?? null, cell ? ids[Number(cell[1])] ?? null : null];
  }

  it("is written as a custom numFmt and applied to the column", () => {
    const code = '#,##0.000" kg"';
    const parts = partsOf({
      sheets: [{ name: "S", columns: [{ header: "Weight", type: "number", format: code }], rows: [[1.5], [2.25]] }],
    });
    const [written, a2] = numFmtOf(parts, code, "A2");
    const [, a3] = numFmtOf(parts, code, "A3");

    expect(written).not.toBeNull();
    expect(a2).toBe(written);
    expect(a3).toBe(written);
  });

  it("overrides the type-derived code but keeps the type deciding the value", () => {
    const parts = partsOf({
      sheets: [{ name: "S", columns: [{ header: "When", type: "date", format: "dd/mm/yyyy" }], rows: [["2026-11-07"]] }],
    });
    const [written, a2] = numFmtOf(parts, "dd/mm/yyyy", "A2");

    expect(a2).not.toBeNull();
    expect(a2).toBe(written);
    expect(text(parts, "xl/styles.xml")).not.toContain("yyyy-mm-dd");
    expect(text(parts, "xl/worksheets/sheet1.xml")).toContain("<v>46333</v>");
  });

  it("applies to an auto column too", () => {
    const parts = partsOf({ sheets: [{ name: "S", columns: [{ header: "Ratio", format: "0.0000" }], rows: [[0.5]] }] });
    const [written, a2] = numFmtOf(parts, "0.0000", "A2");

    expect(written).not.toBeNull();
    expect(a2).toBe(written);
  });

  it("carries onto the totals row", () => {
    const parts = partsOf({
      sheets: [
        {
          name: "S",
          columns: [{ header: "Weight", type: "number", format: "0.000" }],
          rows: [[1.5], [2.5]],
          totals: { Weight: "sum" },
        },
      ],
    });
    const [written, total] = numFmtOf(parts, "0.000", "A4");

    expect(total).not.toBeNull();
    expect(total).toBe(written);
  });

  it("gives way to a displayFormat a cell sets itself", () => {
    const parts = partsOf({
      sheets: [
        {
          name: "S",
          columns: [{ header: "Weight", type: "number", format: "0.000" }],
          rows: [[{ value: 0.5, format: { displayFormat: "percentage" } }]],
        },
      ],
    });

    expect(text(parts, "xl/styles.xml")).not.toContain('formatCode="0.000"');
    expect(text(parts, "xl/styles.xml")).toContain('formatCode="0.0%"');
  });

  it.each([
    ["empty", ""],
    ["a number", 12],
    ["too long", "0".repeat(256)],
  ])("is refused by the validator when it cannot be written: %s", (_name, format) => {
    const errors = Agent.validate({ sheets: [{ name: "S", columns: [{ header: "X", format }], rows: [[1]] }] });

    expect(errors.map((e) => e.path)).toContain("sheets[0].columns[0].format");
  });

  it("is accepted by the validator when it is a code", () => {
    expect(Agent.validate({ sheets: [{ name: "S", columns: [{ header: "X", format: "0.00" }], rows: [[1]] }] })).toEqual([]);
  });

  it("reports the error in PHP's words", () => {
    const errors = Agent.validate({ sheets: [{ name: "S", columns: [{ header: "X", format: "" }], rows: [[1]] }] });
    const error = errors.find((e) => e.path === "sheets[0].columns[0].format")!;

    expect(error.expected).toBe("a non-empty number-format code of at most 255 bytes");
    expect(error.got).toBe("string");
    expect(error.hint).toBe(
      'A custom Excel number-format code applied to every cell in the column, like "0.000" or "dd/mm/yyyy". It overrides the format the column type would give.',
    );
  });
});
