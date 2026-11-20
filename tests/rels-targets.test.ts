import { describe, it, expect } from "vitest";
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { Agent, MissingPartException } from "../src";
import { RelsParser } from "../src/reader/rels-parser";
import { zipSync } from "../src/zip";

/**
 * Relationship targets resolve per OPC (ECMA-376 Part 2): against the
 * directory of the part that owns the relationship, unless the target starts
 * with "/", which is package-absolute. openpyxl writes absolute targets
 * (`/xl/worksheets/sheet1.xml`); this reader used to turn that into
 * `xl/xl/worksheets/sheet1.xml`, find nothing and silently return zero sheets.
 * External targets are never parts, and a relationship whose part is missing
 * throws MissingPartException instead of being skipped.
 *
 * Driven by the PHP repo's tests/fixtures/rels-targets/cases.json (packages
 * assembled by hand from raw XML, expectations written by hand), the same
 * bytes all three ports pass. Found like the active-sheet fixtures; missing is
 * an error, never a skip. The tests at the bottom zip their own raw XML.
 */
const DIR = process.env.HOLY_SHEET_PHP_SRC
  ? join(process.env.HOLY_SHEET_PHP_SRC, "..", "tests", "fixtures", "rels-targets")
  : join(__dirname, "..", "..", "holy-sheet", "tests", "fixtures", "rels-targets");

if (!existsSync(join(DIR, "cases.json"))) {
  throw new Error(`rels-targets fixtures not found at ${DIR}; check out the PHP holy-sheet repo beside this one or set HOLY_SHEET_PHP_SRC.`);
}

type Any = any; // eslint-disable-line @typescript-eslint/no-explicit-any
const CASES: Any = JSON.parse(readFileSync(join(DIR, "cases.json"), "utf8"));
const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(join(DIR, name)));

function projection(schema: Any): Any[] {
  return (schema.sheets as Any[]).map((sheet) => {
    const cells: Any = {};
    const bold: string[] = [];
    const comments: Any = {};
    for (const [address, cell] of Object.entries<Any>(sheet.cells ?? {})) {
      cells[address] = cell.value ?? null;
      if (cell.format?.bold === true) bold.push(address);
      if (cell.comment !== undefined) comments[address] = cell.comment.text;
    }
    return { name: sheet.name, cells, bold, comments, charts: ((sheet.charts ?? []) as Any[]).map((c) => c.type) };
  });
}

function caught(fn: () => unknown): unknown {
  try {
    fn();
  } catch (e) {
    return e;
  }
  return undefined;
}

describe("relationship targets (shared cases)", () => {
  it("has cases to run", () => {
    expect(CASES.read.length).toBeGreaterThan(10);
    expect(CASES.raises.length).toBeGreaterThan(8);
  });

  for (const c of CASES.read) {
    it(`reads: ${c.name}`, () => {
      expect(projection(Agent.read(fixture(c.file), c.options))).toEqual(CASES.projections[c.sheets]);
    });
  }

  for (const c of CASES.raises) {
    it(`raises: ${c.name}`, () => {
      const e = caught(() => Agent.read(fixture(c.file), c.options));
      expect(e).toBeInstanceOf(MissingPartException);
      const err = e as MissingPartException;
      expect(err.part).toBe(c.part);
      expect(err.source).toBe(c.source);
      expect(err.message.startsWith("[holy-sheet] ")).toBe(true);
      if (c.part !== null) expect(err.message).toContain(c.part);
    });
  }
});

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const enc = new TextEncoder();

function rels(...list: [string, string, string][]): string {
  return `<Relationships xmlns="${PKG}">${list.map(([id, t, target]) => `<Relationship Id="${id}" Type="${REL}/${t}" Target="${target}"/>`).join("")}</Relationships>`;
}

function openpyxlLayout(sheetTarget: string, sheetPart: string): Uint8Array {
  const parts: Record<string, string> = {
    "[Content_Types].xml":
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>',
    "_rels/.rels": rels(["rId1", "officeDocument", "/xl/workbook.xml"]),
    "xl/workbook.xml": `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Only" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": rels(["rId1", "worksheet", sheetTarget]),
    [sheetPart]: `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>`,
  };
  return zipSync(Object.entries(parts).map(([name, body]) => ({ name, data: enc.encode(body) })));
}

describe("relationship targets (raw XML)", () => {
  it("reads a package-absolute worksheet target", () => {
    const schema: Any = Agent.read(openpyxlLayout("/xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"));
    expect(schema.sheets[0].cells.A1.value).toBe(42);
  });

  it("raises for an absolute target to nothing instead of reading zero sheets", () => {
    const e = caught(() => Agent.read(openpyxlLayout("/xl/worksheets/sheet1.xml", "xl/worksheets/elsewhere.xml")));
    expect(e).toBeInstanceOf(MissingPartException);
    expect((e as MissingPartException).part).toBe("xl/worksheets/sheet1.xml");
  });

  const resolutions: [string, string, string][] = [
    ["xl/workbook.xml", "worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"],
    ["xl/workbook.xml", "/xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"],
    ["xl/worksheets/sheet1.xml", "../comments1.xml", "xl/comments1.xml"],
    ["xl/worksheets/sheet1.xml", "/xl/comments/comment1.xml", "xl/comments/comment1.xml"],
    ["xl/worksheets/sheet1.xml", "./../drawings/./drawing1.xml", "xl/drawings/drawing1.xml"],
    // The package-level rels (_rels/.rels) belong to the package root, "".
    ["", "xl/workbook.xml", "xl/workbook.xml"],
    ["xl/workbook.xml", "../../../../x.xml", "x.xml"],
  ];
  for (const [source, target, expected] of resolutions) {
    it(`resolves ${target} from ${source || "the package root"}`, () => {
      expect(RelsParser.resolve(source, target)).toBe(expected);
    });
  }

  it("parses the target mode", () => {
    const parsed = RelsParser.parse(
      `<Relationships xmlns="${PKG}">` +
        `<Relationship Id="rId1" Type="${REL}/hyperlink" Target="https://example.com/" TargetMode="External"/>` +
        `<Relationship Id="rId2" Type="${REL}/comments" Target="../comments1.xml"/></Relationships>`,
    );
    expect(RelsParser.isExternal(parsed["rId1"]!)).toBe(true);
    expect(RelsParser.isExternal(parsed["rId2"]!)).toBe(false);
  });
});
