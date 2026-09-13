import { describe, expect, it } from "vitest";
import { Agent, UnsupportedFormatException } from "../../src";
import { odsFixture, odsFixturePath, odsPackage } from "../ods-fixtures";

// Ported from PHP tests/Unit/OdsReaderTest.php, on the same fixture bytes (loaded
// from the PHP repo). What that file says about each fixture applies here.
//
// PHP's `toBe` on arrays checks key ORDER as well as values; `toEqual` does not.
// `same()` compares JSON so a port that emits `format` keys in a different
// order fails here as it would in PHP.
//
// edge.ods is asserted cell by cell in PHP and diffed against PHP wholesale in
// tests/ods-reader-parity.test.ts; the traps a port is most likely to fall into
// on it are repeated below so they fail here even without PHP on the path.

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

function same(actual: unknown, expected: unknown): void {
  expect(JSON.stringify(actual)).toBe(JSON.stringify(expected));
}

function read(name: string): Any {
  return Agent.read(odsFixture(name));
}

/** The xlsx reader reports `displayFormat: auto` everywhere; ods invents none. See the PHP test. */
function withoutAuto(schema: Any): Any {
  for (const sheet of schema.sheets) {
    for (const cell of Object.values(sheet.cells) as Any[]) {
      if (cell.format?.displayFormat === "auto") {
        delete cell.format.displayFormat;
        if (Object.keys(cell.format).length === 0) delete cell.format;
      }
    }
  }
  return schema;
}

describe("ods reader (ported PHP OdsReaderTest)", () => {
  describe("the xlsx-converted fixture", () => {
    it("reads every value type from a LibreOffice-written ods", () => {
      const cells = read("workbook.ods").sheets[0].cells;

      same(cells.A2, { value: 100, comment: { text: "An integer", author: "Fixture" } });
      same(cells.B2, { value: 9800.5 });
      same(cells.C2, { value: -42 });
      same(cells.A3, { value: 0.25, format: { displayFormat: "percentage", decimals: 1 } });
      same(cells.B3, { value: 1234.5, format: { displayFormat: "currency", decimals: 2, currency: "USD" } });
      same(cells.C3, { value: 99, format: { displayFormat: "currency", decimals: 0, currency: "EUR" } });
      same(cells.A4, { value: "2024-03-15", format: { displayFormat: "date" } });
      same(cells.B4, { value: "2024-03-15T10:30:00Z", format: { displayFormat: "datetime" } });
      same(cells.C4, { value: 1234.5678, format: { displayFormat: "number", decimals: 2 } });
      same(cells.A5, { value: true });
      same(cells.B5, { value: false });
      same(cells.A6, { value: "first line\nsecond line" });
      same(cells.B6, { value: "spaced   out" });
      same(cells.C6, { value: "Fish & Chips <tasty>" });
    });

    it("reads cell formatting from automatic styles, parents and column defaults", () => {
      const sheets = read("workbook.ods").sheets;

      same(sheets[0].cells.A1, { value: "Kind", format: { bold: true } });
      same(sheets[0].cells.B1, { value: "Value", format: { bold: true, italic: true, textAlign: "center" } });
      same(sheets[0].cells.C1, {
        value: "Styled",
        format: { color: "#1D4ED8", backgroundColor: "#FEF3C7", fontSize: 14, borderBottom: "#111827" },
      });
      same(sheets[2].cells.A1, { value: "Across", format: { bold: true } });
    });

    it("translates OpenFormula into the A1 syntax the xlsx reader returns", () => {
      const cells = read("workbook.ods").sheets[0].cells;

      same(cells.A7, { value: null, formula: "SUM(A2:C2)", computedValue: 9858.5 });
      same(cells.B7, { value: null, formula: "'Other Sheet'!A1*2", computedValue: 42 });
      same(cells.C7, { value: null, formula: 'IF(A5,LEN("a;b"),0)', computedValue: 3 });
      same(cells.D7, { value: null, formula: "ROUND(B2/4,1)", computedValue: 2450.1 });
    });

    it("expands repeated cells and rows without inventing empty ones", () => {
      const cells = read("workbook.ods").sheets[1].cells;

      same(Object.keys(cells), ["A1", "B1", "C1", "D1", "E1", "A2", "B2", "A3", "B3", "A4", "B4", "H20"]);
      same(cells.E1, { value: 7 });
      same(cells.H20, { value: "far" });
    });

    it("reads merges, sheet names, an empty sheet and document metadata", () => {
      const schema = read("workbook.ods");

      same(
        schema.sheets.map((s: Any) => s.name),
        ["Types", "Repeats", "Merged", "Other Sheet", "Empty"],
      );
      same(schema.sheets[2].mergedRegions, [
        { start: "A1", end: "C1" },
        { start: "A2", end: "A3" },
      ]);
      same(schema.sheets[2].cells.B3, { value: 2 });
      same(schema.sheets[4], { name: "Empty", cells: {} });
      same(schema.meta, { creator: "Holy Sheet ODS fixture", created: "2026-09-13T12:00:00Z" });
    });

    it("describes the same workbook the same way whether it was saved as xlsx or ods", () => {
      const fromXlsx = withoutAuto(read("workbook.xlsx"));
      const fromOds = read("workbook.ods");

      // Frozen panes are view settings a headless conversion does not write.
      delete fromXlsx.sheets[0].frozenRows;
      delete fromXlsx.sheets[0].frozenCols;

      same(fromOds, fromXlsx);
    });

    it("returns a schema that writes straight back out", () => {
      const schema = read("workbook.ods");
      // Leave out the empty sheet, as the PHP test does (a pre-existing
      // read/write contract defect there, not an ods one).
      schema.sheets.splice(4, 1);

      same(Agent.validate(schema), []);
      const back = withoutAuto(Agent.read(Agent.toBytes(schema)));
      same(back.sheets[0].cells.B3, schema.sheets[0].cells.B3);
    });
  });

  describe("the hand-authored native fixture", () => {
    it("reads native OpenDocument values: time, currency, runs, links, tabs and breaks", () => {
      const schema = read("native.ods");
      const cells = schema.sheets[0].cells;

      same(cells.A1, {
        value: "Hello bold and a link",
        format: {
          bold: true,
          italic: true,
          color: "#F8FAFC",
          backgroundColor: "#0F172A",
          fontSize: 16,
          borderTop: "#334155",
          borderRight: "#334155",
          borderBottom: "#334155",
          borderLeft: "#334155",
        },
      });
      same(cells.B1, { value: "1899-12-30T10:30:00Z", format: { displayFormat: "datetime" } });
      same(cells.C1, { value: 19.99, format: { textAlign: "right", displayFormat: "currency", decimals: 2, currency: "EUR" } });
      same(cells.D1, { value: 0.125, format: { displayFormat: "percentage", decimals: 2 } });
      same(cells.A2, { value: false });
      same(cells.B2, { value: "1999-12-31", format: { displayFormat: "date" } });
      same(cells.C2, { value: "a\tb\nc", comment: { text: "First note paragraph\nSecond note paragraph" } });
      same(cells.D2, { value: 1500 });
      same(schema.meta, { creator: "Holy Sheet native ODS fixture", created: "2026-09-13T08:15:00Z" });
    });

    it("translates sheet references, absolute cells, quoted strings and inline arrays", () => {
      const cells = read("native.ods").sheets[0].cells;

      same(cells.A3, { value: null, formula: "SUM(Data!A1:A3)", computedValue: 6 });
      same(cells.B3, { value: null, formula: "$D$2*2", computedValue: 3000 });
      same(cells.C3, { value: null, formula: 'CONCATENATE("x;y","""q""")', computedValue: 'x;y"q"' });
      same(cells.D3, { value: null, formula: "SUMPRODUCT({1,2;3,4},{1,1;1,1})", computedValue: 10 });
    });
  });

  describe("edge.ods: the traps a port falls into", () => {
    it("rounds half a second away from zero, and applies a zone offset", () => {
      const cells = read("edge.ods").sheets[4].cells;

      same(cells.D1, { value: "2025-01-01T00:00:00Z", format: { displayFormat: "datetime" } });
      same(cells.E1, { value: "1899-12-30T00:00:03Z", format: { displayFormat: "datetime" } });
      same(cells.B1, { value: "2024-03-10T06:30:00Z", format: { displayFormat: "datetime" } });
    });

    it("does not let a graphic style named Default, or a style loop, leak into a cell", () => {
      const cells = read("edge.ods").sheets[0].cells;

      same(cells.B3, { value: "unknown style" });
      same(cells.M1, { value: "loop", format: { bold: true, italic: true } });
      same(cells.N1, { value: -5, format: { displayFormat: "currency", decimals: 2, currency: "USD" } });
    });

    it("matches attributes by namespace URI, not by local name or prefix", () => {
      // office: under the prefix "o", beside a decoy namespace reusing the same
      // local names ("value-type", "value"). By URI this is 42; read by local
      // name, the decoy wins and it is an empty string. The xml parser strips
      // prefixes by default, which is exactly this trap.
      same(read("edge.ods").sheets[1].cells.I1, { value: 42 });

      // The same collision LibreOffice writes on every error cell.
      same(read("edge.ods").sheets[2].cells.C1, { value: null, formula: "#REF!+1", computedValue: "#REF!" });
    });

    it("collapses white space as ODF defines it", () => {
      const cells = read("edge.ods").sheets[1].cells;

      same(cells.A1, { value: "leading and inner tabs newline " });
      same(cells.B1, { value: "a   b   " });
      same(cells.D1, { value: "xyz ☺ & <raw>" });
    });
  });

  describe("constructs LibreOffice normalises away, built in memory", () => {
    it("expands rows repeated with content, and skips a trailing million-row repeat", () => {
      const schema: Any = Agent.read(
        odsPackage(
          '<table:table table:name="R">' +
            '<table:table-row table:number-rows-repeated="3">' +
            '<table:table-cell office:value-type="float" office:value="5"/>' +
            '<table:table-cell table:number-columns-repeated="2"/>' +
            '<table:table-cell office:value-type="string"><text:p>tail</text:p></table:table-cell>' +
            "</table:table-row>" +
            '<table:table-row table:number-rows-repeated="1048573"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>' +
            "</table:table>",
        ),
      );

      same(schema.sheets[0].cells, {
        A1: { value: 5 },
        D1: { value: "tail" },
        A2: { value: 5 },
        D2: { value: "tail" },
        A3: { value: 5 },
        D3: { value: "tail" },
      });
    });

    it("reads rows inside header-row and row groups, in document order", () => {
      const schema: Any = Agent.read(
        odsPackage(
          '<table:table table:name="G">' +
            '<table:table-header-rows><table:table-row><table:table-cell office:value-type="string"><text:p>head</text:p></table:table-cell></table:table-row></table:table-header-rows>' +
            '<table:table-row-group><table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row>' +
            '<table:table-row-group><table:table-row><table:table-cell office:value-type="float" office:value="2"/></table:table-row></table:table-row-group>' +
            "</table:table-row-group>" +
            '<table:table-row><table:table-cell office:value-type="float" office:value="3"/></table:table-row>' +
            "</table:table>",
        ),
      );

      same(schema.sheets[0].cells, { A1: { value: "head" }, A2: { value: 1 }, A3: { value: 2 }, A4: { value: 3 } });
    });

    it("keeps content that sits in a covered cell, and records both span directions", () => {
      const schema: Any = Agent.read(
        odsPackage(
          '<table:table table:name="M"><table:table-row>' +
            '<table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2" office:value-type="string"><text:p>big</text:p></table:table-cell>' +
            '<table:covered-table-cell office:value-type="string"><text:p>hidden</text:p></table:covered-table-cell>' +
            "</table:table-row></table:table>",
        ),
      );

      same(schema.sheets[0].cells, { A1: { value: "big" }, B1: { value: "hidden" } });
      same(schema.sheets[0].mergedRegions, [{ start: "A1", end: "B2" }]);
    });

    it("reports a formula result of every type as the xlsx reader would cache it", () => {
      const cells: Any = (
        Agent.read(
          odsPackage(
            '<table:table table:name="F"><table:table-row>' +
              '<table:table-cell table:formula="of:=DATE(2024;3;15)" office:value-type="date" office:date-value="2024-03-15"/>' +
              '<table:table-cell table:formula="of:=[.A1]+0.5" office:value-type="date" office:date-value="2024-03-15T12:00:00"/>' +
              '<table:table-cell table:formula="of:=TIME(6;0;0)" office:value-type="time" office:time-value="PT6H"/>' +
              '<table:table-cell table:formula="of:=1=1" office:value-type="boolean" office:boolean-value="true"/>' +
              '<table:table-cell table:formula="of:=&quot;a&quot;" office:value-type="string"><text:p>a</text:p></table:table-cell>' +
              "</table:table-row></table:table>",
          ),
        ) as Any
      ).sheets[0].cells;

      same(cells.A1, { value: null, formula: "DATE(2024,3,15)", computedValue: 45366, format: { displayFormat: "date" } });
      same(cells.B1, { value: null, formula: "A1+0.5", computedValue: 45366.5, format: { displayFormat: "datetime" } });
      same(cells.C1, { value: null, formula: "TIME(6,0,0)", computedValue: 0.25, format: { displayFormat: "datetime" } });
      same(cells.D1, { value: null, formula: "1=1", computedValue: true });
      same(cells.E1, { value: null, formula: '"a"', computedValue: "a" });
    });

    it("passes Excel-syntax formulas through, and translates the older OpenOffice prefix", () => {
      const cells: Any = (
        Agent.read(
          odsPackage(
            '<table:table table:name="P"><table:table-row>' +
              '<table:table-cell table:formula="msoxl:=SUM(A1:B2,C3)" office:value-type="float" office:value="0"/>' +
              '<table:table-cell table:formula="oooc:=SUM([.A1:.B2];[.C3])" office:value-type="float" office:value="0"/>' +
              '<table:table-cell table:formula="of:=SUM([&apos;It&apos;&apos;s here&apos;.A1:.A2];[$Other.$B$1:.$B$9])" office:value-type="float" office:value="0"/>' +
              "</table:table-row></table:table>",
          ),
        ) as Any
      ).sheets[0].cells;

      expect(cells.A1.formula).toBe("SUM(A1:B2,C3)");
      expect(cells.B1.formula).toBe("SUM(A1:B2,C3)");
      expect(cells.C1.formula).toBe("SUM('It''s here'!A1:A2,Other!$B$1:$B$9)");
    });

    it("converts date and time values exactly", () => {
      const cells: Any = (
        Agent.read(
          odsPackage(
            '<table:table table:name="D"><table:table-row>' +
              '<table:table-cell office:value-type="time" office:time-value="PT36H15M30S"/>' +
              '<table:table-cell office:value-type="date" office:date-value="2024-01-01T23:59:59.6"/>' +
              '<table:table-cell office:value-type="date" office:date-value="2024-06-01T10:00:00+02:00"/>' +
              '<table:table-cell office:value-type="string"><text:p/></table:table-cell>' +
              "<table:table-cell><text:p>untyped</text:p></table:table-cell>" +
              "</table:table-row></table:table>",
          ),
        ) as Any
      ).sheets[0].cells;

      same(cells.A1, { value: "1899-12-31T12:15:30Z", format: { displayFormat: "datetime" } });
      same(cells.B1, { value: "2024-01-02T00:00:00Z", format: { displayFormat: "datetime" } });
      same(cells.C1, { value: "2024-06-01T08:00:00Z", format: { displayFormat: "datetime" } });
      same(cells.D1, { value: "" });
      same(cells.E1, { value: "untyped" });
    });
  });

  describe("dispatch", () => {
    it("still describes xlsx through the same entry point", async () => {
      const schema: Any = await Agent.describe(odsFixturePath("workbook.xlsx"));
      expect(schema.sheets[0].cells.A2.value).toBe(100);
    });

    it("describes an ods from a path too", async () => {
      const schema: Any = await Agent.describe(odsFixturePath("workbook.ods"));
      expect(schema.sheets[0].cells.A2.value).toBe(100);
    });

    it("names what it cannot read instead of failing on a zip it does not know", () => {
      const bytes = odsPackage('<table:table table:name="X"/>', "", "application/vnd.oasis.opendocument.text");

      expect(() => Agent.read(bytes)).toThrow(UnsupportedFormatException);
      expect(() => Agent.read(bytes)).toThrow("application/vnd.oasis.opendocument.text");
    });

    it("refuses bytes that are not a zip at all, and the exception is still an Error", () => {
      const bytes = new TextEncoder().encode("a,b\n1,2\n");
      let caught: unknown;
      try {
        Agent.read(bytes);
      } catch (e) {
        caught = e;
      }

      expect(caught).toBeInstanceOf(UnsupportedFormatException);
      expect(caught).toBeInstanceOf(Error);
      // The phrase the old Error carried, kept for anyone matching on it.
      expect((caught as Error).message).toContain("not a zip archive");
    });
  });
});
