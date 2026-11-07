import { describe, expect, it } from "vitest";
import { Agent, unzipSync } from "../src";

/**
 * Data validation, ported case for case from PHP's
 * `tests/Unit/DataValidationTest.php`. The error paths, `expected` texts and
 * hints are PHP's, byte for byte; the worksheet bytes are pinned against PHP's
 * writer by the `dataValidations` fixture in `parity.test.ts`.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

function sheetXml(schema: Any, n = 1): string {
  const parts = unzipSync(Agent.toBytes(schema));
  return new TextDecoder().decode(parts[`xl/worksheets/sheet${n}.xml`]!);
}

export function dvFixture(): Any {
  return {
    meta: { creator: "Parity", created: "2024-01-01T00:00:00Z" },
    sheets: [
      {
        name: "Catalog",
        columns: [
          { header: "SKU" },
          {
            header: "Category",
            validation: {
              type: "list",
              source: "'Lists'!A1:A3",
              prompt: { title: "Category", text: "Pick a category." },
              error: { style: "stop", title: "Unknown category", text: "Choose a category from the list." },
            },
          },
          { header: "Status", validation: { type: "list", values: ["Draft", "Live", 'Retired "old"'], showDropdown: true } },
          { header: "Qty", type: "integer", validation: { type: "whole", operator: "between", min: 0, max: 1000 } },
        ],
        rows: [
          ["A-1", "Tools", "Draft", 3],
          ["A-2", "Parts", "Live", 10],
        ],
        mergedRegions: [{ start: "K1", end: "L1" }],
        print: { orientation: "landscape" },
        dataValidations: [
          { range: "E2:E500", type: "decimal", operator: "greaterThanOrEqual", value: 0.5, allowBlank: false, error: { style: "warning", text: "Below the minimum & then some <x>" } },
          { range: "F2:F500", type: "date", operator: "notBetween", min: "2026-01-01", max: "2026-12-31", error: { style: "information" } },
          { range: "G2:G500", type: "textLength", operator: "lessThanOrEqual", value: 40, prompt: { text: "At most 40 characters." } },
          { range: "H2:H500", type: "list", values: ["Yes", "No", 7], showDropdown: false },
          { range: "i2:I500", type: "list", source: "A20:A22" },
          { range: "J2", type: "decimal", operator: "lessThan", value: 2.25 },
        ],
      },
      { name: "Lists", cells: { A1: "Tools", A2: "Parts", A3: "Kits" } },
    ],
  };
}

const EXPECTED =
  '<dataValidations count="9">' +
  '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Unknown category" error="Choose a category from the list." promptTitle="Category" prompt="Pick a category." sqref="B2:B3"><formula1>&apos;Lists&apos;!$A$1:$A$3</formula1></dataValidation>' +
  '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="C2:C3"><formula1>&quot;Draft,Live,Retired &quot;&quot;old&quot;&quot;&quot;</formula1></dataValidation>' +
  '<dataValidation type="whole" allowBlank="1" showErrorMessage="1" sqref="D2:D3"><formula1>0</formula1><formula2>1000</formula2></dataValidation>' +
  '<dataValidation type="decimal" errorStyle="warning" operator="greaterThanOrEqual" showErrorMessage="1" error="Below the minimum &amp; then some &lt;x&gt;" sqref="E2:E500"><formula1>0.5</formula1></dataValidation>' +
  '<dataValidation type="date" errorStyle="information" operator="notBetween" allowBlank="1" showErrorMessage="1" sqref="F2:F500"><formula1>46023</formula1><formula2>46387</formula2></dataValidation>' +
  '<dataValidation type="textLength" operator="lessThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" prompt="At most 40 characters." sqref="G2:G500"><formula1>40</formula1></dataValidation>' +
  '<dataValidation type="list" allowBlank="1" showDropDown="1" showErrorMessage="1" sqref="H2:H500"><formula1>&quot;Yes,No,7&quot;</formula1></dataValidation>' +
  '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="I2:I500"><formula1>$A$20:$A$22</formula1></dataValidation>' +
  '<dataValidation type="decimal" operator="lessThan" allowBlank="1" showErrorMessage="1" sqref="J2"><formula1>2.25</formula1></dataValidation>' +
  "</dataValidations>";

function dvErrors(validation: Any, where: "sheet" | "column" = "sheet"): [string, string][] {
  const sheet: Any = { name: "S", columns: [{ header: "A" }], rows: [["x"]] };
  if (where === "sheet") sheet.dataValidations = [validation];
  else sheet.columns[0].validation = validation;
  const schema = { sheets: [sheet, { name: "Lists", cells: { A1: "a" } }] };
  return Agent.validate(schema).map((e) => [e.path, e.expected]);
}

const P = "sheets[0].dataValidations[0]";
const CHOICE = "a non-empty string with no comma, or a whole number";

const CASES: [string, Any, [string, string][], ("sheet" | "column")?][] = [
  ["not an object", ["list"], [[P, "object"]]],
  ["no type", { range: "A2" }, [[`${P}.type`, "one of: list, whole, decimal, date, textLength"]]],
  ["unknown type", { range: "A2", type: "dropdown", values: ["a"] }, [[`${P}.type`, "one of: list, whole, decimal, date, textLength"]]],
  ["no range", { type: "list", values: ["a"] }, [[`${P}.range`, 'a cell or range on this sheet, like "B2:B100"']]],
  ["range on another sheet", { range: "Lists!A2", type: "list", values: ["a"] }, [[`${P}.range`, 'a cell or range on this sheet, like "B2:B100"']]],
  ["range past the grid", { range: "A2:XFE2", type: "list", values: ["a"] }, [[`${P}.range`, 'a cell or range on this sheet, like "B2:B100"']]],
  ["range on a column rule", { range: "A2", type: "list", values: ["a"] }, [["sheets[0].columns[0].validation.range", "absent"]], "column"],
  ["both values and source", { range: "A2", type: "list", values: ["a"], source: "Lists!A1" }, [[P, "exactly one of values or source"]]],
  ["neither values nor source", { range: "A2", type: "list" }, [[P, "exactly one of values or source"]]],
  ["values not a list", { range: "A2", type: "list", values: "a,b" }, [[`${P}.values`, "a non-empty list of choices"]]],
  ["empty values", { range: "A2", type: "list", values: [] }, [[`${P}.values`, "a non-empty list of choices"]]],
  ["a choice with a comma", { range: "A2", type: "list", values: ["a", "b,c", "", 1.5, true] }, [
    [`${P}.values[1]`, CHOICE], [`${P}.values[2]`, CHOICE], [`${P}.values[3]`, CHOICE], [`${P}.values[4]`, CHOICE],
  ]],
  ["source not a range", { range: "A2", type: "list", source: "Lists" }, [[`${P}.source`, 'a cell range like "\'Lists\'!A1:A10"']]],
  ["source on an unknown sheet", { range: "A2", type: "list", source: "'Lsts'!A1:A3" }, [[`${P}.source`, "a sheet in this workbook"]]],
  ["operator on a list", { range: "A2", type: "list", values: ["a"], operator: "equal", value: 1 }, [
    [`${P}.operator`, "absent"], [`${P}.value`, "absent"],
  ]],
  ["list keys on a number rule", { range: "A2", type: "whole", value: 1, operator: "equal", values: ["a"], source: "A1", showDropdown: true }, [
    [`${P}.values`, "absent"], [`${P}.source`, "absent"], [`${P}.showDropdown`, "absent"],
  ]],
  ["unknown operator", { range: "A2", type: "whole", operator: "atLeast", value: 1 }, [[`${P}.operator`, "one of: between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual"]]],
  ["between without bounds", { range: "A2", type: "whole" }, [[`${P}.min`, "a whole number"], [`${P}.max`, "a whole number"]]],
  ["between with value", { range: "A2", type: "decimal", min: 1, max: 2, value: 3 }, [[`${P}.value`, "absent"]]],
  ["one-value operator with bounds", { range: "A2", type: "decimal", operator: "greaterThan", min: 1, max: 2 }, [
    [`${P}.value`, "a number"], [`${P}.min`, "absent"], [`${P}.max`, "absent"],
  ]],
  ["whole with a fraction", { range: "A2", type: "whole", operator: "equal", value: 1.5 }, [[`${P}.value`, "a whole number"]]],
  ["decimal with a string", { range: "A2", type: "decimal", operator: "equal", value: "1.5" }, [[`${P}.value`, "a number"]]],
  ["decimal with a bool", { range: "A2", type: "decimal", operator: "equal", value: true }, [[`${P}.value`, "a number"]]],
  ["date not ISO", { range: "A2", type: "date", operator: "equal", value: "31/12/2026" }, [[`${P}.value`, "a date as YYYY-MM-DD"]]],
  ["date not a day", { range: "A2", type: "date", operator: "equal", value: "2026-02-30" }, [[`${P}.value`, "a date as YYYY-MM-DD"]]],
  ["negative text length", { range: "A2", type: "textLength", operator: "equal", value: -1 }, [[`${P}.value`, "a whole number of characters, 0 or more"]]],
  ["min above max", { range: "A2", type: "date", min: "2026-12-31", max: "2026-01-01" }, [[`${P}.max`, "a value no less than min"]]],
  ["flags not booleans", { range: "A2", type: "list", values: ["a"], allowBlank: 1, showDropdown: "no" }, [
    [`${P}.showDropdown`, "boolean"], [`${P}.allowBlank`, "boolean"],
  ]],
  ["bad prompt", { range: "A2", type: "list", values: ["a"], prompt: "pick" }, [[`${P}.prompt`, "object {title?, text?}"]]],
  ["long prompt", { range: "A2", type: "list", values: ["a"], prompt: { title: "t".repeat(33), text: "p".repeat(256) } }, [
    [`${P}.prompt.title`, "a string of at most 32 characters"], [`${P}.prompt.text`, "a string of at most 255 characters"],
  ]],
  ["bad error", { range: "A2", type: "list", values: ["a"], error: { style: "fatal", title: 7, text: "e".repeat(226) } }, [
    [`${P}.error.style`, "one of: stop, warning, information"],
    [`${P}.error.title`, "a string of at most 32 characters"],
    [`${P}.error.text`, "a string of at most 225 characters"],
  ]],
  ["column rule fine", { type: "list", values: ["a"] }, [], "column"],
];

describe("data validation", () => {
  it("accepts the whole fixture without an error", () => {
    expect(Agent.validate(dvFixture())).toEqual([]);
  });

  it("writes every validation, column rules first, then the sheet list", () => {
    expect(sheetXml(dvFixture())).toContain(EXPECTED);
  });

  it("places dataValidations after mergeCells and before pageSetup", () => {
    const xml = sheetXml(dvFixture());
    const merge = xml.indexOf("</mergeCells>");
    const dv = xml.indexOf("<dataValidations ");
    const page = xml.indexOf("<pageSetup ");
    expect(merge).toBeGreaterThan(0);
    expect(merge).toBeLessThan(dv);
    expect(dv).toBeLessThan(page);
  });

  it("writes nothing for a sheet without validations", () => {
    const schema = dvFixture();
    delete schema.sheets[0].dataValidations;
    for (const col of schema.sheets[0].columns) delete col.validation;
    expect(sheetXml(schema)).not.toContain("dataValidation");
  });

  it("writes a single-row column rule as one cell", () => {
    const xml = sheetXml({ sheets: [{ name: "S", columns: [{ header: "Status", validation: { type: "list", values: ["a", "b"] } }], rows: [["a"]] }] });
    expect(xml).toContain('sqref="A2"><formula1>&quot;a,b&quot;</formula1>');
  });

  it("quotes a sheet name with an apostrophe in a source", () => {
    const xml = sheetXml({
      sheets: [
        { name: "S", cells: { A1: "x" }, dataValidations: [{ range: "A2:A9", type: "list", source: "'Bob''s list'!$B$2:B4" }] },
        { name: "Bob's list", cells: { B2: "y" } },
      ],
    });
    expect(xml).toContain("<formula1>&apos;Bob&apos;&apos;s list&apos;!$B$2:$B$4</formula1>");
  });

  it("counts an inline list at 255 characters, commas included", () => {
    const values = Array(51).fill("abcd");
    const at = (v: string[]) => Agent.validate({ sheets: [{ name: "S", cells: {}, dataValidations: [{ range: "A1", type: "list", values: v }] }] });
    expect(at(values)).toEqual([]);
    values[0] = "abcde";
    expect(at(values)).toEqual([]);
    values[1] = "abcde";
    const errors = at(values);
    expect(errors).toHaveLength(1);
    expect(errors[0]!.path).toBe("sheets[0].dataValidations[0].values");
    expect(errors[0]!.expected).toBe("at most 255 characters, commas included");
    expect(errors[0]!.got).toBe("256 characters");
    expect(errors[0]!.hint).toBe(
      "Excel cannot store an inline list longer than 255 characters. Put the choices in cells and point source at them, like \"'Lists'!A1:A40\".",
    );
  });

  it("counts characters, not UTF-16 units or bytes", () => {
    // An astral character is one character to PHP's mb_strlen and two to .length.
    const v = "\u{1F600}".repeat(127);
    expect(Agent.validate({ sheets: [{ name: "S", cells: {}, dataValidations: [{ range: "A1", type: "list", values: [v, v] }] }] })).toEqual([]);
  });

  for (const [name, validation, expected, where] of CASES) {
    it(`reports each bad validation at its path, in PHP wording: ${name}`, () => {
      expect(dvErrors(validation, where ?? "sheet")).toEqual(expected);
    });
  }

  it("refuses a column rule on a sheet with no data rows, and says what to use instead", () => {
    const errors = Agent.validate({ sheets: [{ name: "Template", columns: [{ header: "Category", validation: { type: "list", values: ["a"] } }], rows: [] }] });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.path).toBe("sheets[0].columns[0].validation");
    expect(errors[0]!.expected).toBe("a sheet with data rows");
    expect(errors[0]!.hint).toBe(
      "A column validation covers the column's data rows, and this sheet has none. For a template, put it in the sheet's dataValidations with a range, like \"B2:B500\".",
    );
  });

  it("refuses a dataValidations that is not a list", () => {
    const errors = Agent.validate({ sheets: [{ name: "S", cells: {}, dataValidations: { range: "A1" } }] });
    expect(errors.map((e) => e.path)).toEqual(["sheets[0].dataValidations"]);
    expect(errors[0]!.expected).toBe("a list of validations");
  });

  it("gives a hint that names the comma rule", () => {
    const errors = Agent.validate({ sheets: [{ name: "S", cells: {}, dataValidations: [{ range: "A1", type: "list", values: ["a,b"] }] }] });
    expect(errors[0]!.hint).toBe(
      'Each choice is text, like "Open", or a whole number. Excel separates inline choices with commas and cannot escape one, so a choice containing a comma must come from cells: use source.',
    );
  });

  it("repairs only the case of a type, operator or error style", () => {
    const out = Agent.validateAndRepair({
      sheets: [
        {
          name: "S",
          columns: [{ header: "A", validation: { type: "List", values: ["a"] } }],
          rows: [["a"]],
          dataValidations: [
            { range: "B2", type: "WHOLE", operator: "greaterthan", value: 1, error: { style: "Warning" } },
            { range: "C2", type: "dropdown", values: ["a"] },
          ],
        },
      ],
    });
    expect(out.repairs).toEqual([
      "changed 'sheets[0].columns[0].validation.type' from 'List' to 'list'",
      "changed 'sheets[0].dataValidations[0].type' from 'WHOLE' to 'whole'",
      "changed 'sheets[0].dataValidations[0].operator' from 'greaterthan' to 'greaterThan'",
      "changed 'sheets[0].dataValidations[0].error.style' from 'Warning' to 'warning'",
    ]);
    expect(out.errors.map((e) => e.path)).toEqual(["sheets[0].dataValidations[1].type"]);
    expect((out.schema as Any).sheets[0].dataValidations[0].operator).toBe("greaterThan");
  });

  it("does not repair a validation that is already right", () => {
    const out = Agent.validateAndRepair(dvFixture());
    expect(out.repairs).toEqual([]);
    expect(out.errors).toEqual([]);
  });
});
