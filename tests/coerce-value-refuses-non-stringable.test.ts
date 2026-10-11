import { describe, expect, it } from "vitest";

import { SchemaException } from "../src/exceptions";
import { Normalizer } from "../src/schema/normalizer";
import { Validator } from "../src/schema/validator";

/**
 * holy-sheet-js#2, mirroring the PHP reference fix in holy-sheet 2.5.0
 * (holy-sheet#9, reported from MOIC production).
 *
 * `coerceValue()` ended with "not a primitive, so it must be safe to
 * stringify". An ARRAY reaches that line whenever a plain `rows` form is nested
 * one level too deep — no object-cell syntax needed, which is exactly the shape
 * an AI-generated edit produces.
 *
 * THIS RUNTIME IS THE DANGEROUS ONE. `Array.prototype.toString` joins, so
 * `["x"]` became `"x"` — exactly what the author probably meant — and
 * `["x","y"]` became `"x,y"`, two cells silently merged into one with the data
 * intact enough to look deliberate. PHP wrote the visibly-wrong `"Array"`,
 * which someone files a bug about. A wrong answer that looks right is the
 * expensive kind.
 *
 * Every expectation below was MEASURED against both engines before being
 * written (`Normalizer` directly, not `String()` — see the dispatcher note in
 * the second block, which is where my own issue report had it wrong).
 */

class Stringy {
  toString(): string {
    return "from-stringable";
  }
}

/** No custom `toString`, so it inherits `Object.prototype`'s — PHP's non-Stringable object. */
class Opaque {
  readonly a = 1;
}

function normalizeRow(row: unknown[], name = "S") {
  return new Normalizer().normalize({ sheets: [{ name, rows: [row] }] });
}

function cellA1(row: unknown[]) {
  return normalizeRow(row).sheets[0]!.cells.get("A1")?.value;
}

describe("coerceValue refuses what it cannot represent, and names the cell", () => {
  it("refuses an array cell and names the cell address", () => {
    try {
      normalizeRow([["x"], "ok"]);
      expect.unreachable("normalize() accepted an array cell instead of refusing it");
    } catch (e) {
      expect(e).toBeInstanceOf(SchemaException);
      const errors = (e as SchemaException).getErrors();
      expect(errors).toHaveLength(1);
      expect(errors[0]!.path).toBe("S!A1");
      expect(errors[0]!.got).toBe("array");
      expect((e as SchemaException).message).toContain("S!A1");
    }
  });

  it("refuses a two-element array rather than joining it into one cell", () => {
    // The specific corruption: "x,y" in A1 and nothing in B1, which reads as a
    // deliberate comma-separated value.
    expect(() => normalizeRow([["x", "y"]])).toThrow(SchemaException);
  });

  it("refuses an object with no toString of its own, and names the cell", () => {
    try {
      normalizeRow([new Opaque()]);
      expect.unreachable("normalize() accepted an opaque object instead of refusing it");
    } catch (e) {
      expect(e).toBeInstanceOf(SchemaException);
      expect((e as SchemaException).getErrors()[0]!.path).toBe("S!A1");
    }
  });

  it("refuses a Map cell", () => {
    // `Map.prototype` has no `toString`, so it inherits Object's and would
    // stringify to "[object Map]".
    expect(() => normalizeRow([new Map([["a", 1]])])).toThrow(SchemaException);
  });

  /**
   * NOT in the PHP reference, because PHP has no analogue: a JS function's
   * `toString` returns its SOURCE TEXT, and `Function.prototype.toString` is
   * "its own", so a naive "has a custom toString" guard ACCEPTS it. This wrote
   * the function body into the spreadsheet cell.
   */
  it("refuses a function cell instead of writing its source code into the sheet", () => {
    expect(() =>
      normalizeRow([
        function hello() {
          return 1;
        },
      ]),
    ).toThrow(SchemaException);
  });

  it("refuses a symbol cell", () => {
    expect(() => normalizeRow([Symbol("x")])).toThrow(SchemaException);
  });

  /**
   * The address has to be right for a cell that is not A1, or the message sends
   * the caller to the wrong place — worse than no address at all.
   */
  it("names the correct address for an array deeper in the sheet", () => {
    try {
      new Normalizer().normalize({
        sheets: [{ name: "Data", rows: [["a", "b"], ["c", ["nested"]]] }],
      });
      expect.unreachable("normalize() accepted an array cell instead of refusing it");
    } catch (e) {
      expect((e as SchemaException).getErrors()[0]!.path).toBe("Data!B2");
    }
  });
});

/**
 * The dispatcher, which is the half my own issue report got wrong.
 *
 * `buildCell` tested `isPlainObject(value)` — `typeof === "object" && !null &&
 * !Array.isArray` — which is true for EVERY non-array object, so a `Date`, a
 * class instance and a `Map` were all read as the `{value, format}` cell-spec
 * form. `value.value ?? null` is then null, and the cell came out EMPTY.
 *
 * So the predictions in holy-sheet-js#2 were wrong in both directions: `{}` did
 * not become `"[object Object]"` and a null-prototype object did not throw —
 * both became empty cells, because neither ever reached `coerceValue`. The
 * table in that issue was measured with raw `String()` instead of through the
 * code path. Measuring the wrong LAYER answers confidently about nothing.
 *
 * The expensive one is `Date`: PHP's guard is `is_array && !array_is_list`, and
 * a PHP `DateTimeInterface` is an object rather than an array, so it fell
 * straight through to the date branch. PHP wrote the serial; this port wrote an
 * empty cell. A date in a plain row is the most ordinary non-string cell there
 * is, and this pair's whole contract is "same bytes either backend" — the
 * byte-parity suite never saw it because its fixtures carry ISO date STRINGS
 * rather than Date objects.
 */
describe("the cell-spec dispatcher only claims a plain object", () => {
  /**
   * Pinned as an EQUIVALENCE, not a constant. `DateConverter` reads a Date's
   * instant in UTC (measured: a UTC-midnight Date and the string "2026-01-02"
   * both give 46024, while a LOCAL-midnight Date gives 46024 + the offset), so
   * `expect(...).toBe(46024)` for `new Date(2026, 0, 2)` passes only in UTC and
   * fails in every other timezone — a test that holds on the machine that wrote
   * it. The invariant that actually matters here is that the plain-row form and
   * the object-cell form produce the SAME cell: the dispatcher no longer
   * swallows the value.
   */
  it("writes a Date in a plain row, rather than an empty cell", () => {
    const viaRow = cellA1([new Date(2026, 0, 2)]);
    const viaObjectCell = cellA1([{ value: new Date(2026, 0, 2) }]);

    expect(viaRow).not.toBeNull();
    expect(typeof viaRow).toBe("number");
    expect(viaRow).toBe(viaObjectCell);
  });

  /** And the value itself, on an input with no timezone to disagree about. */
  it("converts a UTC Date to the Excel serial PHP produces for the same instant", () => {
    expect(cellA1([new Date("2026-01-02T00:00:00Z")])).toBe(46024);
  });

  it("accepts an object whose toString is its own", () => {
    expect(cellA1([new Stringy()])).toBe("from-stringable");
  });

  it("still reads the {value, format} object-cell form as a cell spec", () => {
    const wb = normalizeRow([{ value: "x", format: { bold: true } }]);
    const cell = wb.sheets[0]!.cells.get("A1")!;
    expect(cell.value).toBe("x");
    expect(cell.format?.bold).toBe(true);
  });

  it("still treats a plain object with no value key as an empty cell, as PHP does", () => {
    // PHP: an associative array with no 'value' key normalizes to null rather
    // than throwing. Measured, and this agrees with it.
    expect(cellA1([{ a: 1 }])).toBeNull();
  });
});

/** A refusal must not become a refusal of things that already worked. */
describe("what already worked still works", () => {
  it("accepts the scalar cell values it always did", () => {
    const cells = normalizeRow(["text", 42, 3.5, true, null]).sheets[0]!.cells;
    expect(cells.get("A1")!.value).toBe("text");
    expect(cells.get("B1")!.value).toBe(42);
    expect(cells.get("C1")!.value).toBe(3.5);
    expect(cells.get("D1")!.value).toBe(true);
  });

  it("still stringifies a bigint", () => {
    // Unchanged on purpose: a bigint is a primitive that converts losslessly and
    // deliberately, and PHP has no counterpart to diverge from.
    expect(cellA1([10n])).toBe("10");
  });
});

/**
 * The two halves of the contract have to AGREE.
 *
 * `validate()` returned ZERO errors for the schema `normalize()` now throws on,
 * so `Agent.validate()` told a caller the input was fine and the write then
 * failed — in exactly the validate-then-write order an agent uses, which is how
 * MOIC met it in PHP.
 */
describe("validate() agrees with the writer", () => {
  it("rejects a list-array cell rather than passing it to the writer", () => {
    const errors = new Validator().validate({
      sheets: [{ name: "S", rows: [[["x"], "ok"]] }],
    });

    expect(errors).not.toHaveLength(0);
    expect(errors[0]!.path).toBe("sheets[0].rows[0][0]");
    expect(errors[0]!.got).toBe("array");
  });

  /**
   * And must NOT reject the object cell form, which has always been legal.
   * Rejecting "any object" would trade this bug for a far worse one — every
   * formatted cell in every workbook.
   */
  it("still accepts the object cell form", () => {
    const errors = new Validator().validate({
      sheets: [{ name: "S", rows: [[{ value: "x", format: { bold: true } }, "plain"]] }],
    });

    expect(errors).toHaveLength(0);
  });
});
