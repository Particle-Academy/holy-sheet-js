import { describe, expect, it } from "vitest";

import { Agent, unzipSync } from "../src";
import { formatFloat, isNumericString, numericStringToNumber } from "../src/util";

function cellXmlFor(value: unknown): string {
  const parts = unzipSync(Agent.toBytes({ sheets: [{ name: "S", rows: [[value]] }] }));
  const xml = new TextDecoder().decode(parts["xl/worksheets/sheet1.xml"]!);
  return xml.match(/<c r="A1"[^>]*\/>|<c r="A1"[^>]*>.*?<\/c>/)?.[0] ?? "(no A1)";
}

/**
 * The numeric hazards `.ai/plans/polyglot/parity/conformance-suite.md` §5 names,
 * pinned as cases.
 *
 * Every one of these was a live PHP↔JS disagreement in shipped packages, and
 * none of them was covered — the parity suites diff whole OOXML parts, so they
 * only catch a divergence if some fixture happens to contain the value. None
 * did. Numbers are the thing a spreadsheet writer exists to get right, so they
 * get their own table rather than waiting to be caught incidentally.
 *
 * The reference implementation is the PHP writer, because the pair's contract
 * is "same bytes either backend" and PHP is what shipped first:
 *
 *   (float) $v                                        — coercion
 *   rtrim(rtrim(number_format($v, 14, '.', ''), '0'), '.')  — emission
 */

describe("formatFloat — plain decimal, never exponential", () => {
  /**
   * The severe one. `toFixed` switches to exponential notation at 1e21, and the
   * trailing-zero strip then ate the EXPONENT's zeros rather than a fraction's:
   *
   *   (1e300).toFixed(14)  === "1e+300"
   *   "1e+300".replace(/0+$/, "") === "1e+3"
   *
   * So 1e300 was written to the sheet as 1e3. Not a crash, not invalid XML — a
   * different number, silently, in a file the user opens later and believes.
   * PHP's number_format has no exponential mode and always got this right.
   */
  it("does not turn 1e300 into 1e+3", () => {
    const s = formatFloat(1e300);

    expect(s).not.toContain("e");

    // Golden digits taken from the PHP writer, which is the reference:
    //   rtrim(rtrim(number_format(1e300, 14, '.', ''), '0'), '.')
    // Note it is NOT 1 followed by 300 zeros — number_format prints the EXACT
    // value of the double, and the nearest double to 1e300 is not 10^300. I
    // asserted the round number first and PHP corrected me, which is the whole
    // argument for pinning goldens from the reference rather than from what the
    // value "obviously" is.
    expect(s).toHaveLength(301);
    expect(s.slice(0, 30)).toBe("100000000000000005250476025520");
    expect(s.slice(-30)).toBe("115669472196386865459400540160");
  });

  it("does not emit exponential notation at the 1e21 boundary", () => {
    // `<v>1e+21</v>` is not valid cell content — a reader is entitled to reject
    // the whole sheet, so this one at least failed loudly somewhere downstream.
    expect(formatFloat(1e21)).toBe("1" + "0".repeat(21));
    expect(formatFloat(1e21)).not.toContain("e");
  });

  it("still formats ordinary numbers exactly as before", () => {
    // The fix must not move the values that were already correct — this is the
    // regression half, and it is the reason the fix is not just `String(v)`.
    expect(formatFloat(1.5)).toBe("1.5");
    expect(formatFloat(0)).toBe("0");
    expect(formatFloat(100000)).toBe("100000");
    expect(formatFloat(-2.25)).toBe("-2.25");
    expect(formatFloat(0.1)).toBe("0.1");
  });

  it("keeps 14 decimal places and trims only the fraction", () => {
    expect(formatFloat(1 / 3)).toBe("0.33333333333333");
    // Trailing zeros inside the INTEGER part must survive; only the fraction is
    // trimmed. This is what the old blanket /0+$/ got wrong.
    expect(formatFloat(1200)).toBe("1200");
  });

  it("collapses negative zero to 0", () => {
    // The polyglot plan's §5 lists this as a live divergence -- "PHP <v>-0</v>,
    // JS <v>0</v>". It is NOT one, and I checked rather than fixing what the
    // document said: PHP's number_format(-0.0, 14) returns "0.00000000000000",
    // with no sign, so both engines already emitted 0. Nothing to reconcile.
    //
    // Kept as a regression guard, since both sides drop the sign incidentally
    // (via number_format / toFixed) rather than by decision, and a future
    // "faster" formatter that used a raw string conversion would reintroduce it.
    expect(formatFloat(-0)).toBe("0");
  });

  it("never emits a non-finite value into a cell", () => {
    // NaN/Infinity have no `<v>` representation at all. Whatever we choose it
    // must be finite and identical on both engines.
    for (const v of [NaN, Infinity, -Infinity]) {
      const s = formatFloat(v);
      expect(s).not.toMatch(/nan|inf/i);
      expect(Number.isFinite(Number(s))).toBe(true);
    }
  });
});

describe("numericStringToNumber — matches PHP's numeric-string coercion", () => {
  /**
   * The helper's own docstring said it mirrored PHP's `(int)`/`(float)`. It did
   * not: `parseInt` stops at the first non-digit, so the exponent is discarded
   * and "1e5" became 1 rather than 100000. PHP has read exponent notation in
   * numeric strings since PHP 7.
   */
  it("reads exponent notation instead of truncating at the 'e'", () => {
    expect(numericStringToNumber("1e5")).toBe(100000);
    expect(numericStringToNumber("1E5")).toBe(100000);
    expect(numericStringToNumber("1.5e3")).toBe(1500);
    expect(numericStringToNumber("2e-3")).toBe(0.002);
  });

  it("handles the magnitude where PHP's own (int) cast breaks", () => {
    // PHP's (int) "1e21" clamps to PHP_INT_MAX (9223372036854775807), which is
    // not the number anyone wrote. Both sides now take the float path, so both
    // say 1e21.
    expect(numericStringToNumber("1e21")).toBe(1e21);
  });

  it("keeps the plain cases exactly where they were", () => {
    expect(numericStringToNumber("007")).toBe(7);
    expect(numericStringToNumber(".5")).toBe(0.5);
    expect(numericStringToNumber("-3.25")).toBe(-3.25);
    expect(numericStringToNumber("0")).toBe(0);
    expect(numericStringToNumber("42")).toBe(42);
  });

  it("agrees with isNumericString about what it accepts", () => {
    // A string the gate admits must coerce to a finite number; otherwise a cell
    // ends up holding NaN, which has no `<v>` form.
    for (const s of ["1e5", "007", ".5", "-3.25", "0", "1e21", "+1.5e-2"]) {
      expect(isNumericString(s), `${s} should be numeric`).toBe(true);
      expect(Number.isFinite(numericStringToNumber(s)), `${s} should coerce`).toBe(true);
    }
  });
});

describe("the writer's integer branch", () => {
  /**
   * `formatFloat` being right was not enough: the writer only called it for
   * NON-integers. It mirrored PHP's `is_float($v) ? number_format(...) :
   * (string) $v`, and the mirror leaks — a PHP int never stringifies
   * exponentially, but an integer-VALUED double does. `Number.isInteger(1e21)`
   * is true, so 1e21 took the String() branch and reached the sheet as
   * `<v>1e+21</v>`.
   *
   * Found by the numericHazards parity fixture, not by reasoning: the unit
   * tables above were both green while the two engines still disagreed.
   */
  it("does not write an integral double as exponential", () => {
    expect(cellXmlFor(1e21)).toContain("<v>1000000000000000000000</v>");
    expect(cellXmlFor(1e21)).not.toContain("e+");
    expect(cellXmlFor(1e300)).not.toContain("e+");
  });

  it("leaves ordinary integers alone", () => {
    expect(cellXmlFor(42)).toContain("<v>42</v>");
    expect(cellXmlFor(-7)).toContain("<v>-7</v>");
    expect(cellXmlFor(0)).toContain("<v>0</v>");
  });

  it("never writes a non-finite value into a cell", () => {
    for (const v of [NaN, Infinity, -Infinity]) {
      expect(cellXmlFor(v)).not.toMatch(/nan|inf/i);
    }
  });
});
