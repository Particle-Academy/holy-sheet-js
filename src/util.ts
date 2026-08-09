/** Shared low-level helpers that mirror PHP's loose numeric semantics. */

/** PHP `is_numeric` (approx): optional leading whitespace, sign, int/float, exponent. */
export function isNumericString(v: unknown): boolean {
  if (typeof v === "number") return Number.isFinite(v);
  if (typeof v !== "string") return false;
  return /^\s*[+-]?(\d+(\.\d+)?|\.\d+)([eE][+-]?\d+)?$/.test(v);
}

/**
 * Coerce a numeric string the way PHP does.
 *
 * This used to be `v.includes(".") ? parseFloat(v) : parseInt(v, 10)`, on the
 * theory that it mirrored PHP's `(float)`/`(int)`. It does not. `parseInt`
 * stops at the first character that is not a digit, so it reads "1e5" as 1 and
 * throws the exponent away, while PHP has parsed exponent notation in numeric
 * strings since PHP 7 and says 100000. Every value written as "1e5", "2E-3" or
 * "1.5e3" therefore landed in the sheet as a different number depending on
 * which backend wrote it.
 *
 * `Number` is what the rest of this package already uses for exactly this
 * (see `FormulaLinter.coerceNumber`) — the two disagreed with each other, not
 * just with PHP.
 *
 * The float path is deliberate for integers too: PHP's `(int)` cast CLAMPS at
 * PHP_INT_MAX, so `(int) "1e21"` is 9223372036854775807 — a number nobody
 * wrote. Both engines now take the float path and both say 1e21. The emitted
 * `<v>` is unchanged for integral values, because `formatFloat` trims the
 * fraction away.
 */
export function numericStringToNumber(v: string): number {
  return Number(v);
}

/**
 * Serialize a float the way the PHP writer does:
 * `rtrim(rtrim(number_format($v, 14, '.', ''), '0'), '.')`.
 *
 * `toFixed` alone cannot do this. It switches to exponential notation at 1e21,
 * and the trailing-zero trim then chewed on the EXPONENT rather than a
 * fraction:
 *
 *   (1e300).toFixed(14)          === "1e+300"
 *   "1e+300".replace(/0+$/, "")  === "1e+3"
 *
 * so 1e300 was written to the sheet as 1e3. Not a crash and not invalid XML —
 * simply a different number, silently, in a file someone opens later and
 * believes. At 1e21 exactly it produced `<v>1e+21</v>`, which is not valid cell
 * content at all. PHP's `number_format` has no exponential mode and was right
 * throughout.
 *
 * Above 1e21 every double is an integer, so `BigInt` gives the exact digits
 * PHP prints — verified against it, including that 1e300 expands to
 * 1000000000000000052504760… and not to a 1 with 300 zeros.
 */
export function formatFloat(v: number): string {
  // NaN and Infinity have no `<v>` representation; "NaN" in a cell is a
  // corrupt sheet. PHP's number_format yields 0 for both, so match it.
  if (!Number.isFinite(v)) return "0";

  let s = Math.abs(v) >= 1e21 ? BigInt(v).toString() : v.toFixed(14);

  // Only ever trim a FRACTION. The old unconditional /0+$/ is what turned an
  // exponent into a smaller exponent, and it would equally have eaten the
  // trailing zeros of an integer like 1200 had toFixed not always emitted a
  // decimal point.
  if (s.includes(".")) s = s.replace(/0+$/, "").replace(/\.$/, "");

  return s === "" || s === "-0" ? "0" : s;
}

/** PHP-style debug type name, used by the validator's `got` field. */
export function typeOf(value: unknown): string {
  if (value === null || value === undefined) return "null";
  if (Array.isArray(value)) return "array";
  switch (typeof value) {
    case "object":
      return "object";
    case "boolean":
      return "bool";
    case "number":
      return Number.isInteger(value) ? "int" : "float";
    case "string":
      return "string";
    default:
      return typeof value;
  }
}

/**
 * True for a plain object map (sparse cells / CellData), false for arrays/primitives.
 * Narrows to `Record<string, any>` because the input is loose agent JSON — the
 * Validator, not the type system, is the gate.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function isPlainObject(value: unknown): value is Record<string, any> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
