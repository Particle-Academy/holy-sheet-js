/** Shared low-level helpers that mirror PHP's loose numeric semantics. */

/** PHP `is_numeric` (approx): optional leading whitespace, sign, int/float, exponent. */
export function isNumericString(v: unknown): boolean {
  if (typeof v === "number") return Number.isFinite(v);
  if (typeof v !== "string") return false;
  return /^\s*[+-]?(\d+(\.\d+)?|\.\d+)([eE][+-]?\d+)?$/.test(v);
}

/** Coerce a numeric string the way PHP does: `(float)` when it has a dot, else `(int)`. */
export function numericStringToNumber(v: string): number {
  return v.includes(".") ? parseFloat(v) : parseInt(v, 10);
}

/**
 * Serialize a float the way the PHP writer does:
 * `rtrim(rtrim(number_format($v, 14, '.', ''), '0'), '.')`.
 */
export function formatFloat(v: number): string {
  let s = v.toFixed(14);
  s = s.replace(/0+$/, "").replace(/\.$/, "");
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
