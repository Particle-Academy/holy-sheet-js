import { CellAddress } from "../workbook/cell-address";

/**
 * The few PHP array and cast semantics the op reducer and diff depend on, in one
 * place so `sheet-reducer.ts` and `sheet-diff.ts` read like the PHP they mirror.
 *
 * PHP has one array type for lists and maps; a JSON object and a JSON array both
 * decode to it. `isArr` is PHP's `is_array` over parsed JSON: any object or array.
 * A key whose value is `undefined` is treated as absent, as JSON would.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type Any = any;
export type Obj = Record<string, Any>;

/** PHP `is_array`: a JSON object or a JSON array. */
export function isArr(value: unknown): value is Obj {
  return typeof value === "object" && value !== null;
}

/** PHP `array_key_exists` (a null value counts; `undefined` is absent). */
export function has(value: unknown, key: string): boolean {
  return isArr(value) && Object.prototype.hasOwnProperty.call(value, key) && value[key] !== undefined;
}

/** PHP `$value[$key] ?? null`. */
export function get(value: unknown, key: string): Any {
  return has(value, key) ? ((value as Obj)[key] ?? null) : null;
}

/** PHP `$value === []`: an empty array or an empty object. */
export function isEmptyArr(value: unknown): boolean {
  return isArr(value) && Object.keys(value).every((k) => value[k] === undefined);
}

/** PHP `array_values`. */
export function valuesOf(value: unknown): Any[] {
  if (Array.isArray(value)) return value.slice();
  if (isArr(value)) return Object.keys(value).filter((k) => value[k] !== undefined).map((k) => value[k]);
  return [];
}

/** `foreach ($value as $key => $item)`. */
export function entriesOf(value: unknown): [string, Any][] {
  if (Array.isArray(value)) return value.map((item, i) => [String(i), item]);
  if (isArr(value)) return Object.keys(value).filter((k) => value[k] !== undefined).map((k) => [k, value[k]]);
  return [];
}

/** PHP `(int)` cast. */
export function phpInt(value: unknown): number {
  if (value === null || value === undefined) return 0;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "number") return Number.isFinite(value) ? Math.trunc(value) || 0 : 0;
  if (typeof value === "string") {
    const m = /^[ \t\n\r\v\f]*[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/.exec(value);
    if (!m) return 0;
    const n = Number(m[0].trim());
    return Number.isFinite(n) ? Math.trunc(n) || 0 : 0;
  }
  if (isArr(value)) return isEmptyArr(value) ? 0 : 1;
  return 0;
}

/** PHP `(string)` cast. */
export function phpString(value: unknown): string {
  if (value === null || value === undefined || value === false) return "";
  if (value === true) return "1";
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (isArr(value)) return "Array";
  return String(value);
}

/** PHP `trim` with its default set: space, tab, LF, CR, NUL and vertical tab. Not `String#trim`, which strips NBSP and more. */
export function phpTrim(value: string): string {
  return value.replace(/^[ \t\n\r\x00\x0B]+|[ \t\n\r\x00\x0B]+$/g, "");
}

/**
 * PHP `SheetReducer::integer()`: an int, or a string of digits (`ctype_digit`);
 * anything else is null.
 *
 * One divergence cannot be removed: PHP decodes `2.0` as a float and refuses it,
 * but `JSON.parse` gives the number 2, which is accepted here. An integer past
 * PHP's int range is refused, as PHP would have decoded it as a float too.
 */
export function phpInteger(value: unknown): number | null {
  if (typeof value === "number") {
    return Number.isInteger(value) && Math.abs(value) <= 2 ** 63 ? value || 0 : null;
  }
  return typeof value === "string" && /^[0-9]+$/.test(value) ? phpInt(value) : null;
}

/**
 * Whether PHP reads `$key` of a `foreach` over `columnWidths` as a column index:
 * an int key (a canonical decimal string PHP stores as an int, `-1` included) or
 * a string of digits (`ctype_digit`, so `"007"`). `"abc"`, `"1.5"`, `""` and
 * `"-0"` are not.
 */
export function isIndexKey(key: string): boolean {
  return /^[0-9]+$/.test(key) || (/^-[1-9][0-9]*$/.test(key) && BigInt(key) >= -(2n ** 63n));
}

/** PHP `strtoupper` (ASCII only since PHP 8.2, unlike `toUpperCase`). */
export function asciiUpper(value: string): string {
  return value.replace(/[a-z]+/g, (s) => s.toUpperCase());
}

/**
 * PHP `CellAddress::parse`: `[columnIndex, rowNumber]` or null. Written out
 * rather than reusing this port's `CellAddress.parse`, whose `trim` and
 * `toUpperCase` are Unicode-aware where PHP's are ASCII.
 */
export function parseAddress(address: string): [number, number] | null {
  const m = /^([A-Z]+)(\d+)$/.exec(asciiUpper(phpTrim(address)));
  if (!m) return null;
  return [CellAddress.index(m[1]!), parseInt(m[2]!, 10)];
}

export const letter = CellAddress.letter;
