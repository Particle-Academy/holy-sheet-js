import { CellAddress } from "../workbook/cell-address";

/**
 * What a `columnWidths` entry may be — one rule for the validator, the repairer
 * and the normalizer, and the same rule as PHP `HolySheet\Schema\ColumnWidths`
 * (holy-sheet 2.3.4) and the Python port.
 *
 * - A KEY is a 0-based column index, 0 to {@link ColumnWidths.MAX_INDEX}
 *   (Excel's last column, XFD): a string of one to five digits. JavaScript object
 *   keys are always strings; an integer key from an array reads as its digits.
 * - A WIDTH is a non-negative finite number, or a string of digits with an
 *   optional decimal part (`"120"`, `"80.5"`).
 *
 * Before 2.4.3 this writer turned a key `"abc"` into `NaN` with `parseInt` and
 * wrote a NaN column; PHP overwrote column A and Python raised.
 */
export const ColumnWidths = {
  MAX_INDEX: 16383,

  index(key: unknown): number | null {
    if (typeof key === "number") {
      return Number.isInteger(key) && key >= 0 && key <= ColumnWidths.MAX_INDEX ? key : null;
    }
    if (typeof key === "string" && /^[0-9]{1,5}$/.test(key)) {
      const index = parseInt(key, 10);
      return index <= ColumnWidths.MAX_INDEX ? index : null;
    }
    return null;
  },

  width(px: unknown): number | null {
    if (typeof px === "number") {
      return Number.isFinite(px) && px >= 0 ? px : null;
    }
    if (typeof px === "string" && /^[0-9]+(\.[0-9]+)?$/.test(px)) {
      return Number(px);
    }
    return null;
  },

  /**
   * A column letter key (`"B"`, `"aa"`) as an index, for repair only. One or two
   * letters (A to ZZ): a longer run like `"abc"` is more likely a mistake than
   * column ABC, and a repair must not guess.
   */
  fromLetters(key: unknown): number | null {
    if (typeof key !== "string" || !/^[A-Za-z]{1,2}$/.test(key.trim())) return null;
    return CellAddress.index(key.trim());
  },
};
