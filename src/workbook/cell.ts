import type { CellPrimitive } from "../schema/types";
import type { CellComment } from "./cell-comment";
import type { CellFormat } from "./cell-format";

export type ExcelType = "str" | "inlineStr" | "b" | "n";

/** A single cell. Mirrors PHP `Workbook\Cell`. */
export class Cell {
  /** Canonical: no leading "=". See canonicalFormula(). */
  readonly formula: string | null;

  constructor(
    readonly address: string,
    readonly value: CellPrimitive,
    formula: string | null = null,
    readonly format: CellFormat | null = null,
    readonly comment: CellComment | null = null,
    readonly cachedValue: CellPrimitive = null,
  ) {
    this.formula = formula === null ? null : Cell.canonicalFormula(formula);
  }

  /**
   * The one normalisation of formula text: every leading "=" is stripped, so
   * "=A2*2" and "A2*2" are the same formula. Mirrors PHP
   * `Cell::canonicalFormula()`. It lives where every Cell is made, so the
   * writer and the linter see the same text; until 2.14.2 only the writer
   * stripped it, and lint reported "=A2*2" as a syntax error.
   */
  static canonicalFormula(formula: string): string {
    return formula.replace(/^=+/, "");
  }

  excelType(): ExcelType {
    if (this.formula !== null) return "str";
    if (typeof this.value === "string") return "inlineStr";
    if (typeof this.value === "boolean") return "b";
    return "n";
  }
}
