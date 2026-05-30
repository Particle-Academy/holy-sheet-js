import type { CellPrimitive } from "../schema/types";
import type { CellComment } from "./cell-comment";
import type { CellFormat } from "./cell-format";

export type ExcelType = "str" | "inlineStr" | "b" | "n";

/** A single cell. Mirrors PHP `Workbook\Cell`. */
export class Cell {
  constructor(
    readonly address: string,
    readonly value: CellPrimitive,
    readonly formula: string | null = null,
    readonly format: CellFormat | null = null,
    readonly comment: CellComment | null = null,
    readonly cachedValue: CellPrimitive = null,
  ) {}

  excelType(): ExcelType {
    if (this.formula !== null) return "str";
    if (typeof this.value === "string") return "inlineStr";
    if (typeof this.value === "boolean") return "b";
    return "n";
  }
}
