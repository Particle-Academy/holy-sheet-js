import type { Sheet } from "./sheet";

/** Internal canonical workbook value object. Mirrors PHP `Workbook\Workbook`. */
export class Workbook {
  constructor(
    readonly sheets: Sheet[],
    readonly meta: Record<string, unknown> = {},
    /** What a reader found: an .xltx / .ots reads exactly like its .xlsx / .ods twin and is marked `template`. */
    readonly sourceKind: "document" | "template" = "document",
  ) {}
}
