import type { Sheet } from "./sheet";

/** Internal canonical workbook value object. Mirrors PHP `Workbook\Workbook`. */
export class Workbook {
  constructor(
    readonly sheets: Sheet[],
    readonly meta: Record<string, unknown> = {},
  ) {}
}
