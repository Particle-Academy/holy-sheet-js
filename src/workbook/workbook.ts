import type { Sheet } from "./sheet";

/** Internal canonical workbook value object. Mirrors PHP `Workbook\Workbook`. */
export class Workbook {
  constructor(
    readonly sheets: Sheet[],
    readonly meta: Record<string, unknown> = {},
    /** What a reader found: an .xltx / .ots reads exactly like its .xlsx / .ods twin and is marked `template`. */
    readonly sourceKind: "document" | "template" = "document",
    /**
     * Whether the file counts its date serials from 1904-01-01
     * (`<workbookPr date1904="1"/>`). Read only: the writer always writes the
     * 1900 system, and describe() turns every date into an ISO string.
     */
    readonly date1904: boolean = false,
  ) {}
}
