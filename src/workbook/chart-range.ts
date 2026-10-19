import { CellAddress } from "./cell-address";

/**
 * A cell range a chart reads from. Mirrors PHP `Workbook\ChartRange`:
 * "B5:B7", "B5", "'Q3 Revenue'!B5:B7" or "Data!B5:B7", resolved against the
 * chart's own sheet when no sheet is named.
 */
export class ChartRange {
  constructor(
    readonly sheet: string,
    /** 0-based first column. */
    readonly c1: number,
    /** 1-based first row. */
    readonly r1: number,
    readonly c2: number,
    readonly r2: number,
  ) {}

  static parse(ref: unknown, defaultSheet: string): ChartRange | null {
    if (typeof ref !== "string") return null;
    let text = ref.trim();
    let sheet = defaultSheet;
    const bang = text.lastIndexOf("!");
    if (bang >= 0) {
      let sheetPart = text.slice(0, bang);
      text = text.slice(bang + 1);
      if (sheetPart.length >= 2 && sheetPart.startsWith("'") && sheetPart.endsWith("'")) {
        sheetPart = sheetPart.slice(1, -1).replace(/''/g, "'");
      }
      if (sheetPart === "") return null;
      sheet = sheetPart;
    }
    const ends = text.replace(/\$/g, "").split(":");
    if (ends.length > 2) return null;
    const a = CellAddress.parse(ends[0]!);
    const b = CellAddress.parse(ends[1] ?? ends[0]!);
    if (a === null || b === null || a[1] < 1 || b[1] < 1) return null;
    return new ChartRange(sheet, Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1]));
  }

  /** 'Q3 Revenue'!$B$5:$B$7 */
  formula(): string {
    const from = "$" + CellAddress.letter(this.c1) + "$" + this.r1;
    const to = "$" + CellAddress.letter(this.c2) + "$" + this.r2;
    return "'" + this.sheet.replace(/'/g, "''") + "'!" + from + (from === to ? "" : ":" + to);
  }

  /** "B5:B7" on its own sheet, "'Data'!B5:B7" on another. */
  toSchema(ownSheet: string): string {
    const from = CellAddress.letter(this.c1) + this.r1;
    const to = CellAddress.letter(this.c2) + this.r2;
    const ref = from === to ? from : from + ":" + to;
    return this.sheet === ownSheet ? ref : "'" + this.sheet.replace(/'/g, "''") + "'!" + ref;
  }

  /** A1 addresses, row by row. */
  addresses(): string[] {
    const out: string[] = [];
    for (let r = this.r1; r <= this.r2; r++) {
      for (let c = this.c1; c <= this.c2; c++) out.push(CellAddress.letter(c) + r);
    }
    return out;
  }
}
