import type { Cell } from "./cell";
import type { CellComment } from "./cell-comment";
import type { MergedRegion } from "./merged-region";

/** A worksheet. Mirrors PHP `Workbook\Sheet`. Cells are an A1-keyed map. */
export class Sheet {
  constructor(
    readonly name: string,
    readonly cells: Map<string, Cell> = new Map(),
    readonly mergedRegions: MergedRegion[] = [],
    readonly columnWidths: Map<number, number> = new Map(),
    readonly frozenRows: number = 0,
    readonly frozenCols: number = 0,
  ) {}

  /** rowNumber → (colLetter → Cell), sorted by ascending row number. */
  rows(): Map<number, Map<string, Cell>> {
    const rows = new Map<number, Map<string, Cell>>();
    for (const [address, cell] of this.cells) {
      const m = /^([A-Z]+)(\d+)$/.exec(address);
      if (!m) continue;
      const rowNum = parseInt(m[2]!, 10);
      let row = rows.get(rowNum);
      if (!row) {
        row = new Map();
        rows.set(rowNum, row);
      }
      row.set(m[1]!, cell);
    }
    return new Map([...rows.entries()].sort((a, b) => a[0] - b[0]));
  }

  hasComments(): boolean {
    for (const c of this.cells.values()) if (c.comment !== null) return true;
    return false;
  }

  comments(): { address: string; comment: CellComment }[] {
    const out: { address: string; comment: CellComment }[] = [];
    for (const c of this.cells.values()) {
      if (c.comment !== null) out.push({ address: c.address, comment: c.comment });
    }
    return out;
  }
}
