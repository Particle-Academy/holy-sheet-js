import type { Cell } from "./cell";
import type { CellComment } from "./cell-comment";
import type { MergedRegion } from "./merged-region";
import type { ChartRange } from "./chart-range";

/** How a sheet prints. Mirrors the PHP `Sheet::$print` array. */
export interface PrintSetup {
  orientation: "portrait" | "landscape" | null;
  fitToWidth: boolean;
  /** One page wide AND tall. */
  fitToPage: boolean;
  footer: { left: string; center: string; right: string } | null;
}

/** A native chart, normalised. Mirrors the PHP `Sheet::$charts` entries. */
export interface NormalizedChart {
  type: "column" | "bar" | "line" | "pie";
  title: string | null;
  categories: ChartRange | null;
  series: { name: string | null; values: ChartRange; color: string | null; points: { index: number; color: string }[] }[];
  /** [col, row], both 0-based. */
  from: [number, number];
  to: [number, number];
  legend: boolean | null;
  dataLabels: boolean;
  format: string | null;
}

/** A worksheet. Mirrors PHP `Workbook\Sheet`. Cells are an A1-keyed map. */
export class Sheet {
  constructor(
    readonly name: string,
    readonly cells: Map<string, Cell> = new Map(),
    readonly mergedRegions: MergedRegion[] = [],
    readonly columnWidths: Map<number, number> = new Map(),
    readonly frozenRows: number = 0,
    readonly frozenCols: number = 0,
    /** 1-based row number → pixels. */
    readonly rowHeights: Map<number, number> = new Map(),
    readonly print: PrintSetup | null = null,
    readonly charts: NormalizedChart[] = [],
  ) {}

  withCharts(charts: NormalizedChart[]): Sheet {
    return new Sheet(this.name, this.cells, this.mergedRegions, this.columnWidths, this.frozenRows, this.frozenCols, this.rowHeights, this.print, charts);
  }

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
