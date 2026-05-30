import { Cell } from "../workbook/cell";
import type { CellComment } from "../workbook/cell-comment";
import type { CellFormat } from "../workbook/cell-format";
import type { CellPrimitive } from "../schema/types";
import { MergedRegion } from "../workbook/merged-region";
import { Sheet } from "../workbook/sheet";
import { at, el, els, parseXml, type XmlNode } from "./xml";

/** Parses one xl/worksheets/sheetN.xml → Sheet. Mirrors PHP `WorksheetParser`. */
export const WorksheetParser = {
  parse(
    worksheetXml: string,
    name: string,
    stylesIndex: (CellFormat | null)[],
    comments: Record<string, CellComment> = {},
    sharedStrings: string[] = [],
  ): Sheet {
    const xml = parseXml(worksheetXml);
    if (!xml) return new Sheet(name);

    const cells = parseCells(xml, stylesIndex, comments, sharedStrings);
    const merges = parseMerges(xml);
    const columnWidths = parseColumnWidths(xml);
    const [frozenRows, frozenCols] = parseFrozen(xml);

    return new Sheet(name, cells, merges, columnWidths, frozenRows, frozenCols);
  },
};

function parseCells(
  xml: XmlNode,
  stylesIndex: (CellFormat | null)[],
  comments: Record<string, CellComment>,
  sharedStrings: string[],
): Map<string, Cell> {
  const cells = new Map<string, Cell>();
  const sheetData = el(xml, "sheetData");
  if (!sheetData) return cells;

  for (const row of els(sheetData, "row")) {
    for (const c of els(row, "c")) {
      const address = c.attrs["r"] ?? "";
      if (address === "") continue;

      const type = c.attrs["t"] ?? "n";
      const styleIdx = c.attrs["s"] !== undefined ? parseInt(c.attrs["s"], 10) : 0;
      const format = stylesIndex[styleIdx] ?? null;

      let formula: string | null = null;
      let cachedValue: CellPrimitive = null;
      let value: CellPrimitive = null;

      const f = el(c, "f");
      const v = el(c, "v");
      if (f) {
        formula = f.text;
        cachedValue = v ? coerceValue(v.text, type) : null;
      } else if (type === "inlineStr" && el(c, "is") && el(el(c, "is"), "t")) {
        value = el(el(c, "is"), "t")!.text;
      } else if (type === "s" && v) {
        const idx = parseInt(v.text, 10);
        value = sharedStrings[idx] ?? "";
      } else if (v) {
        value = coerceValue(v.text, type);
      }

      const comment = comments[address] ?? null;
      cells.set(address, new Cell(address, value, formula, format, comment, cachedValue));
    }
  }
  return cells;
}

function coerceValue(raw: string, type: string): CellPrimitive {
  if (type === "b") return raw === "1" || raw === "true" || raw === "TRUE";
  if (type === "str") return raw;
  if (type === "inlineStr") return raw;
  if (raw === "") return null;
  if (/^-?\d+$/.test(raw)) return parseInt(raw, 10);
  if (/^\s*[+-]?(\d+(\.\d+)?|\.\d+)([eE][+-]?\d+)?$/.test(raw)) return parseFloat(raw);
  return raw;
}

function parseMerges(xml: XmlNode): MergedRegion[] {
  const out: MergedRegion[] = [];
  for (const m of els(el(xml, "mergeCells"), "mergeCell")) {
    const ref = m.attrs["ref"] ?? "";
    if (ref.includes(":")) {
      const [start, end] = ref.split(":", 2) as [string, string];
      out.push(new MergedRegion(start, end));
    }
  }
  return out;
}

function parseColumnWidths(xml: XmlNode): Map<number, number> {
  const out = new Map<number, number>();
  for (const col of els(el(xml, "cols"), "col")) {
    if (col.attrs["customWidth"] === undefined) continue;
    const min = parseInt(col.attrs["min"] ?? "0", 10);
    const max = parseInt(col.attrs["max"] ?? "0", 10);
    const excelWidth = Number(col.attrs["width"] ?? "0");
    const px = excelWidth * 7 + 5;
    for (let i = min; i <= max; i++) out.set(i - 1, Math.round(px));
  }
  return out;
}

function parseFrozen(xml: XmlNode): [number, number] {
  for (const view of els(el(xml, "sheetViews"), "sheetView")) {
    const pane = el(view, "pane");
    if (!pane) continue;
    const rows = at(pane, "ySplit") !== undefined ? parseInt(at(pane, "ySplit")!, 10) : 0;
    const cols = at(pane, "xSplit") !== undefined ? parseInt(at(pane, "xSplit")!, 10) : 0;
    return [rows, cols];
  }
  return [0, 0];
}
