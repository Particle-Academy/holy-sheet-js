import { Cell } from "../workbook/cell";
import type { CellComment } from "../workbook/cell-comment";
import type { CellFormat } from "../workbook/cell-format";
import type { CellPrimitive } from "../schema/types";
import { MergedRegion } from "../workbook/merged-region";
import { Sheet, type PrintSetup } from "../workbook/sheet";
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

    return new Sheet(name, cells, merges, columnWidths, frozenRows, frozenCols, parseRowHeights(xml), parsePrint(xml));
  },
};

/** Rows with a CUSTOM height, in pixels (points / 0.75). Mirrors PHP WorksheetParser::parseRowHeights. */
function parseRowHeights(xml: XmlNode): Map<number, number> {
  const out = new Map<number, number>();
  const sheetData = el(xml, "sheetData");
  if (!sheetData) return out;
  for (const row of els(sheetData, "row")) {
    const r = parseInt(at(row, "r") ?? "0", 10);
    const custom = at(row, "customHeight");
    const ht = at(row, "ht");
    if (r > 0 && ht !== undefined && (custom === "1" || custom === "true")) {
      out.set(r, Math.round((parseFloat(ht) / 0.75) * 100) / 100);
    }
  }
  return out;
}

/** Print setup. Mirrors PHP WorksheetParser::parsePrint. */
function parsePrint(xml: XmlNode): PrintSetup | null {
  const fitToPage = ["1", "true"].includes(at(el(el(xml, "sheetPr"), "pageSetUpPr"), "fitToPage") ?? "");
  const setup = el(xml, "pageSetup");
  const o = at(setup, "orientation");
  const orientation = o === "portrait" || o === "landscape" ? o : null;
  const fitToWidth =
    fitToPage && setup !== undefined && at(setup, "fitToWidth") === "1" && ["0", ""].includes(at(setup, "fitToHeight") ?? "");
  const oddFooter = el(el(xml, "headerFooter"), "oddFooter");
  const footer = oddFooter !== undefined ? parseFooterCodes(oddFooter.text) : null;
  if (orientation === null && !fitToWidth && footer === null) return null;
  return { orientation, fitToWidth, footer };
}

/**
 * Excel header/footer codes back to three slots. Mirrors PHP
 * WorksheetParser::parseFooterCodes, code for code.
 */
export function parseFooterCodes(codes: string): PrintSetup["footer"] {
  const slots: Record<"left" | "center" | "right", string> = { left: "", center: "", right: "" };
  let current: "left" | "center" | "right" = "center";
  const len = codes.length;
  for (let i = 0; i < len; i++) {
    const ch = codes[i]!;
    if (ch !== "&" || i + 1 >= len) {
      slots[current] += ch;
      continue;
    }
    const next = codes[++i]!;
    if (next === "&") slots[current] += "&";
    else if (next === "L") current = "left";
    else if (next === "C") current = "center";
    else if (next === "R") current = "right";
    else if (next === "P") slots[current] += "{page}";
    else if (next === "N") slots[current] += "{pages}";
    else if (next === '"') {
      const close = codes.indexOf('"', i + 1);
      i = close === -1 ? len : close;
    } else if (/[0-9]/.test(next)) {
      while (i + 1 < len && /[0-9]/.test(codes[i + 1]!)) i++;
    } else if (next === "K") {
      i = Math.min(len - 1, i + 6);
    }
  }
  return slots.left === "" && slots.center === "" && slots.right === "" ? null : slots;
}

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
