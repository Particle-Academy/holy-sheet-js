import { isNumericString } from "../util";
import type { Cell } from "../workbook/cell";
import type { CellComment } from "../workbook/cell-comment";
import type { CellFormat } from "../workbook/cell-format";
import type { NormalizedChart, Sheet } from "../workbook/sheet";
import { CellAddress } from "../workbook/cell-address";
import type { Workbook } from "../workbook/workbook";
import { DateInverter } from "./format/date-inverter";

/**
 * A read `Workbook` -> the Holy Sheet schema `describe()` returns. Mirrors PHP
 * `Reader\WorkbookSchema`.
 *
 * Shared by every reader (xlsx, ods) so the output shape cannot drift between
 * formats: key order, which empty things are dropped, and how a date-formatted
 * serial becomes an ISO string are decided here and only here.
 */
export const WorkbookSchema = {
  fromWorkbook(workbook: Workbook): Record<string, unknown> {
    const schema: Record<string, unknown> = { sheets: workbook.sheets.map(sheetToSchema) };
    if (Object.keys(workbook.meta).length > 0) schema["meta"] = workbook.meta;
    // Only templates say so; a document read is unchanged.
    if (workbook.sourceKind === "template") schema["sourceKind"] = "template";
    return schema;
  },
};

function sheetToSchema(sheet: Sheet): Record<string, unknown> {
  const cells: Record<string, unknown> = {};
  for (const [address, cell] of sheet.cells) {
    const cellSchema = cellToSchema(cell);
    if (cellSchema === null) continue;
    cells[address] = cellSchema;
  }

  const out: Record<string, unknown> = { name: sheet.name, cells };
  if (sheet.mergedRegions.length > 0) {
    out["mergedRegions"] = sheet.mergedRegions.map((m) => ({ start: m.start, end: m.end }));
  }
  if (sheet.columnWidths.size > 0) {
    const widths: Record<number, number> = {};
    for (const [k, v] of sheet.columnWidths) widths[k] = v;
    out["columnWidths"] = widths;
  }
  if (sheet.frozenRows > 0) out["frozenRows"] = sheet.frozenRows;
  if (sheet.frozenCols > 0) out["frozenCols"] = sheet.frozenCols;
  if (sheet.rowHeights.size > 0) {
    const heights: Record<number, number> = {};
    for (const [k, v] of sheet.rowHeights) heights[k] = v;
    out["rowHeights"] = heights;
  }
  if (sheet.print !== null) {
    const print: Record<string, unknown> = {};
    if (sheet.print.orientation !== null) print["orientation"] = sheet.print.orientation;
    if (sheet.print.fitToWidth) print["fitToWidth"] = true;
    if (sheet.print.fitToPage) print["fitToPage"] = true;
    if (sheet.print.footer !== null) {
      const footer: Record<string, string> = {};
      for (const k of ["left", "center", "right"] as const) if (sheet.print.footer[k] !== "") footer[k] = sheet.print.footer[k];
      print["footer"] = footer;
    }
    out["print"] = print;
  }
  if (sheet.charts.length > 0) {
    out["charts"] = sheet.charts.map((c) => chartToSchema(c, sheet.name));
  }
  return out;
}

/** A normalised chart in the schema's form. Mirrors PHP WorkbookSchema::chartToSchema. */
function chartToSchema(chart: NormalizedChart, sheetName: string): Record<string, unknown> {
  const out: Record<string, unknown> = { type: chart.type };
  if (chart.title !== null) out["title"] = chart.title;
  if (chart.categories !== null) out["categories"] = chart.categories.toSchema(sheetName);
  out["series"] = chart.series.map((s) => {
    const series: Record<string, unknown> = {};
    if (s.name !== null) series["name"] = s.name;
    series["values"] = s.values.toSchema(sheetName);
    if (s.color !== null) series["color"] = "#" + s.color;
    if (s.points.length > 0) series["points"] = s.points.map((p) => ({ index: p.index, color: "#" + p.color }));
    return series;
  });
  out["position"] = {
    from: CellAddress.letter(chart.from[0]) + (chart.from[1] + 1),
    to: CellAddress.letter(chart.to[0]) + (chart.to[1] + 1),
  };
  if (chart.dataLabels) out["dataLabels"] = true;
  if (chart.legend !== null) out["legend"] = chart.legend;
  if (chart.format !== null) out["format"] = chart.format;
  return out;
}

function cellToSchema(cell: Cell): Record<string, unknown> | null {
  let value = cell.value;
  const format = cell.format;

  if (
    format !== null &&
    (format.displayFormat === "date" || format.displayFormat === "datetime") &&
    isNumericString(value as never)
  ) {
    value = DateInverter.toIso(Number(value), format.displayFormat === "datetime");
  }

  const out: Record<string, unknown> = { value };
  if (cell.formula !== null) out["formula"] = cell.formula;
  if (cell.cachedValue !== null) out["computedValue"] = cell.cachedValue;
  if (format !== null && !format.isEmpty()) out["format"] = formatToObject(format);
  if (cell.comment !== null) out["comment"] = commentToObject(cell.comment);

  if (Object.keys(out).length === 1 && out["value"] === null) return null;
  return out;
}

function formatToObject(f: CellFormat): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (f.bold) out["bold"] = true;
  if (f.italic) out["italic"] = true;
  if (f.textAlign !== null) out["textAlign"] = f.textAlign;
  if (f.displayFormat !== null) out["displayFormat"] = f.displayFormat;
  if (f.decimals !== null) out["decimals"] = f.decimals;
  if (f.color !== null) out["color"] = f.color;
  if (f.backgroundColor !== null) out["backgroundColor"] = f.backgroundColor;
  if (f.fontSize !== null) out["fontSize"] = f.fontSize;
  if (f.borderTop !== null) out["borderTop"] = f.borderTop;
  if (f.borderRight !== null) out["borderRight"] = f.borderRight;
  if (f.borderBottom !== null) out["borderBottom"] = f.borderBottom;
  if (f.borderLeft !== null) out["borderLeft"] = f.borderLeft;
  if (f.currency !== null) out["currency"] = f.currency;
  return out;
}

function commentToObject(c: CellComment): Record<string, unknown> {
  const out: Record<string, unknown> = { text: c.text };
  if (c.author !== null) out["author"] = c.author;
  if (c.color !== null) out["color"] = c.color;
  return out;
}
