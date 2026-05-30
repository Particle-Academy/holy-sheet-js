import { isNumericString } from "../util";
import type { Cell } from "../workbook/cell";
import type { CellComment } from "../workbook/cell-comment";
import type { CellFormat } from "../workbook/cell-format";
import type { Sheet } from "../workbook/sheet";
import { Workbook } from "../workbook/workbook";
import { unzipSync } from "../zip";
import { CommentsParser } from "./comments-parser";
import { DateInverter } from "./format/date-inverter";
import { RelsParser } from "./rels-parser";
import { SharedStringsParser } from "./shared-strings-parser";
import { StylesParser } from "./styles-parser";
import { WorksheetParser } from "./worksheet-parser";
import { at, el, els, parseXml } from "./xml";

const decoder = new TextDecoder();

/** XLSX reader — bytes → Holy Sheet schema. Mirrors PHP `Reader\XlsxReader`. */
export class XlsxReader {
  describe(bytes: Uint8Array): Record<string, unknown> {
    return this.workbookToSchema(this.readWorkbook(bytes));
  }

  readWorkbook(bytes: Uint8Array): Workbook {
    const files = unzipSync(bytes);
    const getText = (name: string): string | null => (files[name] ? decoder.decode(files[name]) : null);

    const stylesXml = getText("xl/styles.xml");
    const stylesIndex = stylesXml ? StylesParser.parse(stylesXml) : [];

    const sharedStringsXml = getText("xl/sharedStrings.xml");
    const sharedStrings = sharedStringsXml ? SharedStringsParser.parse(sharedStringsXml) : [];

    const workbookXml = getText("xl/workbook.xml");
    if (workbookXml === null) throw new Error("[holy-sheet] missing xl/workbook.xml");

    const workbookRelsXml = getText("xl/_rels/workbook.xml.rels");
    const workbookRels = workbookRelsXml ? RelsParser.parse(workbookRelsXml) : {};

    const workbookDoc = parseXml(workbookXml);
    if (!workbookDoc) throw new Error("[holy-sheet] failed to parse xl/workbook.xml");

    const sheets: Sheet[] = [];
    let i = 0;
    for (const sheetEl of els(el(workbookDoc, "sheets"), "sheet")) {
      const name = at(sheetEl, "name") ?? "";
      const rId = at(sheetEl, "id") ?? ""; // r:id → local "id"
      const target = workbookRels[rId]?.Target;
      if (target == null) continue;

      const sheetPath = "xl/" + target.replace(/^\//, "");
      const worksheetXml = getText(sheetPath);
      if (worksheetXml === null) continue;

      const sheetNum = i + 1;
      const sheetRelsXml = getText(`xl/worksheets/_rels/sheet${sheetNum}.xml.rels`);
      let comments: Record<string, CellComment> = {};
      if (sheetRelsXml) {
        const sheetRels = RelsParser.parse(sheetRelsXml);
        for (const cr of Object.values(RelsParser.byType(sheetRels, "/comments"))) {
          const commentsPath = resolveRelativePath(sheetPath, cr.Target);
          const commentsXml = getText(commentsPath);
          if (commentsXml) comments = { ...comments, ...CommentsParser.parse(commentsXml) };
        }
      }

      sheets.push(WorksheetParser.parse(worksheetXml, name, stylesIndex, comments, sharedStrings));
      i++;
    }

    const meta = this.parseDocProps(getText("docProps/core.xml"));
    return new Workbook(sheets, meta);
  }

  private parseDocProps(coreXml: string | null): Record<string, unknown> {
    if (coreXml === null) return {};
    const core = parseXml(coreXml);
    if (!core) return {};
    const meta: Record<string, unknown> = {};
    const creator = el(core, "creator");
    if (creator) meta["creator"] = creator.text;
    const created = el(core, "created");
    if (created) meta["created"] = created.text;
    return meta;
  }

  private workbookToSchema(workbook: Workbook): Record<string, unknown> {
    const schema: Record<string, unknown> = { sheets: workbook.sheets.map((s) => this.sheetToSchema(s)) };
    if (Object.keys(workbook.meta).length > 0) schema["meta"] = workbook.meta;
    return schema;
  }

  private sheetToSchema(sheet: Sheet): Record<string, unknown> {
    const cells: Record<string, unknown> = {};
    for (const [address, cell] of sheet.cells) {
      const cellSchema = this.cellToSchema(cell);
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
    return out;
  }

  private cellToSchema(cell: Cell): Record<string, unknown> | null {
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

function resolveRelativePath(base: string, target: string): string {
  const baseDir = base.includes("/") ? base.slice(0, base.lastIndexOf("/")) : ".";
  const combined = baseDir + "/" + target;
  const parts: string[] = [];
  for (const segment of combined.split("/")) {
    if (segment === "..") parts.pop();
    else if (segment !== "" && segment !== ".") parts.push(segment);
  }
  return parts.join("/");
}
