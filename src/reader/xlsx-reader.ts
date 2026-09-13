import type { CellComment } from "../workbook/cell-comment";
import type { Sheet } from "../workbook/sheet";
import { Workbook } from "../workbook/workbook";
import { unzipSync } from "../zip";
import { CommentsParser } from "./comments-parser";
import { RelsParser } from "./rels-parser";
import { SharedStringsParser } from "./shared-strings-parser";
import { StylesParser } from "./styles-parser";
import { WorkbookSchema } from "./workbook-schema";
import { WorksheetParser } from "./worksheet-parser";
import { at, el, els, parseXml } from "./xml";

const decoder = new TextDecoder();

/** XLSX reader — bytes → Holy Sheet schema. Mirrors PHP `Reader\XlsxReader`. */
export class XlsxReader {
  describe(bytes: Uint8Array): Record<string, unknown> {
    return WorkbookSchema.fromWorkbook(this.readWorkbook(bytes));
  }

  readWorkbook(bytes: Uint8Array): Workbook {
    return this.readFiles(unzipSync(bytes));
  }

  /** Read an already-unzipped package (name -> bytes), so a caller that sniffed it does not unzip twice. */
  readFiles(files: Record<string, Uint8Array>): Workbook {
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
