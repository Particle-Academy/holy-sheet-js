import type { CellComment } from "../workbook/cell-comment";
import type { NormalizedChart, Sheet } from "../workbook/sheet";
import { ChartParser } from "./chart-parser";
import { Workbook } from "../workbook/workbook";
import { unzipSync } from "../zip";
import { filesView, sourceKind } from "../zip/office-template";
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
    const kind = sourceKind(filesView(files));
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

      let sheet = WorksheetParser.parse(worksheetXml, name, stylesIndex, comments, sharedStrings);

      // Charts: sheet rels -> drawing -> drawing rels -> chart parts (see PHP).
      if (sheetRelsXml) {
        const charts: NormalizedChart[] = [];
        for (const dr of Object.values(RelsParser.byType(RelsParser.parse(sheetRelsXml), "/drawing"))) {
          const drawingPath = resolveRelativePath(sheetPath, dr.Target);
          const slash = drawingPath.lastIndexOf("/");
          const drawingXml = getText(drawingPath);
          const drawingRelsXml = getText(drawingPath.slice(0, slash) + "/_rels/" + drawingPath.slice(slash + 1) + ".rels");
          if (drawingXml === null || drawingRelsXml === null) continue;
          const drawingRels = RelsParser.parse(drawingRelsXml);
          for (const anchor of ChartParser.anchors(drawingXml)) {
            const target = drawingRels[anchor.rel]?.Target;
            const chartXml = typeof target === "string" ? getText(resolveRelativePath(drawingPath, target)) : null;
            const chart = chartXml !== null ? ChartParser.chart(chartXml, name, anchor.from, anchor.to) : null;
            if (chart !== null) charts.push(chart);
          }
        }
        if (charts.length > 0) sheet = sheet.withCharts(charts);
      }

      sheets.push(sheet);
      i++;
    }

    const meta = this.parseDocProps(getText("docProps/core.xml"));
    return new Workbook(sheets, meta, kind);
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
