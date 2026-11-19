import type { CellComment } from "../workbook/cell-comment";
import type { NormalizedChart, Sheet } from "../workbook/sheet";
import { ChartParser } from "./chart-parser";
import { Workbook } from "../workbook/workbook";
import { unzipSync } from "../zip";
import { filesView, sourceKind } from "../zip/office-template";
import { SheetNotFoundException } from "../exceptions";
import { CommentsParser } from "./comments-parser";
import { ReadLimits } from "./read-limits";
import { RelsParser } from "./rels-parser";
import { SharedStringsParser } from "./shared-strings-parser";
import { StylesParser } from "./styles-parser";
import { WorkbookSchema } from "./workbook-schema";
import { WorksheetParser } from "./worksheet-parser";
import { at, decodePart, el, els, parseXml, type XmlNode } from "./xml";


/** XLSX reader — bytes → Holy Sheet schema. Mirrors PHP `Reader\XlsxReader`. */
export class XlsxReader {
  describe(bytes: Uint8Array): Record<string, unknown> {
    return WorkbookSchema.fromWorkbook(this.readWorkbook(bytes));
  }

  readWorkbook(bytes: Uint8Array): Workbook {
    return this.readFiles(unzipSync(bytes));
  }

  /**
   * Read an already-unzipped package (name -> bytes), so a caller that sniffed
   * it does not unzip twice. `only` reads just the sheet of that exact name
   * (the others are never parsed); an unknown name throws
   * SheetNotFoundException. Mirrors PHP `XlsxReader::readArchive`.
   */
  readFiles(files: Record<string, Uint8Array>, only: string | null = null, limits: ReadLimits | null = null): Workbook {
    limits ??= ReadLimits.from();
    const kind = sourceKind(filesView(files));
    const getText = (name: string): string | null => (files[name] ? decodePart(files[name]) : null);

    const stylesXml = getText("xl/styles.xml");
    const stylesIndex = stylesXml ? StylesParser.parse(stylesXml) : [];

    const sharedStringsXml = getText("xl/sharedStrings.xml");
    const sharedStrings = sharedStringsXml ? SharedStringsParser.parse(sharedStringsXml, limits) : [];

    const workbookXml = getText("xl/workbook.xml");
    if (workbookXml === null) throw new Error("[holy-sheet] missing xl/workbook.xml");

    const workbookRelsXml = getText("xl/_rels/workbook.xml.rels");
    const workbookRels = workbookRelsXml ? RelsParser.parse(workbookRelsXml) : {};

    const workbookDoc = parseXml(workbookXml);
    if (!workbookDoc) throw new Error("[holy-sheet] failed to parse xl/workbook.xml");
    // The 1904 date system: `<workbookPr date1904="1"/>` (or "true").
    const date1904 = ["1", "true"].includes(at(el(workbookDoc, "workbookPr"), "date1904") ?? "");

    if (only !== null) {
      const names = els(el(workbookDoc, "sheets"), "sheet").map((s) => at(s, "name") ?? "");
      if (!names.includes(only)) throw new SheetNotFoundException(only, names);
    }

    const sheets: Sheet[] = [];
    // The position in workbook.xml of each described sheet: activeTab counts
    // the <sheet> elements, and a sheet whose part is missing is skipped
    // below, so the two orders can differ.
    const positions: number[] = [];
    let position = -1;
    let i = 0;
    for (const sheetEl of els(el(workbookDoc, "sheets"), "sheet")) {
      position++;
      const name = at(sheetEl, "name") ?? "";
      if (only !== null && name !== only) {
        // The sheet rels part is named by the position among READ sheets, so a
        // skipped sheet advances it exactly when a full read would have.
        const skipped = workbookRels[at(sheetEl, "id") ?? ""]?.Target;
        if (skipped != null && files["xl/" + skipped.replace(/^\//, "")]) i++;
        continue;
      }
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

      let sheet = WorksheetParser.parse(worksheetXml, name, stylesIndex, comments, sharedStrings, limits);

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
      positions.push(position);
      i++;
    }

    const meta = this.parseDocProps(getText("docProps/core.xml"));
    // A read narrowed to one sheet is not the workbook, so it names no active sheet.
    const active = only === null && sheets.length > 0 ? activeSheet(workbookDoc, positions, position) : null;
    return new Workbook(sheets, meta, kind, date1904, active);
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

/**
 * The described sheet that opens: the first `<workbookView>`'s `activeTab`,
 * clamped to the `<sheet>` elements workbook.xml lists. Absent, or anything but
 * a non-negative integer, is the default, 0. An active sheet the reader skipped
 * (its part is missing) falls back to the first described sheet. Mirrors PHP
 * `XlsxReader::activeSheet` (CC-fancy-0050).
 */
function activeSheet(workbook: XmlNode, positions: number[], lastPosition: number): number {
  let tab = 0;
  const view = els(el(workbook, "bookViews"), "workbookView")[0];
  if (view !== undefined) {
    // xsd:unsignedInt collapses whitespace; only the four XML spaces, as PHP's trim set.
    const raw = (at(view, "activeTab") ?? "").replace(/^[ \t\n\r]+|[ \t\n\r]+$/g, "");
    if (/^[0-9]+$/.test(raw)) {
      // A digit string past any integer becomes a huge number, and the clamp
      // takes it to the last sheet, as PHP's saturating cast does.
      tab = Math.min(Number(raw), Math.max(0, lastPosition));
    }
  }
  const index = positions.indexOf(tab);
  return index < 0 ? 0 : index;
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
