import type { CellComment } from "../workbook/cell-comment";
import type { NormalizedChart, Sheet } from "../workbook/sheet";
import { ChartParser } from "./chart-parser";
import { Workbook } from "../workbook/workbook";
import { unzipSync } from "../zip";
import { filesView, sourceKind } from "../zip/office-template";
import { MissingPartException, SheetNotFoundException } from "../exceptions";
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

    const workbookPath = "xl/workbook.xml";
    const workbookXml = getText(workbookPath);
    if (workbookXml === null) throw new Error("[holy-sheet] missing xl/workbook.xml");

    const workbookRelsXml = getText(RelsParser.relsPathOf(workbookPath));
    const workbookRels = workbookRelsXml ? RelsParser.parse(workbookRelsXml) : {};

    /** A part a relationship points at. Missing, it throws: a part skipped silently is how a workbook read as having no sheets. */
    const requirePart = (part: string, source: string, relId: string): string => {
      const text = getText(part);
      if (text === null) throw new MissingPartException(part, source, relId);
      return text;
    };

    // Styles and shared strings are wherever the workbook's relationships say;
    // the conventional names are only the fallback for a package that declares
    // neither. Matched on the Type's last segment, so transitional and strict
    // namespaces both match. Mirrors PHP `XlsxReader::workbookPart`.
    const workbookPart = (type: string, fallback: string): string | null => {
      for (const [id, rel] of Object.entries(workbookRels)) {
        if (!RelsParser.isExternal(rel) && rel.Type.endsWith("/" + type)) {
          return requirePart(RelsParser.resolve(workbookPath, rel.Target), workbookPath, id);
        }
      }
      return getText(fallback);
    };

    const stylesXml = workbookPart("styles", "xl/styles.xml");
    const stylesIndex = stylesXml ? StylesParser.parse(stylesXml) : [];

    const sharedStringsXml = workbookPart("sharedStrings", "xl/sharedStrings.xml");
    const sharedStrings = sharedStringsXml ? SharedStringsParser.parse(sharedStringsXml, limits) : [];

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
    // the <sheet> elements. Every listed sheet is described (a missing part
    // throws), but the mapping is kept explicit.
    const positions: number[] = [];
    let position = -1;
    for (const sheetEl of els(el(workbookDoc, "sheets"), "sheet")) {
      position++;
      const name = at(sheetEl, "name") ?? "";
      if (only !== null && name !== only) continue;
      const rId = at(sheetEl, "id") ?? ""; // r:id → local "id"
      const sheetRel = workbookRels[rId];
      if (sheetRel === undefined || RelsParser.isExternal(sheetRel)) throw new MissingPartException(null, workbookPath, rId);

      const sheetPath = RelsParser.resolve(workbookPath, sheetRel.Target);
      const worksheetXml = requirePart(sheetPath, workbookPath, rId);

      // The sheet's own rels part sits beside it, named after its actual part,
      // not after its position in the tab strip (sheet3.xml can be the first tab).
      const sheetRelsXml = getText(RelsParser.relsPathOf(sheetPath));
      const sheetRels = sheetRelsXml ? RelsParser.parse(sheetRelsXml) : {};

      let comments: Record<string, CellComment> = {};
      for (const [cId, cr] of Object.entries(RelsParser.byType(sheetRels, "/comments"))) {
        const commentsXml = requirePart(RelsParser.resolve(sheetPath, cr.Target), sheetPath, cId);
        comments = { ...comments, ...CommentsParser.parse(commentsXml) };
      }

      let sheet = WorksheetParser.parse(worksheetXml, name, stylesIndex, comments, sharedStrings, limits);

      // Charts: sheet rels -> drawing -> drawing rels -> chart parts (see PHP).
      const charts: NormalizedChart[] = [];
      for (const [dId, dr] of Object.entries(RelsParser.byType(sheetRels, "/drawing"))) {
        const drawingPath = RelsParser.resolve(sheetPath, dr.Target);
        const drawingXml = requirePart(drawingPath, sheetPath, dId);
        // A drawing with no rels part holds no chart; that is not a missing part.
        const drawingRelsXml = getText(RelsParser.relsPathOf(drawingPath));
        if (drawingRelsXml === null) continue;
        const drawingRels = RelsParser.parse(drawingRelsXml);
        for (const anchor of ChartParser.anchors(drawingXml)) {
          const chartRel = drawingRels[anchor.rel];
          if (chartRel === undefined || RelsParser.isExternal(chartRel)) continue;
          const chartXml = requirePart(RelsParser.resolve(drawingPath, chartRel.Target), drawingPath, anchor.rel);
          const chart = ChartParser.chart(chartXml, name, anchor.from, anchor.to);
          if (chart !== null) charts.push(chart);
        }
      }
      if (charts.length > 0) sheet = sheet.withCharts(charts);

      sheets.push(sheet);
      positions.push(position);
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
 * a non-negative integer, is the default, 0. A tab no described sheet sits at
 * falls back to the first described sheet. Mirrors PHP
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
