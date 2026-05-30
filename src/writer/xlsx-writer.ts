import { Cell } from "../workbook/cell";
import { CellAddress } from "../workbook/cell-address";
import type { Sheet } from "../workbook/sheet";
import type { Workbook } from "../workbook/workbook";
import { xmlEscape } from "../xml";
import { formatFloat } from "../util";
import { zipSync, type ZipFile } from "../zip";
import { StylesRegistry } from "./styles-registry";

const encoder = new TextEncoder();

/**
 * XLSX writer — emits a complete OOXML SpreadsheetML package as bytes.
 * Mirrors PHP `Writer\XlsxWriter` (file I/O moved to the Agent for isomorphism).
 */
export class XlsxWriter {
  toBytes(workbook: Workbook): Uint8Array {
    const styles = new StylesRegistry();

    // Pre-render sheet xml so styles register before we serialize styles.xml.
    const sheetXmls: string[] = workbook.sheets.map((sheet) => this.sheetXml(sheet, styles));

    const sheetsWithComments: number[] = [];
    workbook.sheets.forEach((sheet, i) => {
      if (sheet.hasComments()) sheetsWithComments.push(i);
    });

    const files: ZipFile[] = [];
    const add = (name: string, xml: string): void => {
      files.push({ name, data: encoder.encode(xml) });
    };

    add("[Content_Types].xml", this.contentTypesXml(workbook, sheetsWithComments));
    add("_rels/.rels", this.rootRelsXml());
    add("xl/workbook.xml", this.workbookXml(workbook));
    add("xl/_rels/workbook.xml.rels", this.workbookRelsXml(workbook));
    add("xl/styles.xml", styles.toXml());
    add("docProps/core.xml", this.coreXml(workbook));
    add("docProps/app.xml", this.appXml());

    workbook.sheets.forEach((sheet, i) => {
      const n = i + 1;
      add(`xl/worksheets/sheet${n}.xml`, sheetXmls[i]!);
      if (sheet.hasComments()) {
        add(`xl/worksheets/_rels/sheet${n}.xml.rels`, this.sheetRelsXml(n));
        add(`xl/comments${n}.xml`, this.commentsXml(sheet));
        add(`xl/drawings/vmlDrawing${n}.vml`, this.vmlDrawingXml(sheet));
      }
    });

    return zipSync(files);
  }

  private contentTypesXml(workbook: Workbook, sheetsWithComments: number[]): string {
    let overrides =
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';

    workbook.sheets.forEach((_, i) => {
      const n = i + 1;
      overrides += `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    });

    for (const i of sheetsWithComments) {
      const n = i + 1;
      overrides += `<Override PartName="/xl/comments${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`;
    }

    let defaults =
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>';
    if (sheetsWithComments.length > 0) {
      defaults += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
    }

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      defaults +
      overrides +
      "</Types>"
    );
  }

  private rootRelsXml(): string {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
      "</Relationships>"
    );
  }

  private workbookXml(workbook: Workbook): string {
    let sheets = "";
    workbook.sheets.forEach((sheet, i) => {
      const name = this.escape(sheet.name);
      const sheetId = i + 1;
      const rId = "rId" + (i + 1);
      sheets += `<sheet name="${name}" sheetId="${sheetId}" r:id="${rId}"/>`;
    });
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<sheets>${sheets}</sheets>` +
      "</workbook>"
    );
  }

  private workbookRelsXml(workbook: Workbook): string {
    let rels = "";
    workbook.sheets.forEach((_, i) => {
      const rId = "rId" + (i + 1);
      rels += `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`;
    });
    rels += `<Relationship Id="rId${workbook.sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      rels +
      "</Relationships>"
    );
  }

  private sheetRelsXml(sheetNum: number): string {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${sheetNum}.vml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments${sheetNum}.xml"/>` +
      "</Relationships>"
    );
  }

  private sheetXml(sheet: Sheet, styles: StylesRegistry): string {
    let sheetViewsXml = "";
    if (sheet.frozenRows > 0 || sheet.frozenCols > 0) {
      const topLeft = CellAddress.letter(sheet.frozenCols) + (sheet.frozenRows + 1);
      const pane =
        "<pane" +
        (sheet.frozenCols > 0 ? ` xSplit="${sheet.frozenCols}"` : "") +
        (sheet.frozenRows > 0 ? ` ySplit="${sheet.frozenRows}"` : "") +
        ` topLeftCell="${topLeft}" activePane="bottomRight" state="frozen"/>`;
      sheetViewsXml = `<sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>`;
    }

    let colsXml = "";
    if (sheet.columnWidths.size > 0) {
      colsXml = "<cols>";
      for (const [colIdx, px] of sheet.columnWidths) {
        const excelWidth = Math.max(1.0, (px - 5) / 7);
        const colNum = colIdx + 1;
        colsXml += `<col min="${colNum}" max="${colNum}" width="${excelWidth.toFixed(4)}" customWidth="1"/>`;
      }
      colsXml += "</cols>";
    }

    let rowsXml = "";
    for (const [rowIndex, row] of sheet.rows()) {
      let cellsXml = "";
      const sortedCols = [...row.keys()].sort();
      for (const col of sortedCols) {
        cellsXml += this.cellXml(row.get(col)!, styles);
      }
      rowsXml += `<row r="${rowIndex}">${cellsXml}</row>`;
    }

    let mergesXml = "";
    if (sheet.mergedRegions.length > 0) {
      mergesXml = `<mergeCells count="${sheet.mergedRegions.length}">`;
      for (const merge of sheet.mergedRegions) mergesXml += `<mergeCell ref="${merge.ref()}"/>`;
      mergesXml += "</mergeCells>";
    }

    const legacyDrawing = sheet.hasComments() ? '<legacyDrawing r:id="rId1"/>' : "";

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      sheetViewsXml +
      colsXml +
      `<sheetData>${rowsXml}</sheetData>` +
      mergesXml +
      legacyDrawing +
      "</worksheet>"
    );
  }

  private cellXml(cell: Cell, styles: StylesRegistry): string {
    const ref = cell.address;
    const type = cell.excelType();
    const styleIdx = styles.register(cell.format);
    const sAttr = styleIdx > 0 ? ` s="${styleIdx}"` : "";

    if (cell.formula !== null) {
      const f = this.escape(cell.formula.replace(/^=+/, ""));
      const cached = cell.cachedValue ?? cell.value;
      const v = cached !== null ? `<v>${this.escape(String(cached))}</v>` : "";
      return `<c r="${ref}"${sAttr}><f>${f}</f>${v}</c>`;
    }

    if (type === "inlineStr") {
      const t = this.escape(String(cell.value));
      return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${t}</t></is></c>`;
    }

    if (type === "b") {
      const v = cell.value === true ? "1" : "0";
      return `<c r="${ref}"${sAttr} t="b"><v>${v}</v></c>`;
    }

    if (cell.value === null) {
      return `<c r="${ref}"${sAttr}/>`;
    }

    let v: string;
    if (typeof cell.value === "number") {
      v = Number.isInteger(cell.value) ? String(cell.value) : formatFloat(cell.value);
    } else {
      v = String(cell.value);
    }
    if (v === "") v = "0";
    return `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
  }

  private commentsXml(sheet: Sheet): string {
    const authors: string[] = [];
    let commentList = "";
    for (const entry of sheet.comments()) {
      const author = entry.comment.author ?? "Author";
      if (!authors.includes(author)) authors.push(author);
      const authorIdx = authors.indexOf(author);
      const text = this.escape(entry.comment.text);
      commentList +=
        `<comment ref="${entry.address}" authorId="${authorIdx}">` +
        `<text><r><t xml:space="preserve">${text}</t></r></text>` +
        "</comment>";
    }
    let authorsXml = "<authors>";
    for (const a of authors) authorsXml += `<author>${this.escape(a)}</author>`;
    authorsXml += "</authors>";
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      authorsXml +
      `<commentList>${commentList}</commentList>` +
      "</comments>"
    );
  }

  private vmlDrawingXml(sheet: Sheet): string {
    let shapes = "";
    let shapeId = 1024;
    for (const entry of sheet.comments()) {
      const m = /^([A-Z]+)(\d+)$/.exec(entry.address)!;
      const col = CellAddress.index(m[1]!);
      const row = parseInt(m[2]!, 10) - 1;
      shapeId++;
      shapes +=
        `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t202" ` +
        'style="position:absolute;margin-left:60pt;margin-top:5pt;width:108pt;height:60pt;z-index:1;visibility:hidden" ' +
        'fillcolor="#ffffe1" o:insetmode="auto">' +
        '<v:fill color2="#ffffe1"/>' +
        '<v:shadow on="t" color="black" obscured="t"/>' +
        '<v:path o:connecttype="none"/>' +
        '<v:textbox><div style="text-align:left"></div></v:textbox>' +
        '<x:ClientData ObjectType="Note">' +
        "<x:MoveWithCells/>" +
        "<x:SizeWithCells/>" +
        `<x:Anchor>${col + 1}, 15, ${row}, 10, ${col + 3}, 31, ${row + 4}, 18</x:Anchor>` +
        "<x:AutoFill>False</x:AutoFill>" +
        `<x:Row>${row}</x:Row>` +
        `<x:Column>${col}</x:Column>` +
        "</x:ClientData>" +
        "</v:shape>";
    }
    return (
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" ' +
      'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:x="urn:schemas-microsoft-com:office:excel">' +
      '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
      '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
      '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
      shapes +
      "</xml>"
    );
  }

  private coreXml(workbook: Workbook): string {
    const now = (workbook.meta["created"] as string | undefined) ?? gmDateZ();
    const creator = this.escape(String(workbook.meta["creator"] ?? "Holy Sheet"));
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:creator>${creator}</dc:creator>` +
      `<cp:lastModifiedBy>${creator}</cp:lastModifiedBy>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
      "</cp:coreProperties>"
    );
  }

  private appXml(): string {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      "<Application>Holy Sheet</Application>" +
      "<DocSecurity>0</DocSecurity>" +
      "<ScaleCrop>false</ScaleCrop>" +
      "<SharedDoc>false</SharedDoc>" +
      "<HyperlinksChanged>false</HyperlinksChanged>" +
      "<AppVersion>1.0</AppVersion>" +
      "</Properties>"
    );
  }

  private escape(s: string): string {
    return xmlEscape(s);
  }
}

/** PHP gmdate('Y-m-d\TH:i:s\Z') — current UTC time, seconds precision. */
function gmDateZ(): string {
  return new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
}
