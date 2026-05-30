import type { CellFormat } from "../workbook/cell-format";
import { xmlEscape } from "../xml";
import { NumFmtBuilder } from "./format/num-fmt-builder";

interface Xf {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  align: string | null;
}
interface FontRec {
  size: number;
  name: string;
  color: string | null;
  bold: boolean;
  italic: boolean;
}
interface BorderRec {
  top: string | null;
  right: string | null;
  bottom: string | null;
  left: string | null;
}

/**
 * Deduplicates fonts/fills/borders/numFmts and the cellXfs that combine them.
 * Mirrors PHP `Writer\StylesRegistry` (style index 0 = the unformatted base).
 */
export class StylesRegistry {
  private xfIndex = new Map<string, number>([["__default__", 0]]);
  private xfs: Xf[] = [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, align: null }];

  private fontIndex = new Map<string, number>([["__default__", 0]]);
  private fonts: FontRec[] = [{ size: 11, name: "Calibri", color: null, bold: false, italic: false }];

  private fillIndex = new Map<string, number>([
    ["__none__", 0],
    ["__gray125__", 1],
  ]);
  private fills: ({ type: "none" } | { type: "gray125" } | { type: "solid"; fg: string })[] = [
    { type: "none" },
    { type: "gray125" },
  ];

  private borderIndex = new Map<string, number>([["__default__", 0]]);
  private borders: BorderRec[] = [{ top: null, right: null, bottom: null, left: null }];

  private numFmtIndex = new Map<string, number>();
  private numFmts = new Map<number, string>();
  private nextNumFmtId = 164;

  register(format: CellFormat | null): number {
    if (format === null || format.isEmpty()) return 0;

    const key = format.key();
    const existing = this.xfIndex.get(key);
    if (existing !== undefined) return existing;

    const fontId = this.fontFor(format);
    const fillId = this.fillFor(format);
    const borderId = this.borderFor(format);
    const numFmtId = this.numFmtFor(format);

    this.xfs.push({ fontId, fillId, borderId, numFmtId, align: format.textAlign });
    const idx = this.xfs.length - 1;
    this.xfIndex.set(key, idx);
    return idx;
  }

  private fontFor(f: CellFormat): number {
    const rec: FontRec = {
      size: f.fontSize ?? 11,
      name: "Calibri",
      color: f.color,
      bold: f.bold,
      italic: f.italic,
    };
    const key = JSON.stringify(rec);
    const existing = this.fontIndex.get(key);
    if (existing !== undefined) return existing;
    this.fonts.push(rec);
    const idx = this.fonts.length - 1;
    this.fontIndex.set(key, idx);
    return idx;
  }

  private fillFor(f: CellFormat): number {
    if (f.backgroundColor === null) return 0;
    const key = f.backgroundColor.toUpperCase();
    const existing = this.fillIndex.get(key);
    if (existing !== undefined) return existing;
    this.fills.push({ type: "solid", fg: key });
    const idx = this.fills.length - 1;
    this.fillIndex.set(key, idx);
    return idx;
  }

  private borderFor(f: CellFormat): number {
    if (!f.borderTop && !f.borderRight && !f.borderBottom && !f.borderLeft) return 0;
    const rec: BorderRec = {
      top: f.borderTop,
      right: f.borderRight,
      bottom: f.borderBottom,
      left: f.borderLeft,
    };
    const key = JSON.stringify(rec);
    const existing = this.borderIndex.get(key);
    if (existing !== undefined) return existing;
    this.borders.push(rec);
    const idx = this.borders.length - 1;
    this.borderIndex.set(key, idx);
    return idx;
  }

  private numFmtFor(f: CellFormat): number {
    const code = NumFmtBuilder.build(f);
    if (code === null) return 0;
    const existing = this.numFmtIndex.get(code);
    if (existing !== undefined) return existing;
    const id = this.nextNumFmtId++;
    this.numFmts.set(id, code);
    this.numFmtIndex.set(code, id);
    return id;
  }

  toXml(): string {
    let numFmtsXml = "";
    if (this.numFmts.size > 0) {
      let records = "";
      for (const [id, code] of this.numFmts) {
        records += `<numFmt numFmtId="${id}" formatCode="${xmlEscape(code)}"/>`;
      }
      numFmtsXml = `<numFmts count="${this.numFmts.size}">${records}</numFmts>`;
    }

    let fontsXml = `<fonts count="${this.fonts.length}">`;
    for (const f of this.fonts) {
      fontsXml += "<font>";
      fontsXml += `<sz val="${Number(f.size)}"/>`;
      if (f.bold) fontsXml += "<b/>";
      if (f.italic) fontsXml += "<i/>";
      if (f.color) fontsXml += `<color rgb="${hexToArgb(f.color)}"/>`;
      fontsXml += `<name val="${xmlEscape(f.name)}"/>`;
      fontsXml += "</font>";
    }
    fontsXml += "</fonts>";

    let fillsXml = `<fills count="${this.fills.length}">`;
    for (const fill of this.fills) {
      if (fill.type === "none") {
        fillsXml += '<fill><patternFill patternType="none"/></fill>';
      } else if (fill.type === "gray125") {
        fillsXml += '<fill><patternFill patternType="gray125"/></fill>';
      } else {
        fillsXml += `<fill><patternFill patternType="solid"><fgColor rgb="${hexToArgb(fill.fg)}"/></patternFill></fill>`;
      }
    }
    fillsXml += "</fills>";

    let bordersXml = `<borders count="${this.borders.length}">`;
    for (const b of this.borders) {
      bordersXml += "<border>";
      for (const side of ["left", "right", "top", "bottom"] as const) {
        const color = b[side];
        if (!color) {
          bordersXml += `<${side}/>`;
        } else {
          bordersXml += `<${side} style="thin"><color rgb="${hexToArgb(color)}"/></${side}>`;
        }
      }
      bordersXml += "<diagonal/></border>";
    }
    bordersXml += "</borders>";

    let cellXfsXml = `<cellXfs count="${this.xfs.length}">`;
    for (const xf of this.xfs) {
      let apply = "";
      if (xf.fontId > 0) apply += ' applyFont="1"';
      if (xf.fillId > 0) apply += ' applyFill="1"';
      if (xf.borderId > 0) apply += ' applyBorder="1"';
      if (xf.numFmtId > 0) apply += ' applyNumberFormat="1"';
      if (xf.align) apply += ' applyAlignment="1"';
      cellXfsXml += `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0"${apply}>`;
      if (xf.align) cellXfsXml += `<alignment horizontal="${xmlEscape(xf.align)}"/>`;
      cellXfsXml += "</xf>";
    }
    cellXfsXml += "</cellXfs>";

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      numFmtsXml +
      fontsXml +
      fillsXml +
      bordersXml +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      cellXfsXml +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
      "</styleSheet>"
    );
  }
}

/** "#RRGGBB" → "FFRRGGBB" (Excel ARGB). */
function hexToArgb(hex: string): string {
  const h = hex.replace(/^#/, "");
  if (h.length === 6) return "FF" + h.toUpperCase();
  if (h.length === 8) return h.toUpperCase();
  return "FF000000";
}
