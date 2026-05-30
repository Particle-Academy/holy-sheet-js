import { CellFormat } from "../workbook/cell-format";
import { NumFmtParser, type NumFmtParsed } from "./format/num-fmt-parser";
import { at, el, els, parseXml, type XmlNode } from "./xml";

interface FontRec {
  bold: boolean;
  italic: boolean;
  size: number;
  color: string | null;
}
interface BorderRec {
  top: string | null;
  right: string | null;
  bottom: string | null;
  left: string | null;
}

/** Parses xl/styles.xml → (CellFormat | null)[] by xf id. Mirrors PHP `StylesParser`. */
export const StylesParser = {
  parse(stylesXml: string): (CellFormat | null)[] {
    const xml = parseXml(stylesXml);
    if (!xml) return [];

    const numFmts = parseNumFmts(xml);
    const fonts = parseFonts(xml);
    const fills = parseFills(xml);
    const borders = parseBorders(xml);

    const xfs: (CellFormat | null)[] = [];
    for (const xf of els(el(xml, "cellXfs"), "xf")) {
      const fontId = xf.attrs["fontId"] !== undefined ? parseInt(xf.attrs["fontId"], 10) : 0;
      const fillId = xf.attrs["fillId"] !== undefined ? parseInt(xf.attrs["fillId"], 10) : 0;
      const borderId = xf.attrs["borderId"] !== undefined ? parseInt(xf.attrs["borderId"], 10) : 0;
      const numFmtId = xf.attrs["numFmtId"] !== undefined ? parseInt(xf.attrs["numFmtId"], 10) : 0;

      const alignment = el(xf, "alignment");
      const align = alignment && alignment.attrs["horizontal"] !== undefined ? alignment.attrs["horizontal"] : null;

      const font = fonts[fontId] ?? null;
      const fill = fills[fillId] ?? null;
      const border = borders[borderId] ?? null;
      const customCode = numFmts[numFmtId];
      const numFmtParsed =
        customCode !== undefined ? NumFmtParser.parse(customCode) : NumFmtParser.parseBuiltin(numFmtId);

      if (xfs.length === 0 && font === null && fill === null && border === null && !numFmtParsed && align === null) {
        xfs.push(null);
        continue;
      }

      xfs.push(buildCellFormat(font, fill, border, numFmtParsed, align));
    }

    return xfs;
  },
};

function parseNumFmts(xml: XmlNode): Record<number, string> {
  const map: Record<number, string> = {};
  for (const nf of els(el(xml, "numFmts"), "numFmt")) {
    map[parseInt(nf.attrs["numFmtId"] ?? "0", 10)] = nf.attrs["formatCode"] ?? "";
  }
  return map;
}

function parseFonts(xml: XmlNode): FontRec[] {
  return els(el(xml, "fonts"), "font").map((f) => {
    const sz = el(f, "sz");
    const color = el(f, "color");
    return {
      bold: !!el(f, "b"),
      italic: !!el(f, "i"),
      size: sz ? Math.trunc(Number(at(sz, "val"))) : 11,
      color: color && at(color, "rgb") !== undefined ? argbToHex(at(color, "rgb")!) : null,
    };
  });
}

function parseFills(xml: XmlNode): (string | null)[] {
  return els(el(xml, "fills"), "fill").map((fill) => {
    const pf = el(fill, "patternFill");
    const pattern = pf ? (at(pf, "patternType") ?? "") : "none";
    if (pattern !== "solid") return null;
    const fg = el(pf, "fgColor");
    return fg && at(fg, "rgb") !== undefined ? argbToHex(at(fg, "rgb")!) : null;
  });
}

function parseBorders(xml: XmlNode): BorderRec[] {
  return els(el(xml, "borders"), "border").map((b) => {
    const rec: BorderRec = { top: null, right: null, bottom: null, left: null };
    for (const side of ["top", "right", "bottom", "left"] as const) {
      const sideEl = el(b, side);
      const color = el(sideEl, "color");
      if (sideEl && at(sideEl, "style") !== undefined && at(sideEl, "style") !== "" && color && at(color, "rgb") !== undefined) {
        rec[side] = argbToHex(at(color, "rgb")!);
      }
    }
    return rec;
  });
}

function buildCellFormat(
  font: FontRec | null,
  fill: string | null,
  border: BorderRec | null,
  numFmt: NumFmtParsed | null,
  align: string | null,
): CellFormat {
  return new CellFormat({
    bold: font?.bold ?? false,
    italic: font?.italic ?? false,
    textAlign: align,
    displayFormat: numFmt?.displayFormat ?? null,
    decimals: numFmt?.decimals ?? null,
    color: font?.color ?? null,
    backgroundColor: fill,
    fontSize: font && font.size !== 11 ? font.size : null,
    borderTop: border?.top ?? null,
    borderRight: border?.right ?? null,
    borderBottom: border?.bottom ?? null,
    borderLeft: border?.left ?? null,
    currency: numFmt?.currency ?? null,
  });
}

function argbToHex(argb: string): string {
  const a = argb.toUpperCase();
  if (a.length === 8) return "#" + a.slice(2);
  if (a.length === 6) return "#" + a;
  return "#000000";
}
