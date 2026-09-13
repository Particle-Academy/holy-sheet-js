import { CellFormat } from "../../workbook/cell-format";
import type { DisplayFormat } from "../../schema/types";
import { SYMBOL_TO_ISO } from "../format/num-fmt-parser";
import type { XmlNode } from "../xml";
import { NS, asciiLower, attr, child, childrenIn, ownText, phpInt, phpIsNumeric, phpTrim } from "./ns";

interface DataFormat {
  displayFormat: DisplayFormat;
  decimals?: number;
  currency?: string;
}

interface Resolved {
  bold: boolean;
  italic: boolean;
  textAlign: string | null;
  color: string | null;
  backgroundColor: string | null;
  fontSize: number | null;
  borderTop: string | null;
  borderRight: string | null;
  borderBottom: string | null;
  borderLeft: string | null;
  data: DataFormat | null;
}

/**
 * Cell styles and data styles -> `CellFormat`. Mirrors PHP `Reader\Ods\OdsStyles`,
 * where what is and is not mapped is documented.
 *
 * A cell's style is found in content.xml's automatic styles, then styles.xml's
 * common styles, and inherits along `style:parent-style-name` down to the
 * table-cell `style:default-style`. The nearest style that sets a property wins.
 */
export class OdsStyles {
  private readonly automatic = new Map<string, XmlNode>();
  private readonly common = new Map<string, XmlNode>();
  private defaultStyle: XmlNode | null = null;
  private readonly dataStyles = new Map<string, XmlNode>();
  private readonly resolved = new Map<string, Resolved>();
  private readonly defaultFontSize: number | null;

  constructor(content: XmlNode | null, styles: XmlNode | null) {
    if (content) this.index(child(content, NS.OFFICE, "automatic-styles"), this.automatic);
    if (styles) {
      this.index(child(styles, NS.OFFICE, "styles"), this.common);
      this.index(child(styles, NS.OFFICE, "automatic-styles"), this.common);
    }
    this.defaultFontSize = this.properties("Default").fontSize;
  }

  /**
   * The format of one cell.
   *
   * @param styleName  the cell's effective style (cell, row default, column default, or "Default")
   * @param valueType  office:value-type
   * @param currency   office:currency, the ISO code a currency value carries
   * @param dateHasTime whether an office:date-value carries a time of day
   */
  cellFormat(styleName: string, valueType: string | null, currency: string | null, dateHasTime: boolean): CellFormat | null {
    const p = this.properties(styleName);
    let data: DataFormat | null = p.data;
    const display = data?.displayFormat ?? null;

    // The value type says what the value IS; the data style only says how it is shown.
    switch (valueType) {
      case "date":
        data = { displayFormat: display === "date" || display === "datetime" ? display : dateHasTime ? "datetime" : "date" };
        break;
      case "time":
        data = { displayFormat: "datetime" };
        break;
      case "percentage":
        if (display !== "percentage") {
          data = { displayFormat: "percentage", ...(data?.decimals !== undefined ? { decimals: data.decimals } : {}) };
        }
        break;
      case "currency":
        if (display !== "currency") {
          data = { displayFormat: "currency", ...(data?.decimals !== undefined ? { decimals: data.decimals } : {}) };
        } else {
          data = { ...data! };
        }
        if (currency !== null && currency !== "") data.currency = currency;
        break;
    }

    const format = new CellFormat({
      bold: p.bold,
      italic: p.italic,
      textAlign: p.textAlign,
      displayFormat: data?.displayFormat ?? null,
      decimals: data?.decimals ?? null,
      color: p.color,
      backgroundColor: p.backgroundColor,
      fontSize: p.fontSize !== null && p.fontSize !== this.defaultFontSize ? p.fontSize : null,
      borderTop: p.borderTop,
      borderRight: p.borderRight,
      borderBottom: p.borderBottom,
      borderLeft: p.borderLeft,
      currency: data?.currency ?? null,
    });

    return format.isEmpty() ? null : format;
  }

  private index(container: XmlNode | undefined, into: Map<string, XmlNode>): void {
    if (!container) return;

    for (const el of childrenIn(container, NS.STYLE)) {
      const family = attr(el, NS.STYLE, "family") ?? "";
      if (el.name === "default-style" && family === "table-cell") {
        this.defaultStyle ??= el;
      } else if (el.name === "style" && family === "table-cell") {
        const name = attr(el, NS.STYLE, "name") ?? "";
        if (!into.has(name)) into.set(name, el);
      }
    }
    for (const el of childrenIn(container, NS.NUMBER)) {
      const name = attr(el, NS.STYLE, "name") ?? "";
      if (!this.dataStyles.has(name)) this.dataStyles.set(name, el);
    }
  }

  /** The style and its ancestors, nearest first, ending with the default style. */
  private chain(name: string): XmlNode[] {
    const chain: XmlNode[] = [];
    const seen = new Set<string>();
    // The cell's own style is normally automatic; a parent is always common.
    let el = this.automatic.get(name) ?? this.common.get(name);
    while (el !== undefined && !seen.has(name)) {
      seen.add(name);
      chain.push(el);
      name = attr(el, NS.STYLE, "parent-style-name") ?? "";
      el = name === "" ? undefined : this.common.get(name) ?? this.automatic.get(name);
    }
    if (this.defaultStyle) chain.push(this.defaultStyle);
    return chain;
  }

  private properties(name: string): Resolved {
    const cached = this.resolved.get(name);
    if (cached) return cached;

    const chain = this.chain(name);
    const weight = first(chain, "text-properties", NS.FO, "font-weight");
    const style = first(chain, "text-properties", NS.FO, "font-style");
    const size = first(chain, "text-properties", NS.FO, "font-size");

    let dataStyle: string | null = null;
    for (const el of chain) {
      const candidate = attr(el, NS.STYLE, "data-style-name") ?? "";
      if (candidate !== "") {
        dataStyle = candidate;
        break;
      }
    }

    const sizeMatch = size !== null ? /^([0-9]*\.?[0-9]+)pt$/.exec(size) : null;
    const resolved: Resolved = {
      bold: weight !== null && (weight === "bold" || (phpIsNumeric(weight) && phpInt(weight) >= 600)),
      italic: style === "italic" || style === "oblique",
      textAlign: textAlign(chain),
      color: color(chain),
      backgroundColor: background(chain),
      fontSize: sizeMatch ? Math.trunc(parseFloat(sizeMatch[1]!)) : null,
      borderTop: border(chain, "top"),
      borderRight: border(chain, "right"),
      borderBottom: border(chain, "bottom"),
      borderLeft: border(chain, "left"),
      data: this.dataFormat(dataStyle, 0),
    };
    this.resolved.set(name, resolved);
    return resolved;
  }

  /** A data style -> displayFormat, decimals, currency. */
  private dataFormat(name: string | null, depth: number): DataFormat | null {
    if (name === null || depth > 4) return null;
    const el = this.dataStyles.get(name);
    if (!el) return null;

    // LibreOffice writes a signed format as the NEGATIVE sub-format with a map
    // to the positive one; the positive one is the format as authored.
    for (const map of childrenIn(el, NS.STYLE)) {
      if (map.name !== "map") continue;
      const condition = (attr(map, NS.STYLE, "condition") ?? "").replace(/[ \t\n\r\f\v]+/g, "");
      if (condition === "value()>=0" || condition === "value()>0") {
        const mapped = this.dataFormat(attr(map, NS.STYLE, "apply-style-name") ?? "", depth + 1);
        if (mapped !== null) return mapped;
      }
    }

    const parts = childrenIn(el, NS.NUMBER);
    const has = (local: string): boolean => parts.some((p) => p.name === local);
    const number = parts.find((p) => p.name === "number");
    let decimals: number | null = null;
    if (number) {
      const places = attr(number, NS.NUMBER, "decimal-places") ?? attr(number, NS.NUMBER, "min-decimal-places");
      decimals = places !== undefined ? phpInt(places) : null;
    }

    switch (el.name) {
      case "number-style": {
        if (has("scientific-number") || has("fraction")) return null;
        // A currency written as literal text around a number: "$#,##0.00"
        // converted from xlsx arrives this way, not as a currency-style.
        for (const text of parts.filter((p) => p.name === "text")) {
          const iso = currencyCode(ownText(text).replace(/[-()"\u00A0]/g, ""));
          if (iso !== null) return { displayFormat: "currency", decimals: decimals ?? 0, currency: iso };
        }
        if (!number) return null;
        // A number with no decimal places set is "General".
        return decimals === null ? { displayFormat: "auto" } : { displayFormat: "number", decimals };
      }
      case "percentage-style":
        return { displayFormat: "percentage", decimals: decimals ?? 0 };
      case "currency-style": {
        const out: DataFormat = { displayFormat: "currency", decimals: decimals ?? 0 };
        const symbol = parts.find((p) => p.name === "currency-symbol");
        const iso = symbol ? currencyCode(ownText(symbol)) : null;
        if (iso !== null) out.currency = iso;
        return out;
      }
      case "date-style":
        return ["hours", "minutes", "seconds", "am-pm"].some(has) ? { displayFormat: "datetime" } : { displayFormat: "date" };
      case "time-style":
        return { displayFormat: "datetime" };
      case "text-style":
        return { displayFormat: "text" };
      default: // boolean-style, and anything newer than this reader
        return null;
    }
  }
}

/** The nearest value of one property attribute. */
function first(chain: XmlNode[], properties: string, ns: string, name: string): string | null {
  for (const el of chain) {
    const props = child(el, NS.STYLE, properties);
    const value = attr(props, ns, name);
    if (value !== undefined) return phpTrim(value);
  }
  return null;
}

function color(chain: XmlNode[]): string | null {
  for (const el of chain) {
    const props = child(el, NS.STYLE, "text-properties");
    if (!props) continue;
    // "Automatic" colour: whatever the viewer's window text is. No fixed colour.
    if (attr(props, NS.STYLE, "use-window-font-color") === "true") return null;
    const value = attr(props, NS.FO, "color");
    if (value !== undefined) return hex(value);
  }
  return null;
}

function background(chain: XmlNode[]): string | null {
  const value = first(chain, "table-cell-properties", NS.FO, "background-color");
  return value === null || value === "transparent" ? null : hex(value);
}

function border(chain: XmlNode[], side: string): string | null {
  for (const el of chain) {
    const props = child(el, NS.STYLE, "table-cell-properties");
    if (!props) continue;
    const raw = attr(props, NS.FO, `border-${side}`) ?? attr(props, NS.FO, "border");
    if (raw === undefined) continue;

    const value = asciiLower(phpTrim(raw));
    if (value === "" || /\b(none|hidden)\b/.test(value)) return null;
    const m = /#([0-9a-f]{6})\b/.exec(value);
    return m ? "#" + m[1]!.toUpperCase() : "#000000";
  }
  return null;
}

function textAlign(chain: XmlNode[]): string | null {
  // text-align-source="value-type" means "align by what the value is": the
  // stored fo:text-align is not in effect.
  if (first(chain, "table-cell-properties", NS.STYLE, "text-align-source") === "value-type") return null;

  const align = first(chain, "paragraph-properties", NS.FO, "text-align");
  if (align === null || align === "") return null;
  if (align === "start") return "left";
  if (align === "end") return "right";
  return align;
}

function hex(value: string): string | null {
  const m = /^#([0-9a-fA-F]{6})$/.exec(phpTrim(value));
  return m ? "#" + m[1]!.toUpperCase() : null;
}

function currencyCode(symbol: string): string | null {
  const s = phpTrim(symbol);
  if (s === "") return null;
  if (/^[A-Z]{3}$/.test(s)) return s;
  return SYMBOL_TO_ISO[s] ?? null;
}
