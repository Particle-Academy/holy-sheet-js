import type { CellFormatInput, DisplayFormat } from "../schema/types";

export interface CellFormatProps {
  bold?: boolean;
  italic?: boolean;
  textAlign?: string | null;
  displayFormat?: DisplayFormat | null;
  decimals?: number | null;
  color?: string | null;
  backgroundColor?: string | null;
  fontSize?: number | null;
  borderTop?: string | null;
  borderRight?: string | null;
  borderBottom?: string | null;
  borderLeft?: string | null;
  currency?: string | null;
  formatCode?: string | null;
  /** A side's line style when it is not thin; null is thin. */
  borderTopStyle?: string | null;
  borderRightStyle?: string | null;
  borderBottomStyle?: string | null;
  borderLeftStyle?: string | null;
}

/** The line styles a border side can have. Mirrors PHP `CellFormat::BORDER_STYLES`. */
export const BORDER_STYLES = ["thin", "medium", "thick", "dashed", "dotted", "double"] as const;

/**
 * One side from the schema: a bare colour (thin), or `{style?, color?}` (thin
 * and black when omitted). Thin is carried as a null style, so a thin object
 * and the bare colour it means are one format. Mirrors PHP `CellFormat::border`.
 */
export function borderSide(side: unknown): [string | null, string | null] {
  if (typeof side === "string") return [side, null];
  // PHP's `[]` is an empty map, so an empty array is `{}` here too.
  if (side === null || typeof side !== "object" || (Array.isArray(side) && side.length > 0)) return [null, null];
  const o = side as Record<string, unknown>;
  const color = typeof o["color"] === "string" ? o["color"] : "#000000";
  const style = o["style"] ?? "thin";
  return [color, typeof style === "string" && style !== "thin" && (BORDER_STYLES as readonly string[]).includes(style) ? style : null];
}

/**
 * Per-cell format. The writer's StylesRegistry deduplicates equal formats.
 * Mirrors PHP `Workbook\CellFormat` (PHP named args → an options object here).
 */
export class CellFormat {
  readonly bold: boolean;
  readonly italic: boolean;
  readonly textAlign: string | null;
  readonly displayFormat: DisplayFormat | null;
  readonly decimals: number | null;
  readonly color: string | null;
  readonly backgroundColor: string | null;
  readonly fontSize: number | null;
  readonly borderTop: string | null;
  readonly borderRight: string | null;
  readonly borderBottom: string | null;
  readonly borderLeft: string | null;
  readonly currency: string | null;
  /**
   * A number-format code no displayFormat describes, written verbatim. Set by a
   * column's `format` (PHP's `formatCode`, which PHP's reader also fills from a
   * foreign code; this port's reader does not yet).
   */
  readonly formatCode: string | null;
  readonly borderTopStyle: string | null;
  readonly borderRightStyle: string | null;
  readonly borderBottomStyle: string | null;
  readonly borderLeftStyle: string | null;

  constructor(p: CellFormatProps = {}) {
    this.bold = p.bold ?? false;
    this.italic = p.italic ?? false;
    this.textAlign = p.textAlign ?? null;
    this.displayFormat = p.displayFormat ?? null;
    this.decimals = p.decimals ?? null;
    this.color = p.color ?? null;
    this.backgroundColor = p.backgroundColor ?? null;
    this.fontSize = p.fontSize ?? null;
    this.borderTop = p.borderTop ?? null;
    this.borderRight = p.borderRight ?? null;
    this.borderBottom = p.borderBottom ?? null;
    this.borderLeft = p.borderLeft ?? null;
    this.currency = p.currency ?? null;
    this.formatCode = p.formatCode ?? null;
    this.borderTopStyle = p.borderTopStyle ?? null;
    this.borderRightStyle = p.borderRightStyle ?? null;
    this.borderBottomStyle = p.borderBottomStyle ?? null;
    this.borderLeftStyle = p.borderLeftStyle ?? null;
  }

  /** Stable dedup key (JSON of ordered fields — replaces PHP md5(serialize)). */
  key(): string {
    return JSON.stringify([
      this.bold,
      this.italic,
      this.textAlign,
      this.displayFormat,
      this.decimals,
      this.color,
      this.backgroundColor,
      this.fontSize,
      this.borderTop,
      this.borderRight,
      this.borderBottom,
      this.borderLeft,
      this.currency,
      this.formatCode,
      this.borderTopStyle,
      this.borderRightStyle,
      this.borderBottomStyle,
      this.borderLeftStyle,
    ]);
  }

  isEmpty(): boolean {
    return (
      !this.bold &&
      !this.italic &&
      this.textAlign === null &&
      this.displayFormat === null &&
      this.decimals === null &&
      this.color === null &&
      this.backgroundColor === null &&
      this.fontSize === null &&
      this.borderTop === null &&
      this.borderRight === null &&
      this.borderBottom === null &&
      this.borderLeft === null &&
      this.currency === null &&
      this.formatCode === null
    );
  }

  static fromInput(a: CellFormatInput): CellFormat {
    const [top, topStyle] = borderSide(a.borderTop ?? null);
    const [right, rightStyle] = borderSide(a.borderRight ?? null);
    const [bottom, bottomStyle] = borderSide(a.borderBottom ?? null);
    const [left, leftStyle] = borderSide(a.borderLeft ?? null);
    return new CellFormat({
      bold: Boolean(a.bold ?? false),
      italic: Boolean(a.italic ?? false),
      textAlign: a.textAlign ?? null,
      displayFormat: a.displayFormat ?? null,
      decimals: a.decimals != null ? Math.trunc(a.decimals) : null,
      color: a.color ?? null,
      backgroundColor: a.backgroundColor ?? null,
      fontSize: a.fontSize != null ? Math.trunc(a.fontSize) : null,
      borderTop: top,
      borderRight: right,
      borderBottom: bottom,
      borderLeft: left,
      currency: a.currency ?? null,
      borderTopStyle: topStyle,
      borderRightStyle: rightStyle,
      borderBottomStyle: bottomStyle,
      borderLeftStyle: leftStyle,
    });
  }

  /** Merge another format on top of this one (other wins where set). */
  mergeWith(other: CellFormat | null): CellFormat {
    if (other === null) return this;
    return new CellFormat({
      bold: other.bold || this.bold,
      italic: other.italic || this.italic,
      textAlign: other.textAlign ?? this.textAlign,
      displayFormat: other.displayFormat ?? this.displayFormat,
      decimals: other.decimals ?? this.decimals,
      color: other.color ?? this.color,
      backgroundColor: other.backgroundColor ?? this.backgroundColor,
      fontSize: other.fontSize ?? this.fontSize,
      borderTop: other.borderTop ?? this.borderTop,
      borderRight: other.borderRight ?? this.borderRight,
      borderBottom: other.borderBottom ?? this.borderBottom,
      borderLeft: other.borderLeft ?? this.borderLeft,
      currency: other.currency ?? this.currency,
      // A displayFormat set on top replaces a code beneath it (PHP).
      formatCode: other.displayFormat !== null ? other.formatCode : (other.formatCode ?? this.formatCode),
      // A side's style travels with its colour: whichever side wins brings its own.
      borderTopStyle: other.borderTop !== null ? other.borderTopStyle : this.borderTopStyle,
      borderRightStyle: other.borderRight !== null ? other.borderRightStyle : this.borderRightStyle,
      borderBottomStyle: other.borderBottom !== null ? other.borderBottomStyle : this.borderBottomStyle,
      borderLeftStyle: other.borderLeft !== null ? other.borderLeftStyle : this.borderLeftStyle,
    });
  }
}
