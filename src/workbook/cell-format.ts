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
      this.currency === null
    );
  }

  static fromInput(a: CellFormatInput): CellFormat {
    return new CellFormat({
      bold: Boolean(a.bold ?? false),
      italic: Boolean(a.italic ?? false),
      textAlign: a.textAlign ?? null,
      displayFormat: a.displayFormat ?? null,
      decimals: a.decimals != null ? Math.trunc(a.decimals) : null,
      color: a.color ?? null,
      backgroundColor: a.backgroundColor ?? null,
      fontSize: a.fontSize != null ? Math.trunc(a.fontSize) : null,
      borderTop: a.borderTop ?? null,
      borderRight: a.borderRight ?? null,
      borderBottom: a.borderBottom ?? null,
      borderLeft: a.borderLeft ?? null,
      currency: a.currency ?? null,
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
    });
  }
}
