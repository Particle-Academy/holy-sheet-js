import { unzipSync } from "../zip";
import { Cell } from "../workbook/cell";
import { CellAddress } from "../workbook/cell-address";
import { CellComment } from "../workbook/cell-comment";
import { MergedRegion } from "../workbook/merged-region";
import { Sheet } from "../workbook/sheet";
import { Workbook } from "../workbook/workbook";
import type { CellPrimitive } from "../schema/types";
import { NS, asciiLower, attr, child, phpInt, phpIsNumeric, phpTrim } from "./ods/ns";
import { odsFormulaToA1 } from "./ods/ods-formula";
import { OdsStyles } from "./ods/ods-styles";
import { paragraphs } from "./ods/ods-text";
import { WorkbookSchema } from "./workbook-schema";
import { parseXml, type XmlNode } from "./xml";

/** Sheet bounds. A repeat past them is padding, and is not followed. */
const MAX_ROWS = 1048576;
const MAX_COLUMNS = 16384;

/** Days from the spreadsheet epoch (1899-12-30) to the Unix epoch. */
const UNIX_EPOCH_SERIAL = 25569;

const decoder = new TextDecoder();

type ColumnStyle = [first: number, last: number, style: string];

interface ReadCell {
  value: CellPrimitive;
  formula: string | null;
  cached: CellPrimitive;
  type: string | null;
  currency: string | null;
  dateHasTime: boolean;
  comment: CellComment | null;
  /** A date or time value as Unix seconds (UTC), until the format decides date or datetime. */
  seconds: number | null;
}

interface SheetState {
  row: number;
  cells: Map<string, Cell>;
  merges: MergedRegion[];
}

/**
 * OpenDocument Spreadsheet (.ods) reader — bytes → Holy Sheet schema, the same
 * schema the xlsx reader returns. Mirrors PHP `Reader\OdsReader`, where what is
 * and is not mapped is documented.
 */
export class OdsReader {
  describe(bytes: Uint8Array): Record<string, unknown> {
    return WorkbookSchema.fromWorkbook(this.readWorkbook(bytes));
  }

  readWorkbook(bytes: Uint8Array): Workbook {
    return this.readFiles(unzipSync(bytes));
  }

  /** Read an already-unzipped package (name -> bytes). */
  readFiles(files: Record<string, Uint8Array>): Workbook {
    const getText = (name: string): string | null => (files[name] ? decoder.decode(files[name]) : null);

    const content = getText("content.xml");
    if (content === null) throw new Error("[holy-sheet] missing content.xml");
    const contentDoc = parseXml(content, { namespaces: true });
    if (!contentDoc) throw new Error("[holy-sheet] failed to parse content.xml");
    const stylesText = getText("styles.xml");
    const stylesDoc = stylesText === null ? null : parseXml(stylesText, { namespaces: true });

    const styles = new OdsStyles(contentDoc, stylesDoc);

    const sheets: Sheet[] = [];
    const spreadsheet = child(child(contentDoc, NS.OFFICE, "body"), NS.OFFICE, "spreadsheet");
    for (const table of spreadsheet?.children ?? []) {
      if (table.ns === NS.TABLE && table.name === "table") sheets.push(this.readSheet(table, styles));
    }

    const meta = getText("meta.xml");
    return new Workbook(sheets, meta === null ? {} : readMeta(meta));
  }

  private readSheet(table: XmlNode, styles: OdsStyles): Sheet {
    const columns: ColumnStyle[] = [];
    collectColumns(table, columns, { next: 0 });

    const state: SheetState = { row: 0, cells: new Map(), merges: [] };
    this.walkRows(table, styles, columns, state);

    return new Sheet(attr(table, NS.TABLE, "name") ?? "", state.cells, state.merges);
  }

  private walkRows(container: XmlNode, styles: OdsStyles, columns: ColumnStyle[], state: SheetState): void {
    for (const el of container.children) {
      if (state.row >= MAX_ROWS) return;
      if (el.ns !== NS.TABLE) continue;

      if (el.name === "table-row") {
        this.readRow(el, styles, columns, state);
      } else if (el.name === "table-header-rows" || el.name === "table-rows" || el.name === "table-row-group") {
        this.walkRows(el, styles, columns, state);
      }
    }
  }

  private readRow(row: XmlNode, styles: OdsStyles, columns: ColumnStyle[], state: SheetState): void {
    const repeat = Math.max(1, phpInt(attr(row, NS.TABLE, "number-rows-repeated") ?? "1"));
    const rowStyle = attr(row, NS.TABLE, "default-cell-style-name") ?? "";

    // One pass over the cells, reading each one that holds anything ONCE,
    // however many columns and rows it repeats across.
    const entries: Array<[number, ReadCell | null, [number, number], string]> = [];
    let column = 0;
    for (const cell of row.children) {
      if (cell.ns !== NS.TABLE || (cell.name !== "table-cell" && cell.name !== "covered-table-cell")) continue;

      const span: [number, number] =
        cell.name === "table-cell"
          ? [
              Math.max(1, phpInt(attr(cell, NS.TABLE, "number-columns-spanned") ?? "1")),
              Math.max(1, phpInt(attr(cell, NS.TABLE, "number-rows-spanned") ?? "1")),
            ]
          : [1, 1];
      const read = hasContent(cell) ? readCell(cell) : null;

      const count = Math.max(1, phpInt(attr(cell, NS.TABLE, "number-columns-repeated") ?? "1"));
      if (read !== null || span[0] !== 1 || span[1] !== 1) {
        const style = attr(cell, NS.TABLE, "style-name") ?? "";
        for (let i = 0; i < count && column + i < MAX_COLUMNS; i++) {
          entries.push([column + i, read, span, style]);
        }
      }

      column += count;
      if (column >= MAX_COLUMNS) break;
    }

    if (entries.length === 0) {
      state.row += repeat;
      return;
    }

    for (let r = 0; r < repeat && state.row < MAX_ROWS; r++) {
      const rowNumber = state.row + 1;
      for (const [col, read, span, style] of entries) {
        const address = CellAddress.letter(col) + rowNumber;

        if (span[0] !== 1 || span[1] !== 1) {
          state.merges.push(
            new MergedRegion(
              address,
              CellAddress.letter(Math.min(col + span[0], MAX_COLUMNS) - 1) + Math.min(rowNumber + span[1] - 1, MAX_ROWS),
            ),
          );
        }
        if (read === null) continue;

        const styleName = style !== "" ? style : rowStyle !== "" ? rowStyle : columnStyle(columns, col);
        const format = styles.cellFormat(styleName, read.type, read.currency, read.dateHasTime);
        let value = read.value;
        if (read.seconds !== null) {
          // A date or time value always has a date or datetime format by now.
          value = isoFromSeconds(read.seconds, format?.displayFormat !== "date");
        }
        state.cells.set(address, new Cell(address, value, read.formula, format, read.comment, read.cached));
      }
      state.row++;
    }
  }
}

/** Column default cell styles as [first index, last index, style name]. */
function collectColumns(container: XmlNode, columns: ColumnStyle[], cursor: { next: number }): void {
  for (const el of container.children) {
    if (el.ns === NS.TABLE && el.name === "table-column") {
      const repeat = Math.max(1, phpInt(attr(el, NS.TABLE, "number-columns-repeated") ?? "1"));
      const style = attr(el, NS.TABLE, "default-cell-style-name") ?? "";
      if (style !== "") columns.push([cursor.next, cursor.next + repeat - 1, style]);
      cursor.next += repeat;
    } else if (
      el.ns === NS.TABLE &&
      (el.name === "table-columns" || el.name === "table-header-columns" || el.name === "table-column-group")
    ) {
      collectColumns(el, columns, cursor);
    }
    if (cursor.next >= MAX_COLUMNS) return;
  }
}

function columnStyle(columns: ColumnStyle[], column: number): string {
  for (const [first, last, style] of columns) {
    if (column >= first && column <= last) return style;
  }
  return "Default";
}

/**
 * A cell holds something when it has a value type, a formula, a comment, or
 * text. A cell with only a style is formatting over nothing, as in xlsx.
 */
function hasContent(cell: XmlNode): boolean {
  if (attr(cell, NS.OFFICE, "value-type") !== undefined) return true;
  if (attr(cell, NS.TABLE, "formula") !== undefined) return true;
  if (child(cell, NS.OFFICE, "annotation") !== undefined) return true;
  return (paragraphs(cell) ?? "") !== "";
}

/** Everything about a cell that does not depend on where it is. */
function readCell(cell: XmlNode): ReadCell {
  const type = attr(cell, NS.OFFICE, "value-type") ?? null;
  const text = paragraphs(cell);

  let value: CellPrimitive = null;
  let seconds: number | null = null;
  let dateHasTime = false;
  switch (type) {
    case "float":
    case "percentage":
    case "currency":
      value = toNumber(attr(cell, NS.OFFICE, "value") ?? "");
      break;
    case "boolean": {
      const b = asciiLower(phpTrim(attr(cell, NS.OFFICE, "boolean-value") ?? ""));
      value = b === "true" || b === "1";
      break;
    }
    case "date": {
      const raw = phpTrim(attr(cell, NS.OFFICE, "date-value") ?? "");
      seconds = dateSeconds(raw);
      dateHasTime = raw.includes("T");
      break;
    }
    case "time": {
      const duration = durationSeconds(attr(cell, NS.OFFICE, "time-value") ?? "");
      seconds = duration === null ? null : duration - UNIX_EPOCH_SERIAL * 86400;
      break;
    }
    case "string": {
      const error = attr(cell, NS.CALCEXT, "value-type") === "error";
      const stored = attr(cell, NS.OFFICE, "string-value");
      value = stored !== undefined && !error ? stored : (text ?? "");
      break;
    }
    default:
      value = text !== null && text !== "" ? text : null;
  }

  let formula: string | null = null;
  let cached: CellPrimitive = null;
  const formulaAttr = attr(cell, NS.TABLE, "formula");
  if (formulaAttr !== undefined) {
    const translated = odsFormulaToA1(formulaAttr);
    const literalBoolean = type === "boolean" && /^[ \t\n\r\f\v]*(TRUE|FALSE)\(\)[ \t\n\r\f\v]*$/i.test(translated);
    if (!literalBoolean) {
      formula = translated;
      cached = seconds !== null ? serial(seconds) : value;
      value = null;
      seconds = null;
    }
  }

  let comment: CellComment | null = null;
  const annotation = child(cell, NS.OFFICE, "annotation");
  if (annotation) {
    const creator = child(annotation, NS.DC, "creator");
    const author = creator ? phpTrim(creator.text) : "";
    comment = new CellComment(paragraphs(annotation) ?? "", author === "" ? null : author);
  }

  const currency = attr(cell, NS.OFFICE, "currency");
  return {
    value,
    formula,
    cached,
    type,
    currency: currency !== undefined ? phpTrim(currency) : null,
    dateHasTime,
    comment,
    seconds,
  };
}

/** `office:value` as the xlsx reader coerces `<v>`. */
function toNumber(raw: string): number | null {
  const s = phpTrim(raw);
  if (s === "") return null;
  if (/^-?\d+$/.test(s)) return parseInt(s, 10);
  return phpIsNumeric(s) ? Number(s) : null;
}

/** `office:date-value` -> whole seconds since the Unix epoch, in UTC. */
function dateSeconds(raw: string): number | null {
  const m = /^(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?)?(Z|[+-]\d{2}:?\d{2})?$/.exec(raw);
  if (!m) return null;

  const days = daysFromCivil(parseInt(m[1]!, 10), parseInt(m[2]!, 10), parseInt(m[3]!, 10));
  let seconds = days * 86400 + phpInt(m[4] ?? "0") * 3600 + phpInt(m[5] ?? "0") * 60 + phpInt(m[6] ?? "0");
  const fraction = m[7] !== undefined && m[7] !== "" ? Number("0" + m[7]) : 0;

  const zone = m[8] ?? "";
  if (zone !== "" && zone !== "Z") {
    const digits = zone.slice(1).replace(":", "");
    const offset = phpInt(digits.slice(0, 2)) * 3600 + phpInt(digits.slice(2, 4)) * 60;
    seconds -= zone[0] === "-" ? -offset : offset;
  }

  return seconds + Math.round(fraction);
}

/** `office:time-value` (`PT10H30M00S`, `P1DT2H`, `-PT1H`) -> whole seconds. */
function durationSeconds(raw: string): number | null {
  const m = /^(-)?P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(phpTrim(raw));
  if (!m) return null;

  const total =
    phpInt(m[2] ?? "0") * 86400 + phpInt(m[3] ?? "0") * 3600 + phpInt(m[4] ?? "0") * 60 + (m[5] ? Number(m[5]) : 0);
  const seconds = Math.round(total);
  return m[1] === "-" ? -seconds : seconds;
}

/** Days since 1970-01-01 in the proleptic Gregorian calendar (Hinnant's algorithm). */
function daysFromCivil(year: number, month: number, day: number): number {
  year -= month <= 2 ? 1 : 0;
  const era = Math.trunc((year >= 0 ? year : year - 399) / 400);
  const yoe = year - era * 400;
  const doy = Math.trunc((153 * (month + (month > 2 ? -3 : 9)) + 2) / 5) + day - 1;
  const doe = yoe * 365 + Math.trunc(yoe / 4) - Math.trunc(yoe / 100) + doy;
  return era * 146097 + doe - 719468;
}

/** Unix seconds -> spreadsheet serial. */
function serial(seconds: number): number {
  return (seconds + UNIX_EPOCH_SERIAL * 86400) / 86400;
}

/** PHP `gmdate('Y-m-d')` / `gmdate('Y-m-d\TH:i:s\Z')`. */
function isoFromSeconds(seconds: number, withTime: boolean): string {
  const d = new Date(seconds * 1000);
  const pad = (n: number, width = 2): string => String(n).padStart(width, "0");
  const date = `${pad(d.getUTCFullYear(), 4)}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  if (!withTime) return date;
  return `${date}T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}Z`;
}

function readMeta(xml: string): Record<string, string> {
  const doc = parseXml(xml, { namespaces: true });
  const meta = child(doc, NS.OFFICE, "meta");
  if (!meta) return {};

  const out: Record<string, string> = {};
  const creator = child(meta, NS.META, "initial-creator") ?? child(meta, NS.DC, "creator");
  if (creator && phpTrim(creator.text) !== "") out["creator"] = phpTrim(creator.text);
  const created = child(meta, NS.META, "creation-date");
  if (created && phpTrim(created.text) !== "") {
    const value = phpTrim(created.text);
    // ODF dates carry no zone unless one is written; the schema's are UTC.
    out["created"] = /(Z|[+-]\d{2}:?\d{2})$/.test(value) ? value : value + "Z";
  }
  return out;
}
