import { isNumericString, numericStringToNumber } from "../util";
import type { BuilderOptions, HolySheetSchema } from "../schema/types";
import { ArrayBuilder } from "./array-builder";

/**
 * CSV string → Holy Sheet schema. Mirrors PHP `Helpers\CsvBuilder`, but takes
 * CSV *content* (isomorphic). To convert a file in Node, read it yourself and
 * pass the string. First row is treated as headers.
 */
export const CsvBuilder = {
  build(csv: string, options: BuilderOptions = {}): HolySheetSchema {
    const delimiter = options.delimiter ?? ",";
    const enclosure = options.enclosure ?? '"';
    const sheetName = options.sheetName ?? "Sheet 1";

    const rows = parseRows(csv, delimiter, enclosure);

    if (rows.length === 0) {
      return { sheets: [{ name: sheetName, columns: [], rows: [] }] };
    }

    const headers = rows[0]!.map((v) => String(v));
    const dataRows: (string | number)[][] = rows.slice(1).map((row) =>
      row.map((cell) => (isNumericString(cell) ? numericStringToNumber(cell) : cell)),
    );

    return ArrayBuilder.build(dataRows, headers, sheetName, options);
  },
};

/** RFC-4180-ish CSV parser: quoted fields, doubled quotes, embedded newlines. */
function parseRows(csv: string, delimiter: string, quote: string): string[][] {
  const rows: string[][] = [];
  let field = "";
  let row: string[] = [];
  let inQuotes = false;
  let started = false; // did this row have any content/field?

  const pushField = (): void => {
    row.push(field);
    field = "";
    started = true;
  };
  const pushRow = (): void => {
    pushField();
    rows.push(row);
    row = [];
    started = false;
  };

  for (let i = 0; i < csv.length; i++) {
    const ch = csv[i]!;
    if (inQuotes) {
      if (ch === quote) {
        if (csv[i + 1] === quote) {
          field += quote;
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }
    if (ch === quote) {
      inQuotes = true;
      started = true;
      continue;
    }
    if (ch === delimiter) {
      pushField();
      continue;
    }
    if (ch === "\r") {
      if (csv[i + 1] === "\n") i++;
      pushRow();
      continue;
    }
    if (ch === "\n") {
      pushRow();
      continue;
    }
    field += ch;
    started = true;
  }
  // trailing field/row (skip a purely-empty trailing line)
  if (field !== "" || row.length > 0 || started) pushRow();

  // Drop blank lines (a single empty field), matching fgetcsv's blank-line skip.
  return rows.filter((r) => !(r.length === 1 && r[0] === ""));
}
