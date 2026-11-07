import { CellAddress } from "../workbook/cell-address";
import { ChartRange } from "../workbook/chart-range";
import { DateConverter } from "../writer/format/date-converter";
import { formatFloat, isPlainObject, typeOf } from "../util";
import { ColumnWidths } from "./column-widths";

/**
 * Data validation: the schema shape, its problems, its normal form and its XML.
 * A line-by-line port of PHP `Schema\DataValidations`; the error paths,
 * wording and check order are PHP's, and so are the bytes.
 *
 * Lengths are counted in characters (code points), as PHP's `mb_strlen`
 * does, never in UTF-16 units.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/** [subpath, expected, got, value, hint] */
export type DvProblem = [string, string, string, unknown, string];

export interface NormalizedValidation {
  sqref: string;
  type: string;
  operator: string | null;
  formula1: string;
  formula2: string | null;
  allowBlank: boolean;
  hideDropdown: boolean;
  prompt: { title: string; text: string } | null;
  error: { style: string; title: string; text: string };
}

export const DV_TYPES = ["list", "whole", "decimal", "date", "textLength"];
export const DV_OPERATORS = ["between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"];
export const DV_STYLES = ["stop", "warning", "information"];

const MAX_LIST = 255;
const MAX_WHOLE = 9007199254740991;
const MAX_ROW = 1048576;

const has = (o: Any, k: string): boolean => Object.prototype.hasOwnProperty.call(o, k);
const chars = (s: string): number => Array.from(s).length;

/** PHP's array: an object, or an EMPTY array (PHP cannot tell [] from {}). */
const isObj = (v: unknown): v is Record<string, Any> => isPlainObject(v) || (Array.isArray(v) && v.length === 0);

/** PHP's list: an array, or an EMPTY object. */
const isList = (v: unknown): v is Any[] => Array.isArray(v) || (isPlainObject(v) && Object.keys(v).length === 0);

const asList = (v: Any): Any[] => (Array.isArray(v) ? v : []);

export const DataValidations = {
  problems(dv: unknown, ownSheet: string, sheetNames: string[] | null, column: boolean, dataRows: number): DvProblem[] {
    if (!isObj(dv)) {
      return [["", "object", typeOf(dv), dv, column
        ? 'A column validation is an object: {"type": "list", "values": ["Open", "Closed"]}.'
        : 'A validation is an object: {"range": "B2:B100", "type": "list", "values": ["Open", "Closed"]}.']];
    }
    const d = dv as Record<string, Any>;
    const out: DvProblem[] = [];
    const type = d.type ?? null;
    const known = typeof type === "string" && DV_TYPES.includes(type);
    if (!known) {
      out.push([".type", "one of: " + DV_TYPES.join(", "), typeOf(type), type,
        "list is a dropdown; whole, decimal and date bound a number or a date; textLength bounds how long the text is."]);
    }

    if (column) {
      if (has(d, "range")) {
        out.push([".range", "absent", typeOf(d.range), d.range,
          "A column validation covers that column's data rows. To choose the cells yourself, put it in the sheet's dataValidations with a range."]);
      } else if (dataRows === 0) {
        out.push(["", "a sheet with data rows", "no rows", null,
          "A column validation covers the column's data rows, and this sheet has none. For a template, put it in the sheet's dataValidations with a range, like \"B2:B500\"."]);
      }
    } else if (DataValidations.sqref(d.range ?? null) === null) {
      out.push([".range", 'a cell or range on this sheet, like "B2:B100"', typeOf(d.range ?? null), d.range ?? null,
        'The cells the rule applies to, on this sheet. A dropdown for a template column usually runs well past the sample rows: "B2:B500".']);
    }

    if (type === "list") out.push(...listProblems(d, ownSheet, sheetNames));
    else if (known) out.push(...boundProblems(d, type));

    if (has(d, "allowBlank") && typeof d.allowBlank !== "boolean") {
      out.push([".allowBlank", "boolean", typeOf(d.allowBlank), d.allowBlank,
        "true (the default) lets a cell be left empty; false makes an empty cell break the rule."]);
    }

    if (has(d, "prompt")) {
      const p = d.prompt;
      if (!isObj(p)) {
        out.push([".prompt", "object {title?, text?}", typeOf(p), p,
          'Shown when the cell is selected: {"title": "Status", "text": "Pick one."}.']);
      } else {
        out.push(...textProblems(p, ".prompt", 255, "Excel limits an input prompt to 255 characters."));
      }
    }

    if (has(d, "error")) {
      const e = d.error;
      if (!isObj(e)) {
        out.push([".error", "object {style?, title?, text?}", typeOf(e), e,
          'Shown when an entry breaks the rule: {"style": "stop", "title": "Invalid status", "text": "Pick a status from the list."}.']);
      } else {
        if (has(e, "style") && !(typeof e.style === "string" && DV_STYLES.includes(e.style))) {
          out.push([".error.style", "one of: " + DV_STYLES.join(", "), typeOf(e.style), e.style,
            "stop refuses the entry (the default); warning lets the person keep it after a warning; information only tells them."]);
        }
        out.push(...textProblems(e, ".error", 225, "Excel limits an error message to 225 characters."));
      }
    }

    return out;
  },

  normalize(dv: unknown, sqref: string, ownSheet: string): NormalizedValidation | null {
    // The range was checked by the caller and is carried as sqref; the rest is
    // checked as a column rule with rows, which is every other check.
    let d = dv as Record<string, Any>;
    if (isPlainObject(dv)) {
      d = { ...(dv as Record<string, Any>) };
      delete d.range;
    }
    if (DataValidations.problems(d, ownSheet, null, true, 1).length > 0) return null;
    const type: string = d.type;
    let operator: string | null = null;
    let formula1: string;
    let formula2: string | null = null;
    if (type === "list") {
      if (has(d, "values")) {
        formula1 = '"' + (d.values as Any[]).map((v) => choice(v)!).join(",").replace(/"/g, '""') + '"';
      } else {
        formula1 = sourceFormula(gridRange(d.source, ownSheet)!, ownSheet);
      }
    } else {
      const op: string = d.operator ?? "between";
      operator = op === "between" ? null : op;
      if (op === "between" || op === "notBetween") {
        formula1 = operand(d.min, type)!;
        formula2 = operand(d.max, type);
      } else {
        formula1 = operand(d.value, type)!;
      }
    }

    const prompt = isObj(d.prompt ?? null)
      ? { title: d.prompt.title != null ? String(d.prompt.title) : "", text: d.prompt.text != null ? String(d.prompt.text) : "" }
      : null;
    const error = isObj(d.error ?? null) ? d.error : {};

    return {
      sqref,
      type,
      operator,
      formula1,
      formula2,
      allowBlank: (d.allowBlank ?? true) === true,
      hideDropdown: type === "list" && (d.showDropdown ?? true) === false,
      prompt,
      error: {
        style: error.style != null ? String(error.style) : "stop",
        title: error.title != null ? String(error.title) : "",
        text: error.text != null ? String(error.text) : "",
      },
    };
  },

  /** Column rules first, in column order, then the sheet's own list. */
  forSheet(sheet: Record<string, Any>, name: string): NormalizedValidation[] {
    const out: NormalizedValidation[] = [];
    const columns: Any[] = Array.isArray(sheet.columns) ? sheet.columns : [];
    const rows = Array.isArray(sheet.rows) ? sheet.rows.length : 0;
    if (rows > 0) {
      columns.forEach((col, i) => {
        if (!isPlainObject(col) || !has(col, "validation")) return;
        const letter = CellAddress.letter(i);
        const sqref = letter + "2" + (rows > 1 ? ":" + letter + (rows + 1) : "");
        const v = col.validation;
        if (isObj(v) && !has(v, "range")) {
          const n = DataValidations.normalize(v, sqref, name);
          if (n !== null) out.push(n);
        }
      });
    }
    if (isList(sheet.dataValidations ?? null)) {
      for (const dv of asList(sheet.dataValidations)) {
        const sqref = isObj(dv) ? DataValidations.sqref((dv as Any).range ?? null) : null;
        if (sqref === null) continue;
        const n = DataValidations.normalize(dv, sqref, name);
        if (n !== null) out.push(n);
      }
    }
    return out;
  },

  /** `<dataValidations>`, or "" for none. Attribute order is PHP's (CT_DataValidation's). */
  xml(rules: NormalizedValidation[], escape: (s: string) => string): string {
    if (rules.length === 0) return "";
    let xml = `<dataValidations count="${rules.length}">`;
    for (const r of rules) {
      xml +=
        `<dataValidation type="${r.type}"` +
        (r.error.style !== "stop" ? ` errorStyle="${r.error.style}"` : "") +
        (r.operator !== null ? ` operator="${r.operator}"` : "") +
        (r.allowBlank ? ' allowBlank="1"' : "") +
        // OOXML's showDropDown="1" HIDES the arrow; the schema's showDropdown says what a person sees.
        (r.hideDropdown ? ' showDropDown="1"' : "") +
        (r.prompt !== null ? ' showInputMessage="1"' : "") +
        ' showErrorMessage="1"' +
        (r.error.title !== "" ? ` errorTitle="${escape(r.error.title)}"` : "") +
        (r.error.text !== "" ? ` error="${escape(r.error.text)}"` : "") +
        (r.prompt !== null && r.prompt.title !== "" ? ` promptTitle="${escape(r.prompt.title)}"` : "") +
        (r.prompt !== null && r.prompt.text !== "" ? ` prompt="${escape(r.prompt.text)}"` : "") +
        ` sqref="${r.sqref}">` +
        `<formula1>${escape(r.formula1)}</formula1>` +
        (r.formula2 !== null ? `<formula2>${escape(r.formula2)}</formula2>` : "") +
        "</dataValidation>";
    }
    return xml + "</dataValidations>";
  },

  /** A sheet-level range as the canonical sqref, or null when it is not one on this sheet inside the grid. */
  sqref(range: unknown): string | null {
    if (typeof range !== "string" || range.includes("!")) return null;
    const r = gridRange(range, "");
    return r === null ? null : r.toSchema("");
  },
};

function listProblems(d: Record<string, Any>, ownSheet: string, sheetNames: string[] | null): DvProblem[] {
  const out: DvProblem[] = [];
  const hasValues = has(d, "values");
  const hasSource = has(d, "source");
  if (hasValues === hasSource) {
    out.push(["", "exactly one of values or source", hasValues ? "both" : "neither", null,
      "A dropdown lists its choices inline, {\"values\": [\"Open\", \"Closed\"]}, or reads them from cells, {\"source\": \"'Lists'!A1:A10\"}."]);
  } else if (hasValues) {
    const values = d.values;
    if (!Array.isArray(values) || values.length === 0) {
      out.push([".values", "a non-empty list of choices", typeOf(values), values, 'The choices, in order: ["Open", "Closed"].']);
    } else {
      let bad = false;
      values.forEach((v, k) => {
        if (choice(v) === null) {
          bad = true;
          out.push([`.values[${k}]`, "a non-empty string with no comma, or a whole number", typeOf(v), v,
            'Each choice is text, like "Open", or a whole number. Excel separates inline choices with commas and cannot escape one, so a choice containing a comma must come from cells: use source.']);
        }
      });
      if (!bad) {
        const length = chars(values.map((v) => choice(v)!).join(","));
        if (length > MAX_LIST) {
          out.push([".values", "at most 255 characters, commas included", `${length} characters`, values,
            "Excel cannot store an inline list longer than 255 characters. Put the choices in cells and point source at them, like \"'Lists'!A1:A40\"."]);
        }
      }
    }
  } else {
    const range = gridRange(d.source, ownSheet);
    if (range === null) {
      out.push([".source", "a cell range like \"'Lists'!A1:A10\"", typeOf(d.source), d.source,
        "The cells holding the choices: \"A2:A20\" on this sheet, or \"'Lists'!A1:A10\" on another."]);
    } else if (sheetNames !== null && !sheetNames.includes(range.sheet)) {
      out.push([".source", "a sheet in this workbook", "unknown sheet", d.source,
        `There is no sheet named "${range.sheet}". Name the sheet exactly as sheets[].name has it.`]);
    }
  }

  if (has(d, "showDropdown") && typeof d.showDropdown !== "boolean") {
    out.push([".showDropdown", "boolean", typeOf(d.showDropdown), d.showDropdown,
      "true (the default) shows the in-cell dropdown arrow; false hides the arrow and still enforces the list."]);
  }

  for (const key of ["operator", "min", "max", "value"]) {
    if (has(d, key)) out.push([`.${key}`, "absent", typeOf(d[key]), d[key], `${key} does not apply to a list: its choices are the rule.`]);
  }
  return out;
}

function boundProblems(d: Record<string, Any>, type: string): DvProblem[] {
  const out: DvProblem[] = [];
  const operator = d.operator ?? "between";
  const known = typeof operator === "string" && DV_OPERATORS.includes(operator);
  if (!known) {
    out.push([".operator", "one of: " + DV_OPERATORS.join(", "), typeOf(operator), operator,
      "Omit for between. between and notBetween take min and max; the others take value."]);
  }

  const [expected, hint] =
    type === "whole" ? ["a whole number", "A whole number, like 10."]
    : type === "decimal" ? ["a number", "A number, like 0.5."]
    : type === "date" ? ["a date as YYYY-MM-DD", 'A calendar date, like "2026-12-31".']
    : ["a whole number of characters, 0 or more", "A length in characters, like 50."];

  if (known) {
    const two = operator === "between" || operator === "notBetween";
    for (const key of two ? ["min", "max"] : ["value"]) {
      const v = d[key] ?? null;
      if (operand(v, type) === null) out.push([`.${key}`, expected, typeOf(v), v, hint]);
    }
    if (two && operand(d.min ?? null, type) !== null && operand(d.max ?? null, type) !== null && compare(d.min, d.max, type) > 0) {
      out.push([".max", "a value no less than min", typeOf(d.max), d.max,
        "min is the lower bound and max the upper; as written, nothing could ever pass this rule."]);
    }
    for (const key of two ? ["value"] : ["min", "max"]) {
      if (has(d, key)) {
        out.push([`.${key}`, "absent", typeOf(d[key]), d[key], two
          ? `${operator} takes min and max, not value.`
          : `${operator} compares against one value: use value, not min and max.`]);
      }
    }
  }

  for (const key of ["values", "source", "showDropdown"]) {
    if (has(d, key)) out.push([`.${key}`, "absent", typeOf(d[key]), d[key], `${key} applies only to a list.`]);
  }
  return out;
}

function textProblems(box: Record<string, Any>, at: string, max: number, hint: string): DvProblem[] {
  const out: DvProblem[] = [];
  if (has(box, "title") && (typeof box.title !== "string" || chars(box.title) > 32)) {
    out.push([`${at}.title`, "a string of at most 32 characters", typeOf(box.title), box.title,
      "Excel limits a prompt or error title to 32 characters."]);
  }
  if (has(box, "text") && (typeof box.text !== "string" || chars(box.text) > max)) {
    out.push([`${at}.text`, `a string of at most ${max} characters`, typeOf(box.text), box.text, hint]);
  }
  return out;
}

function gridRange(ref: unknown, ownSheet: string): ChartRange | null {
  const r = ChartRange.parse(ref, ownSheet);
  if (r === null || r.c2 > ColumnWidths.MAX_INDEX || r.r2 > MAX_ROW) return null;
  return r;
}

/** $A$1:$A$9 on the rule's own sheet, 'Lists'!$A$1:$A$9 on another. */
function sourceFormula(r: ChartRange, ownSheet: string): string {
  const from = "$" + CellAddress.letter(r.c1) + "$" + r.r1;
  const to = "$" + CellAddress.letter(r.c2) + "$" + r.r2;
  const ref = from === to ? from : from + ":" + to;
  return r.sheet === ownSheet ? ref : "'" + r.sheet.replace(/'/g, "''") + "'!" + ref;
}

/** An inline choice as text, or null when it cannot be one. */
function choice(v: unknown): string | null {
  if (typeof v === "string") return v !== "" && !v.includes(",") ? v : null;
  if (typeof v === "number" && Number.isInteger(v) && Math.abs(v) <= MAX_WHOLE) return String(v);
  return null;
}

/** A number with no fraction inside the exact-integer range. */
function whole(v: unknown): number | null {
  return typeof v === "number" && Number.isInteger(v) && Math.abs(v) <= MAX_WHOLE ? v : null;
}

/** An operand as the text of its formula, or null when it is not one for this type. */
function operand(v: unknown, type: string): string | null {
  switch (type) {
    case "whole":
    case "textLength": {
      const w = whole(v);
      if (w === null || (type === "textLength" && w < 0)) return null;
      return String(w === 0 ? 0 : w);
    }
    case "decimal":
      if (typeof v !== "number" || !Number.isFinite(v)) return null;
      return Number.isSafeInteger(v) ? String(v === 0 ? 0 : v) : formatFloat(v);
    case "date": {
      if (typeof v !== "string") return null;
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
      if (!m) return null;
      const y = +m[1]!, mo = +m[2]!, day = +m[3]!;
      const t = new Date(Date.UTC(2000, mo - 1, day));
      t.setUTCFullYear(y);
      if (y < 1 || t.getUTCFullYear() !== y || t.getUTCMonth() !== mo - 1 || t.getUTCDate() !== day) return null;
      return String(DateConverter.toSerial(v));
    }
  }
  return null;
}

function compare(a: Any, b: Any, type: string): number {
  if (type === "date") return a < b ? -1 : a > b ? 1 : 0;
  return a < b ? -1 : a > b ? 1 : 0;
}
