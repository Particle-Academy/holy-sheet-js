import { ArrayBuilder } from "./helpers/array-builder";
import { CsvBuilder } from "./helpers/csv-builder";
import { FormulaLinter } from "./schema/formula-linter";
import { Normalizer } from "./schema/normalizer";
import { Validator } from "./schema/validator";
import type {
  BuilderOptions,
  FormulaProblem,
  HolySheetSchema,
  RepairResult,
  ValidationError,
  WriteResult,
} from "./schema/types";
import { FormatSniffer } from "./reader/format-sniffer";
import { OdsReader } from "./reader/ods-reader";
import { WorkbookSchema } from "./reader/workbook-schema";
import { XlsxReader } from "./reader/xlsx-reader";
import { XlsxWriter } from "./writer/xlsx-writer";
import { SheetDiff } from "./ops/sheet-diff";
import { SheetOpSchema } from "./ops/sheet-op-schema";
import { SheetReducer } from "./ops/sheet-reducer";
import type { SheetOp } from "./ops/types";
import toolSchema from "./holy-sheet.schema.json";

/** This package's own version, pinned to package.json by `version.test.ts`. */
export const VERSION = "2.4.0";

type Any = any;

function toU8(input: Uint8Array | ArrayBuffer): Uint8Array {
  return input instanceof Uint8Array ? input : new Uint8Array(input);
}

function readSpreadsheet(bytes: Uint8Array, path: string | null): Record<string, unknown> {
  const { format, files } = FormatSniffer.sniff(bytes, path);
  const workbook = format === "ods" ? new OdsReader().readFiles(files) : new XlsxReader().readFiles(files);
  return WorkbookSchema.fromWorkbook(workbook);
}

/**
 * Agent — the structured-tool surface for Holy Sheet. Mirrors PHP `Agent`.
 * Universal methods are synchronous; file-touching methods (`write`,
 * `describe`-from-path) are async and Node-only (browsers have no sync FS).
 */
export const Agent = {
  /** Validate a schema without writing. Empty array = valid. */
  validate(schema: Any): ValidationError[] {
    return new Validator().validate(schema);
  },

  /** xlsx bytes without touching disk. Universal. Throws SchemaException if invalid. */
  toBytes(schema: Any): Uint8Array {
    new Validator().assert(schema);
    const workbook = new Normalizer().normalize(schema);
    return new XlsxWriter().toBytes(workbook);
  },

  /** Write a workbook to disk (Node only). Throws SchemaException if invalid. */
  async write(schema: Any, path: string): Promise<WriteResult> {
    new Validator().assert(schema);
    const workbook = new Normalizer().normalize(schema);
    const bytes = new XlsxWriter().toBytes(workbook);
    const fs = await import("node:fs");
    fs.writeFileSync(path, bytes);
    return { path, bytes: bytes.length, sheets: workbook.sheets.length };
  },

  /** JSON Schema for LLM tool-use. */
  toolDefinition(): Record<string, unknown> {
    return toolSchema as Record<string, unknown>;
  },

  /**
   * Round-trip xlsx or ods bytes back to a schema. Universal. The format is
   * told apart by content, never by name, and both describe to the same
   * schema. Throws `UnsupportedFormatException` for anything else.
   */
  read(input: Uint8Array | ArrayBuffer): Record<string, unknown> {
    return readSpreadsheet(toU8(input), null);
  },

  /** Round-trip an xlsx or ods file on disk back to a schema (Node only). */
  async describe(path: string): Promise<Record<string, unknown>> {
    const fs = await import("node:fs");
    if (!fs.existsSync(path)) return { error: "not_found", path };
    const bytes = fs.readFileSync(path);
    return readSpreadsheet(new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength), path);
  },

  /** Validate + conservative repairs in one call. */
  validateAndRepair(schema: Any): RepairResult {
    return new Validator().validateAndRepair(schema);
  },

  /** Build a schema from a flat array of rows (+ optional headers). */
  fromArray(
    rows: Any[][],
    headers: string[] | null = null,
    sheetName = "Sheet 1",
    options: BuilderOptions = {},
  ): HolySheetSchema {
    return ArrayBuilder.build(rows, headers, sheetName, options);
  },

  /** Build a schema from CSV content (read files yourself in Node). */
  fromCsv(csv: string, options: BuilderOptions = {}): HolySheetSchema {
    return CsvBuilder.build(csv, options);
  },

  /** Evaluate every formula and report Excel-style errors. Empty array = clean. */
  lint(schema: Any): FormulaProblem[] {
    return new FormulaLinter().lint(schema);
  },

  /**
   * The ops that turn schema `a` into schema `b`. Mirrors PHP `Agent::diff`, and
   * gives the same ops in the same order for the same inputs.
   *
   * - `reduce(a, diff(a, b))` equals `b` (key order aside).
   * - Schemas that write the same workbook diff to `[]`, so
   *   `diff(s, read(toBytes(s)))` is `[]`: a save without a change records
   *   nothing.
   * - One changed cell is one `set_cell`; an inserted row is one
   *   `insert_rows` plus its cells.
   *
   * Store `diff(newer, older)` to keep a version as the ops that restore it.
   * Both schemas must be valid: the "same workbook" check writes them.
   */
  diff(a: Any, b: Any): SheetOp[] {
    return SheetDiff.diff(a, b);
  },

  /**
   * Apply one op, or a list of them, to a schema; returns a new schema and never
   * modifies the input. An op naming a sheet that is not there is skipped.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  reduce(schema: Any, opOrOps: SheetOp | readonly SheetOp[]): Record<string, any> {
    return SheetReducer.applyAll(schema, isOpList(opOrOps) ? Object.values(opOrOps) : [opOrOps]);
  },

  /**
   * JSON Schema for one op. `set_cell`, `set_range` and `set_workbook` are
   * fancy-sheets' `SheetOp` shapes.
   */
  opSchema(): Record<string, unknown> {
    return SheetOpSchema.jsonSchema();
  },

  /**
   * Whether two schemas write the same workbook: a columns/rows sheet and the
   * cell map it becomes are equivalent, and so are a schema without
   * `meta.created` and its written copy.
   */
  equivalent(a: Any, b: Any): boolean {
    return SheetDiff.equivalent(a, b);
  },

  version(): string {
    return VERSION;
  },
};

/**
 * PHP's `$opOrOps === [] || array_is_list($opOrOps)`: an array, or an object
 * keyed exactly "0".."n-1" (which is how PHP sees `{}` too).
 */
function isOpList(opOrOps: unknown): boolean {
  if (Array.isArray(opOrOps)) return true;
  if (typeof opOrOps !== "object" || opOrOps === null) return false;
  return Object.keys(opOrOps).every((key, i) => key === String(i));
}
