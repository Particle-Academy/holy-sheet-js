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
import { XlsxReader } from "./reader/xlsx-reader";
import { XlsxWriter } from "./writer/xlsx-writer";
import toolSchema from "./holy-sheet.schema.json";

/** Feature-parity baseline with PHP holy-sheet; bumped independently on npm. */
export const VERSION = "1.0.0";

type Any = any;

function toU8(input: Uint8Array | ArrayBuffer): Uint8Array {
  return input instanceof Uint8Array ? input : new Uint8Array(input);
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

  /** Round-trip xlsx bytes back to a schema. Universal. */
  read(input: Uint8Array | ArrayBuffer): Record<string, unknown> {
    return new XlsxReader().describe(toU8(input));
  },

  /** Round-trip an xlsx file on disk back to a schema (Node only). */
  async describe(path: string): Promise<Record<string, unknown>> {
    const fs = await import("node:fs");
    if (!fs.existsSync(path)) return { error: "not_found", path };
    const bytes = fs.readFileSync(path);
    return new XlsxReader().describe(new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength));
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

  version(): string {
    return VERSION;
  },
};
