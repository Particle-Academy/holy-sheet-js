import { Agent, VERSION, type ReadOptions } from "./agent";
import type { RowsOptions } from "./reader/rows-view";
import type {
  BuilderOptions,
  FormulaProblem,
  HolySheetSchema,
  RepairResult,
  ValidationError,
  WriteResult,
} from "./schema/types";

type Any = any;

/**
 * Holy Sheet — instance entry point mirroring PHP `HolySheet`. Thin wrappers
 * over `Agent`. Provided for parity / DI ergonomics; most callers use `Agent`.
 */
export class HolySheet {
  static readonly VERSION = VERSION;

  validate(schema: Any): ValidationError[] {
    return Agent.validate(schema);
  }
  toBytes(schema: Any): Uint8Array {
    return Agent.toBytes(schema);
  }
  write(schema: Any, path: string): Promise<WriteResult> {
    return Agent.write(schema, path);
  }
  toolDefinition(): Record<string, unknown> {
    return Agent.toolDefinition();
  }
  read(input: Uint8Array | ArrayBuffer, options: ReadOptions = {}): Record<string, unknown> {
    return Agent.read(input, options);
  }
  describe(path: string, options: ReadOptions = {}): Promise<Record<string, unknown>> {
    return Agent.describe(path, options);
  }
  /** {@link Agent.rows}: a read sheet as rows of values. */
  rows(sheet: Record<string, unknown>, options: RowsOptions = {}): unknown[][] {
    return Agent.rows(sheet, options);
  }
  /** {@link Agent.convertTemplate}: .xltx/.ots to .xlsx/.ods and back. */
  convertTemplate(input: Uint8Array | ArrayBuffer, to: "document" | "template" | null = null): Uint8Array {
    return Agent.convertTemplate(input, to);
  }
  validateAndRepair(schema: Any): RepairResult {
    return Agent.validateAndRepair(schema);
  }
  fromArray(
    rows: Any[][],
    headers: string[] | null = null,
    sheetName = "Sheet 1",
    options: BuilderOptions = {},
  ): HolySheetSchema {
    return Agent.fromArray(rows, headers, sheetName, options);
  }
  fromCsv(csv: string, options: BuilderOptions = {}): HolySheetSchema {
    return Agent.fromCsv(csv, options);
  }
  lint(schema: Any): FormulaProblem[] {
    return Agent.lint(schema);
  }
  getVersion(): string {
    return VERSION;
  }

  static version(): string {
    return VERSION;
  }
}
