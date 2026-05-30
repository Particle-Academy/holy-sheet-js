import { Agent, VERSION } from "./agent";
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
  read(input: Uint8Array | ArrayBuffer): Record<string, unknown> {
    return Agent.read(input);
  }
  describe(path: string): Promise<Record<string, unknown>> {
    return Agent.describe(path);
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
