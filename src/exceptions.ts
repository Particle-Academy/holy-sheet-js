import type { ValidationError } from "./schema/types";

/** Thrown when a Holy Sheet schema fails validation. Mirrors PHP `SchemaException`. */
export class SchemaException extends Error {
  readonly errors: ValidationError[];

  constructor(errors: ValidationError[], message?: string) {
    super(message ?? summarize(errors));
    this.name = "SchemaException";
    this.errors = errors;
    Object.setPrototypeOf(this, SchemaException.prototype);
  }

  static fromErrors(errors: ValidationError[]): SchemaException {
    return new SchemaException(errors);
  }

  getErrors(): ValidationError[] {
    return this.errors;
  }
}

function summarize(errors: ValidationError[]): string {
  if (errors.length === 1) {
    const e = errors[0]!;
    return `[holy-sheet] schema invalid at ${e.path}: expected ${e.expected}, got ${e.got}`;
  }
  const first = errors[0]!;
  const rest = errors.length - 1;
  return `[holy-sheet] schema invalid at ${first.path}: expected ${first.expected}, got ${first.got} (+${rest} more)`;
}

/**
 * Thrown by `Agent.read()` / `Agent.describe()` for a file that is neither an
 * xlsx workbook nor an OpenDocument spreadsheet. Mirrors PHP
 * `UnsupportedFormatException`.
 *
 * It extends Error, which is what an unreadable file threw before it existed,
 * so an existing catch still catches it. `mimetype` carries what the package
 * declared about itself when it declared anything (an OpenDocument TEXT file,
 * say), so a caller can say "that is a document, not a spreadsheet".
 */
export class UnsupportedFormatException extends Error {
  readonly path: string | null;
  readonly mimetype: string | null;

  constructor(message: string, path: string | null = null, mimetype: string | null = null) {
    super(message);
    this.name = "UnsupportedFormatException";
    this.path = path;
    this.mimetype = mimetype;
    Object.setPrototypeOf(this, UnsupportedFormatException.prototype);
  }

  static notAZip(path: string | null = null): UnsupportedFormatException {
    return new UnsupportedFormatException(
      `[holy-sheet] cannot read ${path ?? "this file"}: it is not a zip package, and xlsx and ods both are`,
      path,
    );
  }

  static unknownPackage(path: string | null, mimetype: string | null): UnsupportedFormatException {
    const what =
      mimetype !== null && mimetype !== ""
        ? `its mimetype is ${mimetype}`
        : "it has neither xl/workbook.xml nor an OpenDocument mimetype";
    return new UnsupportedFormatException(
      `[holy-sheet] cannot read ${path ?? "this file"}: ${what}. Supported: xlsx, ods`,
      path,
      mimetype !== "" ? mimetype : null,
    );
  }
}
