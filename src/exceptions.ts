import type { ValidationError } from "./schema/types";

/** Bounded archive/security refusal; contains no document names or content. */
export class ArchiveRefused extends Error {
  constructor(public readonly reason: string) {
    super(`Source archive refused: ${reason}`);
    this.name = "ArchiveRefused";
    Object.setPrototypeOf(this, ArchiveRefused.prototype);
  }
}

const TEMPLATE_MESSAGES: Record<string, string> = {
  macro_enabled: "Macro-enabled templates (.xltm) are refused: re-save it as an .xltx in Excel.",
  family_mismatch: "This package declares a template of a different kind than its contents.",
  already_template: "This package is already a template.",
  already_document: "This package is already a document.",
  not_convertible: "This package has no single content-type declaration that can be rewritten.",
  odf_mimetype: "An OpenDocument package must store its mimetype entry first and uncompressed.",
  bad_target: 'The conversion target must be "document" or "template".',
};

/**
 * An Office or LibreOffice template package refused, with a constant `reason`.
 * Mirrors PHP `OfficeTemplateRefused`:
 *
 * - `macro_enabled`   an .xltm (or a template carrying a VBA project). Never read
 *                     and never converted.
 * - `family_mismatch` it declares a template of another family, or its parts
 *                     are not the family it declares.
 * - `already_template` / `already_document`  conversion asked for what it is.
 * - `not_convertible` there is no single declaration that can be rewritten.
 * - `odf_mimetype`    an ODF package whose `mimetype` is not first and stored.
 * - `bad_target`      the conversion target is not `document` or `template`.
 *
 * Branch on `reason`; the message is for people.
 */
export class OfficeTemplateRefused extends Error {
  constructor(public readonly reason: string) {
    super(`Template refused (${reason}): ${TEMPLATE_MESSAGES[reason] ?? reason}`);
    this.name = "OfficeTemplateRefused";
    Object.setPrototypeOf(this, OfficeTemplateRefused.prototype);
  }
}

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
      `[holy-sheet] cannot read ${path ?? "this file"}: it is not a zip archive, and xlsx and ods both are`,
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
