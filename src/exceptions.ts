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
