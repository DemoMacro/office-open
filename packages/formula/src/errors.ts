import type { FormulaDialect } from "./ast";
import type { FormulaErrorValue } from "./value";

export type FormulaErrorCode =
  | "empty-formula"
  | "expected-expression"
  | "expected-end"
  | "invalid-character"
  | "invalid-array"
  | "invalid-error-literal"
  | "invalid-number"
  | "invalid-reference"
  | "invalid-string"
  | "invalid-table-reference"
  | "unsupported-syntax";

export interface FormulaDiagnostic {
  code: FormulaErrorCode;
  message: string;
  start: number;
  end: number;
  dialect: FormulaDialect;
}

export class FormulaParseError extends Error implements FormulaDiagnostic {
  readonly code: FormulaErrorCode;
  readonly start: number;
  readonly end: number;
  readonly dialect: FormulaDialect;

  constructor(diagnostic: FormulaDiagnostic, options?: { cause?: unknown }) {
    super(diagnostic.message, options);
    this.name = "FormulaParseError";
    this.code = diagnostic.code;
    this.start = diagnostic.start;
    this.end = diagnostic.end;
    this.dialect = diagnostic.dialect;
  }
}

export class FormulaValueError extends Error {
  readonly value: FormulaErrorValue;
  readonly reason?: string;

  constructor(value: FormulaErrorValue, reason?: string, options?: { cause?: unknown }) {
    super(reason === undefined ? value : `${value}: ${reason}`, options);
    this.name = "FormulaValueError";
    this.value = value;
    this.reason = reason;
  }
}
