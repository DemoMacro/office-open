export const formulaErrorValues = [
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#GETTING_DATA",
  "#SPILL!",
  "#CALC!",
  "#FIELD!",
  "#CONNECT!",
  "#BLOCKED!",
  "#BUSY!",
  "#UNKNOWN!",
] as const;

export type FormulaErrorValue = (typeof formulaErrorValues)[number];

export class FormulaError {
  readonly kind = "formula-error" as const;

  constructor(readonly value: FormulaErrorValue) {}

  toString(): string {
    return this.value;
  }
}

export type FormulaScalar = null | boolean | number | string | FormulaError;
export type FormulaMatrix = FormulaScalar[][];
export type FormulaValue = FormulaScalar | FormulaMatrix;

export const maximumFormulaCells = 2_000_000;

export function isMatrix(value: FormulaValue): value is FormulaMatrix {
  return Array.isArray(value);
}

export function formulaError(value: FormulaErrorValue): FormulaError {
  return new FormulaError(value);
}

export function isFormulaErrorValue(value: unknown): value is FormulaError {
  return value instanceof FormulaError;
}

export function isFormulaErrorToken(value: unknown): value is FormulaErrorValue {
  return typeof value === "string" && (formulaErrorValues as readonly string[]).includes(value);
}

export function normalizeFormulaError(value: FormulaScalar): FormulaScalar {
  return isFormulaErrorToken(value) ? formulaError(value) : value;
}

export function serializeFormulaError(value: FormulaValue): FormulaValue {
  if (isFormulaErrorValue(value)) return serializeFormulaError(value.value);
  return isMatrix(value)
    ? value.map((row) =>
        row.map((cell): FormulaScalar => (isFormulaErrorValue(cell) ? cell.value : cell)),
      )
    : isFormulaErrorValue(value)
      ? value.value
      : value;
}
