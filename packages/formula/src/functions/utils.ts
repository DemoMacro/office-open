import { FormulaValueError } from "../errors";
import { isFormulaErrorValue, isMatrix, maximumFormulaCells, type FormulaValue } from "../value";
import type {
  FormulaFunctionImplementation,
  FormulaFunctionMetadata,
  RegisteredFunction,
} from "./types";

export type FormulaFunctionTable = ReadonlyMap<string, FormulaFunctionImplementation>;

export function define(
  metadata: FormulaFunctionMetadata,
  implementation: FormulaFunctionImplementation,
): RegisteredFunction {
  return Object.assign(implementation, { metadata });
}

export function scalarNumber(value: FormulaValue): number {
  const scalar = Array.isArray(value) ? value.flat()[0] : value;
  if (scalar === null) return 0;
  if (typeof scalar === "object")
    throw new FormulaValueError(isFormulaErrorValue(scalar) ? scalar.value : "#VALUE!");
  if (typeof scalar === "number") return scalar;
  if (typeof scalar === "boolean") return scalar ? 1 : 0;
  if (typeof scalar === "string" && scalar.trim() !== "" && Number.isFinite(Number(scalar)))
    return Number(scalar);
  throw new FormulaValueError("#VALUE!");
}

export function scalarText(value: FormulaValue): string {
  const scalar = Array.isArray(value) ? value.flat()[0] : value;
  if (scalar === null) return "";
  if (typeof scalar === "object")
    throw new FormulaValueError(isFormulaErrorValue(scalar) ? scalar.value : "#VALUE!");
  if (typeof scalar === "string") return scalar;
  if (typeof scalar === "number") return String(scalar);
  if (typeof scalar === "boolean") return scalar ? "TRUE" : "FALSE";
  throw new FormulaValueError("#VALUE!");
}

export function scalarBoolean(value: FormulaValue): boolean {
  const scalar = Array.isArray(value) ? value.flat()[0] : value;
  if (typeof scalar === "boolean") return scalar;
  if (typeof scalar === "number") return scalar !== 0;
  if (scalar === null) return false;
  const text = scalarText(scalar).toUpperCase();
  if (text === "TRUE") return true;
  if (text === "FALSE") return false;
  throw new FormulaValueError("#VALUE!");
}

export function values(value: FormulaValue): FormulaValue[] {
  return isMatrix(value) ? value.flat() : [value];
}

export function numbers(value: FormulaValue): number[] {
  if (isMatrix(value)) {
    return value
      .flat()
      .filter((entry) => typeof entry === "number" || isFormulaErrorValue(entry))
      .map(scalarNumber);
  }
  if (value === null) return [];
  if (typeof value === "boolean" || typeof value === "string") {
    const number = scalarNumber(value);
    return Number.isFinite(number) ? [number] : [];
  }
  return [scalarNumber(value)];
}

export function numericSeries(value: FormulaValue): Array<number | undefined> {
  return values(value).map((entry) => {
    if (entry === null) return undefined;
    if (typeof entry === "string" && (entry.trim() === "" || !Number.isFinite(Number(entry))))
      return undefined;
    return scalarNumber(entry);
  });
}

export function pairedNumbers(left: FormulaValue, right: FormulaValue): Array<[number, number]> {
  const leftValues = numericSeries(left);
  const rightValues = numericSeries(right);
  const pairs: Array<[number, number]> = [];
  for (const [index, leftValue] of leftValues.entries()) {
    const rightValue = rightValues[index];
    if (leftValue !== undefined && rightValue !== undefined) pairs.push([leftValue, rightValue]);
  }
  return pairs;
}

export function isError(value: FormulaValue): boolean {
  return firstError([value]) !== undefined;
}

export function firstError(functionArguments: readonly FormulaValue[]): FormulaValue | undefined {
  return functionArguments
    .flatMap((value) => (isMatrix(value) ? value.flat() : [value]))
    .find(isFormulaErrorValue);
}

export function assertFormulaCellCount(rows: number, columns: number): void {
  if (!Number.isInteger(rows) || !Number.isInteger(columns) || rows < 1 || columns < 1)
    throw new FormulaValueError("#VALUE!");
  if (rows * columns > maximumFormulaCells) throw new FormulaValueError("#VALUE!");
}
