import { FormulaValueError } from "../errors";
import { isFormulaErrorValue, isMatrix, type FormulaScalar, type FormulaValue } from "../value";

export type CriterionOperator = "=" | "<>" | "<" | "<=" | ">" | ">=";

type Criterion = {
  operator: CriterionOperator;
  operand: FormulaScalar;
};

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}

function numericValue(value: FormulaScalar): number | undefined {
  if (typeof value === "number") return value;
  if (typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value)))
    return Number(value);
  return undefined;
}

function isCriterionOperator(value: string): value is CriterionOperator {
  return ["=", "<>", "<", "<=", ">", ">="].some((operator) => value === operator);
}

function hasWildcard(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    if (value[index] === "~") index += 1;
    else if (value[index] === "*" || value[index] === "?") return true;
  }
  return false;
}

function wildcardRegExp(value: string): RegExp {
  let pattern = "";
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index]!;
    const next = value[index + 1];
    if (character === "~" && (next === "*" || next === "?" || next === "~")) {
      pattern += escapeRegExp(next);
      index += 1;
    } else if (character === "*") pattern += ".*";
    else if (character === "?") pattern += ".";
    else pattern += escapeRegExp(character);
  }
  return new RegExp(`^${pattern}$`, "iu");
}

function scalarText(value: FormulaScalar): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

export function parseCriterion(criterion: FormulaScalar): Criterion {
  if (typeof criterion === "string" && criterion !== "") {
    const match = /^(<>|<=|>=|=|<|>)(.*)$/u.exec(criterion);
    if (match && isCriterionOperator(match[1]!)) {
      const operandText = match[2] ?? "";
      return {
        operator: match[1],
        operand: operandText === "" ? null : parseCriterion(operandText).operand,
      };
    }
  }
  if (criterion === null || criterion === "") return { operator: "=", operand: null };
  if (typeof criterion === "boolean") return { operator: "=", operand: criterion };
  if (typeof criterion === "string") {
    const upper = criterion.toUpperCase();
    if (upper === "TRUE" || upper === "FALSE") return { operator: "=", operand: upper === "TRUE" };
    const number = numericValue(criterion);
    if (number !== undefined) return { operator: "=", operand: number };
    return { operator: "=", operand: criterion };
  }
  return { operator: "=", operand: criterion };
}

export function scalarCriterion(criterion: FormulaValue): FormulaScalar {
  return isMatrix(criterion) ? (criterion.flat()[0] ?? null) : criterion;
}

function equalValue(value: FormulaScalar, expected: FormulaScalar): boolean {
  if (typeof expected === "number") {
    if (typeof value === "boolean") return false;
    return numericValue(value) === expected;
  }
  if (typeof expected === "boolean")
    return typeof value === "boolean"
      ? value === expected
      : scalarText(value) === scalarText(expected);
  const actual = scalarText(value);
  const target = scalarText(expected);
  return hasWildcard(target)
    ? wildcardRegExp(target).test(actual)
    : actual.localeCompare(target, undefined, { sensitivity: "base" }) === 0;
}

function compareValue(value: FormulaScalar, expected: FormulaScalar): number | undefined {
  if (typeof expected === "number") {
    if (typeof value === "boolean") return undefined;
    const actual = numericValue(value);
    return actual === undefined ? undefined : actual - expected;
  }
  if (typeof expected === "string" && typeof value === "string")
    return value.localeCompare(expected, undefined, { sensitivity: "base" });
  return undefined;
}

export function matchesCriterion(value: FormulaValue, criterion: FormulaValue): boolean {
  const scalar = isMatrix(value) ? (value.flat()[0] ?? null) : value;
  const criterionScalar = isMatrix(criterion) ? (criterion.flat()[0] ?? null) : criterion;
  if (isFormulaErrorValue(criterion)) throw new FormulaValueError(criterion.value);
  if (isFormulaErrorValue(criterionScalar)) throw new FormulaValueError(criterionScalar.value);
  const { operator, operand } = parseCriterion(criterionScalar);
  if (operator === "=" && isFormulaErrorValue(operand)) throw new FormulaValueError(operand.value);
  if (operand === null) {
    const blank = scalar === null || scalar === "";
    return operator === "<>" ? !blank : blank;
  }
  if (operator === "=") return equalValue(scalar, operand);
  if (operator === "<>") return !equalValue(scalar, operand);
  const comparison = compareValue(scalar, operand);
  if (comparison === undefined) return false;
  if (operator === "<") return comparison < 0;
  if (operator === "<=") return comparison <= 0;
  if (operator === ">") return comparison > 0;
  return comparison >= 0;
}
