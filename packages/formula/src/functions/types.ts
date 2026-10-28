import type { FormulaAst } from "../ast";
import type { FormulaFunctionExternalContext } from "../catalog-types";
import type { FormulaEvaluationContext } from "../context";
import type { FormulaValue } from "../value";

export type FormulaFunctionImplementation = (
  args: FormulaValue[],
  context: FormulaEvaluationContext,
) => FormulaValue;

export type FormulaFunctionAstEvaluator = (
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
) => FormulaValue;

export interface FormulaFunctionMetadata {
  minArgs: number;
  maxArgs: number | "variadic";
  astOnly?: boolean;
  acceptsErrors?: boolean;
  volatile?: boolean;
  dynamicDependencies?: boolean;
  broadcast?: boolean;
  hiddenRowSensitive?: boolean;
  externalContext?: FormulaFunctionExternalContext;
}

export type RegisteredFunction = FormulaFunctionImplementation & {
  readonly metadata: FormulaFunctionMetadata;
  readonly evaluateAst?: FormulaFunctionAstEvaluator;
};
