import {
  type FormulaAst,
  type FormulaBinaryOperator,
  type FormulaCellCoordinate,
  type FormulaReferenceArea,
  type FormulaReferenceAst,
  type FormulaTableSelector,
} from "./ast";
import {
  createFormulaContext,
  type FormulaCellEntry,
  type FormulaLambdaDefinition,
  type FormulaEvaluationContext,
  type FormulaEvaluationOptions,
  type FormulaWorkbookContext,
} from "./context";
import { FormulaValueError } from "./errors";
import { getFunction, setFunctionAstEvaluator } from "./functions";
import type { RegisteredFunction } from "./functions/types";
import { assertFormulaCellCount, firstError, scalarBoolean } from "./functions/utils";
import { sqlRequest } from "./functions/web-host";
import { collectDependencies, formatFormulaDependencyTarget } from "./graph";
import { parseFormula } from "./parser";
import { isWholeColumn, isWholeRow, maximumColumns, maximumRows } from "./references";
import { maximumFormulaCells } from "./value";
import {
  type FormulaMatrix,
  type FormulaScalar,
  type FormulaErrorValue,
  formulaError,
  isFormulaErrorValue,
  isFormulaErrorToken,
  normalizeFormulaError,
  serializeFormulaError,
  isMatrix,
  type FormulaValue,
} from "./value";

export interface FormulaEvaluationResult {
  value: FormulaValue;
  dependencies: string[];
}

function flatten(value: FormulaValue): FormulaScalar[] {
  if (typeof value === "string" && value.startsWith(lambdaPrefix))
    throw new FormulaValueError("#VALUE!");
  return isMatrix(value) ? value.flat() : [value];
}

function firstScalar(value: FormulaValue): FormulaScalar {
  return flatten(value)[0] ?? null;
}

function toNumber(value: FormulaValue): number {
  const scalar = firstScalar(value);
  if (typeof scalar === "number") return scalar;
  if (typeof scalar === "boolean") return scalar ? 1 : 0;
  if (typeof scalar === "string") {
    const parsed = Number(scalar.trim());
    if (scalar.trim() !== "" && Number.isFinite(parsed)) return parsed;
  }
  if (scalar === null) return 0;
  throw new FormulaValueError("#VALUE!");
}

function toText(value: FormulaValue): string {
  const scalar = firstScalar(value);
  if (typeof scalar === "string") return scalar;
  if (typeof scalar === "number") return String(scalar);
  if (typeof scalar === "boolean") return scalar ? "TRUE" : "FALSE";
  if (scalar === null) return "";
  throw new FormulaValueError(
    isFormulaErrorValue(scalar) ? scalar.value : "#VALUE!",
    isFormulaErrorValue(scalar) ? scalar.value : undefined,
  );
}

function typeRank(value: FormulaScalar): number {
  if (typeof value === "boolean") return 3;
  if (typeof value === "string") return 2;
  if (typeof value === "number") return 1;
  return 0;
}

function compareScalars(left: FormulaScalar, right: FormulaScalar): number {
  if (left === null || right === null) {
    const other = left === null ? right : left;
    if (other === "") return 0;
    const leftNumber = left === null ? 0 : coerceComparableNumber(left);
    const rightNumber = right === null ? 0 : coerceComparableNumber(right);
    if (leftNumber !== undefined && rightNumber !== undefined)
      return Math.sign(leftNumber - rightNumber);
  }
  const leftRank = typeRank(left);
  const rightRank = typeRank(right);
  if (leftRank !== rightRank) return Math.sign(leftRank - rightRank);
  if (leftRank === 3) return Math.sign(Number(left) - Number(right));
  if (leftRank === 2 && typeof left === "string" && typeof right === "string")
    return Math.sign(
      left.toLowerCase().localeCompare(right.toLowerCase(), undefined, {
        sensitivity: "base",
      }),
    );
  if (leftRank === 1 && typeof left === "number" && typeof right === "number")
    return Math.sign(left - right);
  return 0;
}

function coerceComparableNumber(value: FormulaScalar): number | undefined {
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return Number(value);
  if (typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value)))
    return Number(value);
  return undefined;
}

export function evaluateFormulaAst(
  ast: FormulaAst,
  workbook: FormulaWorkbookContext,
  options: FormulaEvaluationOptions = {},
): FormulaValue {
  const baseContext = createFormulaContext(workbook, options);
  const context: FormulaEvaluationContext = {
    ...baseContext,
    invokeLambdaValue: (handle, args) =>
      invokeLambda(lambdaValue(handle, baseContext), [...args], baseContext, new Set([ast.raw])),
  };
  const evaluated = formulaResult(() => evaluateNode(ast, context, new Set([ast.raw])));
  const result =
    typeof evaluated === "string" &&
    evaluated.startsWith(lambdaPrefix) &&
    context.lambdaScope.has(evaluated)
      ? "#CALC!"
      : evaluated;
  return options.errors === "object" ? result : serializeFormulaError(result);
}

function evaluateNode(
  ast: FormulaAst,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (ast.kind === "number") return ast.value;
  if (ast.kind === "string") return ast.value;
  if (ast.kind === "boolean") return ast.value;
  if (ast.kind === "missing") return null;
  if (ast.kind === "error") return formulaError(ast.value);
  if (ast.kind === "reference") {
    const variable = context.variables?.get(ast.raw.toUpperCase());
    if (variable !== undefined) return variable;
    return evaluateReferenceAst(ast, context);
  }
  if (ast.kind === "name") return evaluateName(ast, context, visited);
  if (ast.kind === "group") return evaluateNode(ast.operand, context, visited);
  if (ast.kind === "unary") {
    return unaryScalars(evaluateNode(ast.operand, context, visited), (value) =>
      ast.operator === "-" ? -toNumber(value) : toNumber(value),
    );
  }
  if (ast.kind === "postfix")
    return unaryScalars(
      evaluateNode(ast.operand, context, visited),
      (value) => toNumber(value) / 100,
    );
  if (ast.kind === "array") {
    return ast.rows.map((row) =>
      row.map((cell) => firstScalar(evaluateNode(cell, context, visited))),
    );
  }
  if (ast.kind === "binary")
    return evaluateBinary(ast.operator, ast.left, ast.right, context, visited);
  if (ast.kind === "application") {
    const callee = evaluateNode(ast.callee, context, visited);
    if (!isLambdaValue(callee, context)) return formulaError("#VALUE!");
    const lambda = lambdaValue(callee, context);
    const arguments_ = ast.args.map((argument) => evaluateNode(argument, context, visited));
    return invokeLambda(lambda, arguments_, context, visited);
  }

  return evaluateCall(ast, context, visited);
}

function evaluateCall(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  const name = ast.name.toUpperCase();
  const lambdaVariable = context.variables?.get(name);
  if (lambdaVariable !== undefined && isLambdaValue(lambdaVariable, context)) {
    const lambda = lambdaValue(lambdaVariable, context);
    const lambdaArguments = ast.args.map((argument) => evaluateNode(argument, context, visited));
    return invokeLambda(lambda, lambdaArguments, context, visited);
  }
  const implementation = getFunction(ast.name);
  if (!implementation) return formulaError("#NAME?");
  const metadata = implementation.metadata;
  if (
    ast.args.length < metadata.minArgs ||
    (metadata.maxArgs !== "variadic" && ast.args.length > metadata.maxArgs)
  )
    return formulaError("#VALUE!");
  if (implementation.evaluateAst)
    return normalizeNumber(implementation.evaluateAst(ast, context, visited));
  const args = ast.args.map((argument) => evaluateNode(argument, context, visited));
  const argumentError = firstError(args);
  if (argumentError !== undefined && !metadata.acceptsErrors) return argumentError;
  try {
    return metadata.broadcast
      ? broadcastCall(implementation, args, context)
      : normalizeNumber(implementation(args, context));
  } catch (error) {
    if (error instanceof FormulaValueError) return formulaError(error.value);
    throw error;
  }
}

function normalizeNumber(value: FormulaValue): FormulaValue {
  if (isMatrix(value)) {
    return value.map((row) => row.map(normalizeScalar));
  }
  return normalizeScalar(value);
}

function normalizeScalar(value: FormulaScalar): FormulaScalar {
  if (typeof value === "number" && !Number.isFinite(value)) return formulaError("#NUM!");
  return normalizeFormulaError(value);
}

function broadcastDimensions(
  values: readonly FormulaValue[],
): { rows: number; columns: number } | undefined {
  const matrices = values.filter(isMatrix);
  if (matrices.length === 0) return undefined;
  const heights = new Set(matrices.map((matrix) => (matrix.length === 1 ? 1 : matrix.length)));
  const widths = new Set(
    matrices.map((matrix) => (matrix[0]?.length === 1 ? 1 : matrix[0]?.length)),
  );
  if (heights.size > 2 || widths.size > 2) throw new FormulaValueError("#VALUE!");
  const rowValues = [...heights];
  const columnValues = [...widths];
  if (rowValues.length === 2 && !rowValues.includes(1)) throw new FormulaValueError("#VALUE!");
  if (columnValues.length === 2 && !columnValues.includes(1))
    throw new FormulaValueError("#VALUE!");
  return {
    rows: Math.max(...(rowValues.length ? rowValues : [1])),
    columns: Math.max(
      ...(columnValues.length && columnValues[0] !== undefined ? columnValues : [1]),
    ),
  };
}

function broadcastScalar(value: FormulaValue, row: number, column: number): FormulaScalar {
  if (!isMatrix(value)) return value;
  const sourceRow = value.length === 1 ? 0 : row;
  const sourceColumn = value[sourceRow]?.length === 1 ? 0 : column;
  return value[sourceRow]?.[sourceColumn] ?? null;
}

function broadcastCall(
  implementation: RegisteredFunction,
  args: readonly FormulaValue[],
  context: FormulaEvaluationContext,
): FormulaValue {
  const dimensions = broadcastDimensions(args);
  if (!dimensions) return normalizeNumber(implementation([...args], context));
  assertFormulaCellCount(dimensions.rows, dimensions.columns);
  const output: FormulaMatrix = Array.from<unknown, FormulaScalar[]>(
    { length: dimensions.rows },
    (_, row) => {
      const matrixRow = Array.from<unknown, FormulaScalar>(
        { length: dimensions.columns },
        (_, column) => {
          const broadcastArguments = args.map((value) => {
            if (!isMatrix(value)) return value;
            const sourceRow = value.length === 1 ? 0 : row;
            const sourceColumn = value[sourceRow]?.length === 1 ? 0 : column;
            return value[sourceRow]?.[sourceColumn] ?? null;
          });
          const argumentError = firstError(broadcastArguments);
          if (argumentError !== undefined && !implementation.metadata.acceptsErrors)
            return isMatrix(argumentError) ? "#VALUE!" : argumentError;
          try {
            const evaluated = normalizeNumber(implementation(broadcastArguments, context));
            if (isMatrix(evaluated)) throw new FormulaValueError("#VALUE!");
            return evaluated;
          } catch (error) {
            if (error instanceof FormulaValueError) return formulaError(error.value);
            throw error;
          }
        },
      );
      return matrixRow;
    },
  );
  return output;
}

function childContext(
  context: FormulaEvaluationContext,
  variables: ReadonlyMap<string, FormulaValue>,
): FormulaEvaluationContext {
  return { ...context, variables };
}

function referencePosition(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  name: "ROW" | "COLUMN",
): FormulaValue {
  const argument = ast.args[0];
  if (argument === undefined || argument.kind === "missing") {
    const fallback = name === "ROW" ? context.row : context.column;
    return fallback ?? "#VALUE!";
  }
  if (argument.kind !== "reference") return "#VALUE!";
  const bounds = referenceBounds(argument, context);
  if (!bounds) return "#REF!";
  if (name === "ROW") {
    return Array.from({ length: bounds.lastRow - bounds.firstRow + 1 }, (_, index) => [
      bounds.firstRow + index,
    ]);
  }
  return [
    Array.from({ length: bounds.lastColumn - bounds.firstColumn + 1 }, (_, index) => {
      return bounds.firstColumn + index;
    }),
  ];
}

function referenceBounds(
  ast: FormulaReferenceAst,
  context: FormulaEvaluationContext,
): { firstRow: number; firstColumn: number; lastRow: number; lastColumn: number } | undefined {
  const area = ast.reference?.areas?.[0];
  if (area) return areaBounds(area, context);
  const start = ast.reference?.start;
  if (!start) return undefined;
  return areaBounds({ start, ...(ast.reference?.end ? { end: ast.reference.end } : {}) }, context);
}

function evaluateIndirect(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (ast.args.length < 1 || ast.args.length > 2) return "#VALUE!";
  const value = evaluateNode(ast.args[0]!, context, visited);
  const source = firstScalar(value);
  if (typeof source !== "string" || source.trim() === "") return "#REF!";
  if (ast.args[1] !== undefined) {
    const style = firstScalar(evaluateNode(ast.args[1]!, context, visited));
    const usesA1 =
      style === null ||
      style === true ||
      (typeof style === "number" && style !== 0) ||
      (typeof style === "string" && style.toUpperCase() === "TRUE");
    if (!usesA1) return evaluateR1C1(source.trim(), context);
  }
  let target: FormulaAst;
  try {
    target = parseFormula(source.trim(), { dialect: context.dialect ?? "excel" });
  } catch {
    return "#REF!";
  }
  if (target.kind !== "reference" && target.kind !== "name") return "#REF!";
  return evaluateNode(target, context, visited);
}

function evaluateR1C1(source: string, context: FormulaEvaluationContext): FormulaValue {
  const sheetSeparator = source.lastIndexOf("!");
  const sheetName = sheetSeparator === -1 ? undefined : source.slice(0, sheetSeparator);
  const reference = sheetSeparator === -1 ? source : source.slice(sheetSeparator + 1);
  const match =
    /^(?:(R(?:\[-?\d+\]|\d+)?)(?::R(?:\[-?\d+\]|\d+)?)?)(?:(C(?:\[-?\d+\]|\d+)?)(?::C(?:\[-?\d+\]|\d+)?)?)$/u.exec(
      reference,
    );
  if (!match) return "#REF!";
  const rows = (match[1] ?? "")
    .split(":")
    .filter(Boolean)
    .map((edge) => r1c1Coordinate(edge, "R", context));
  const columns = (match[2] ?? "")
    .split(":")
    .filter(Boolean)
    .map((edge) => r1c1Coordinate(edge, "C", context));
  if (rows.includes(undefined) || columns.includes(undefined)) return "#REF!";
  const startRow = rows[0]!;
  const endRow = rows.at(-1)!;
  const startColumn = columns[0]!;
  const endColumn = columns.at(-1)!;
  return readArea(
    {
      ...(sheetName === undefined ? {} : { fromSheet: sheetName.replace(/^'|'$/gu, "") }),
      start: {
        row: Math.min(startRow, endRow),
        column: Math.min(startColumn, endColumn),
        absoluteRow: true,
        absoluteColumn: true,
      },
      end: {
        row: Math.max(startRow, endRow),
        column: Math.max(startColumn, endColumn),
        absoluteRow: true,
        absoluteColumn: true,
      },
    },
    context,
  );
}

function r1c1Coordinate(
  token: string,
  axis: "R" | "C",
  context: FormulaEvaluationContext,
): number | undefined {
  const body = token.slice(1);
  if (body === "") return axis === "R" ? context.row : context.column;
  if (body.startsWith("[")) {
    const offset = Number(body.slice(1, -1));
    const base = axis === "R" ? context.row : context.column;
    if (!Number.isFinite(offset) || base === undefined) return undefined;
    return base + offset;
  }
  const value = Number(body);
  return Number.isInteger(value) && value >= 1 ? value : undefined;
}

function evaluateOffset(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  const reference = ast.args[0];
  if (reference === undefined || reference.kind === "missing") return "#REF!";
  if (reference.kind !== "reference") return "#VALUE!";
  const bounds = referenceBounds(reference, context);
  if (!bounds) return "#REF!";
  const baseHeight = bounds.lastRow - bounds.firstRow + 1;
  const baseWidth = bounds.lastColumn - bounds.firstColumn + 1;
  const argumentValue = (position: number) => {
    const argument = ast.args[position];
    return argument === undefined || argument.kind === "missing"
      ? undefined
      : evaluateNode(argument, context, visited);
  };
  const rows = Math.trunc(toNumber(argumentValue(1) ?? 0));
  const columns = Math.trunc(toNumber(argumentValue(2) ?? 0));
  const height = Math.trunc(toNumber(argumentValue(3) ?? baseHeight));
  const width = Math.trunc(toNumber(argumentValue(4) ?? baseWidth));
  if (height <= 0 || width <= 0) return "#REF!";
  const firstRow = bounds.firstRow + rows;
  const firstColumn = bounds.firstColumn + columns;
  const lastRow = firstRow + height - 1;
  const lastColumn = firstColumn + width - 1;
  const area = reference.reference?.areas?.[0];
  const sheetId =
    area?.fromSheet === undefined
      ? (context.sheet ?? 0)
      : context.workbook.getSheetId(area.fromSheet);
  if (sheetId === undefined) return "#REF!";
  if (
    firstRow < 1 ||
    firstColumn < 1 ||
    lastRow > context.workbook.getRowCount(sheetId) ||
    lastColumn > context.workbook.getColumnCount(sheetId)
  )
    return "#REF!";
  return readArea(
    {
      ...(area?.fromSheet === undefined ? {} : { fromSheet: area.fromSheet }),
      start: { row: firstRow, column: firstColumn, absoluteRow: true, absoluteColumn: true },
      end: { row: lastRow, column: lastColumn, absoluteRow: true, absoluteColumn: true },
    },
    context,
  );
}

function cellFormulaText(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
): FormulaValue {
  const reference = ast.args[0];
  if (reference === undefined || reference.kind === "missing") return "#N/A";
  if (reference.kind !== "reference") return "#VALUE!";
  const bounds = referenceBounds(reference, context);
  if (!bounds || bounds.firstRow !== bounds.lastRow || bounds.firstColumn !== bounds.lastColumn)
    return "#VALUE!";
  const area = reference.reference?.areas?.[0];
  const sheetId =
    area?.fromSheet === undefined
      ? (context.sheet ?? 0)
      : context.workbook.getSheetId(area.fromSheet);
  if (sheetId === undefined) return "#REF!";
  const snapshot = context.workbook.getCell(sheetId, bounds.firstRow, bounds.firstColumn);
  return snapshot?.formula ?? "#N/A";
}

const lambdaPrefix = "__FORMULA_LAMBDA__";
let nextLambdaId = 0;

function createLambda(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
): FormulaValue {
  if (ast.args.length < 1) return "#VALUE!";
  const parameterNodes = ast.args.slice(0, -1);
  const parameters = parameterNodes.map((node) => {
    if (node.kind === "name") return node.name.toUpperCase();
    if (node.kind === "reference" && /^[A-Za-z_][A-Za-z0-9_.]*$/u.test(node.raw))
      return node.raw.toUpperCase();
    throw new FormulaValueError("#VALUE!");
  });
  const handle = `${lambdaPrefix}${nextLambdaId++}`;
  context.lambdaScope.set(handle, {
    parameterNames: parameters,
    body: ast.args.at(-1)!,
    variables: context.variables ?? new Map(),
  });
  return handle;
}

function evaluateLet(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (ast.args.length < 3 || ast.args.length % 2 === 0) return "#VALUE!";
  let variables = context.variables ?? new Map<string, FormulaValue>();
  for (let index = 0; index < ast.args.length - 1; index += 2) {
    const nameNode = ast.args[index]!;
    const name =
      nameNode.kind === "name"
        ? nameNode.name
        : nameNode.kind === "reference" && /^[A-Za-z_][A-Za-z0-9_.]*$/u.test(nameNode.raw)
          ? nameNode.raw
          : undefined;
    if (name === undefined) return "#VALUE!";
    const value = evaluateNode(ast.args[index + 1]!, childContext(context, variables), visited);
    variables = new Map([...variables, [name.toUpperCase(), value]]);
  }
  return evaluateNode(ast.args.at(-1)!, childContext(context, variables), visited);
}

function isLambdaValue(value: FormulaValue, context: FormulaEvaluationContext): value is string {
  return (
    typeof value === "string" && value.startsWith(lambdaPrefix) && context.lambdaScope.has(value)
  );
}

function lambdaHandle(
  argument: FormulaAst,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  const value = evaluateNode(argument, context, visited);
  return isLambdaValue(value, context) ? value : "#VALUE!";
}

function lambdaValue(handle: string, context: FormulaEvaluationContext): FormulaLambdaDefinition {
  const lambda = context.lambdaScope.get(handle);
  if (!lambda) throw new FormulaValueError("#VALUE!");
  return lambda;
}

function scalarLambdaResult(value: FormulaValue): FormulaScalar {
  if (isMatrix(value)) throw new FormulaValueError("#VALUE!");
  if (typeof value === "string" && value.startsWith(lambdaPrefix))
    throw new FormulaValueError("#VALUE!");
  return value;
}

function invokeLambda(
  lambda: FormulaLambdaDefinition,
  arguments_: FormulaValue[],
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (lambda.parameterNames.length !== arguments_.length) return formulaError("#VALUE!");
  const variables = new Map<string, FormulaValue>([
    ...(context.variables ?? new Map<string, FormulaValue>()),
    ...lambda.variables,
    ...lambda.parameterNames.map((name, index): [string, FormulaValue] => [
      name,
      arguments_[index] ?? null,
    ]),
  ]);
  return evaluateNode(lambda.body, childContext(context, variables), visited);
}

function formulaResult(operation: () => FormulaValue): FormulaValue {
  try {
    return operation();
  } catch (error) {
    if (error instanceof FormulaValueError) return formulaError(error.value);
    throw error;
  }
}

function evaluateIf(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (ast.args.length < 2 || ast.args.length > 3) return "#VALUE!";
  const condition = evaluateNode(ast.args[0]!, context, visited);
  const conditionError = firstError([condition]);
  if (conditionError !== undefined) return conditionError;
  if (isMatrix(condition)) {
    const selectedBranches = new Set<FormulaAst>();
    for (const selected of condition.flat()) {
      const branch = evaluateIfBranchAst(ast, scalarBoolean(selected));
      if (branch !== undefined && branch.kind !== "missing") selectedBranches.add(branch);
    }
    const branchResults = new Map<FormulaAst, FormulaValue>();
    for (const branch of selectedBranches)
      branchResults.set(
        branch,
        formulaResult(() => evaluateNode(branch, context, visited)),
      );
    const dimensions = broadcastDimensions([condition, ...branchResults.values()]);
    if (!dimensions) return condition;
    assertFormulaCellCount(dimensions.rows, dimensions.columns);
    return Array.from({ length: dimensions.rows }, (_, row) =>
      Array.from({ length: dimensions.columns }, (_, column): FormulaScalar => {
        const selected = broadcastScalar(condition, row, column);
        const isSelected = scalarBoolean(selected);
        const branch = evaluateIfBranchAst(ast, isSelected);
        if (branch === undefined || branch.kind === "missing") return isSelected ? 0 : false;
        const branchResult = branchResults.get(branch);
        return branchResult === undefined ? null : broadcastScalar(branchResult, row, column);
      }),
    );
  }
  return formulaResult(() => {
    return evaluateIfBranch(ast, scalarBoolean(condition), context, visited);
  });
}

function evaluateIfBranch(
  ast: Extract<FormulaAst, { kind: "call" }>,
  selected: boolean,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  const branch = evaluateIfBranchAst(ast, selected);
  if (branch === undefined) return false;
  if (branch.kind === "missing") return selected ? 0 : false;
  return evaluateNode(branch, context, visited);
}

function evaluateIfBranchAst(
  ast: Extract<FormulaAst, { kind: "call" }>,
  selected: boolean,
): FormulaAst | undefined {
  return selected ? ast.args[1] : ast.args[2];
}

function evaluateIfError(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
  onlyNa: boolean,
): FormulaValue {
  if (ast.args.length !== 2) return "#VALUE!";
  const value = evaluateNode(ast.args[0]!, context, visited);
  const error = firstError([value]);
  const errorValue = isFormulaErrorValue(error) ? error.value : undefined;
  const shouldFallback = errorValue !== undefined && (onlyNa ? errorValue === "#N/A" : true);
  if (isMatrix(value)) {
    if (!shouldFallback) return value;
    const fallbackResult = formulaResult(() => evaluateNode(ast.args[1]!, context, visited));
    const dimensions = broadcastDimensions([value, fallbackResult]);
    if (!dimensions) return fallbackResult;
    assertFormulaCellCount(dimensions.rows, dimensions.columns);
    return Array.from({ length: dimensions.rows }, (_, row) =>
      Array.from({ length: dimensions.columns }, (_, column): FormulaScalar => {
        const cell = broadcastScalar(value, row, column);
        const cellError = isFormulaErrorValue(cell) ? cell : undefined;
        const useFallback = cellError !== undefined && (!onlyNa || cellError.value === "#N/A");
        return useFallback ? broadcastScalar(fallbackResult, row, column) : cell;
      }),
    );
  }
  return shouldFallback ? evaluateNode(ast.args[1]!, context, visited) : value;
}

function evaluateChoose(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (ast.args.length < 2) return "#VALUE!";
  const indexValue = evaluateNode(ast.args[0]!, context, visited);
  const indexError = firstError([indexValue]);
  if (indexError !== undefined) return indexError;
  return formulaResult(() => {
    const index = Math.trunc(toNumber(indexValue));
    const selected = index >= 1 ? ast.args[index] : undefined;
    return selected === undefined ? "#VALUE!" : evaluateNode(selected, context, visited);
  });
}

function evaluateIfs(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (ast.args.length < 2 || ast.args.length % 2 !== 0) return "#VALUE!";
  for (let index = 0; index < ast.args.length; index += 2) {
    const condition = evaluateNode(ast.args[index]!, context, visited);
    const error = firstError([condition]);
    if (error !== undefined) return error;
    const matches = formulaResult(() => scalarBoolean(condition));
    if (isFormulaErrorValue(matches)) return matches;
    if (matches) return evaluateNode(ast.args[index + 1]!, context, visited);
  }
  return "#N/A";
}

function evaluateSwitch(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (ast.args.length < 3) return "#VALUE!";
  const expression = evaluateNode(ast.args[0]!, context, visited);
  const expressionError = firstError([expression]);
  if (expressionError !== undefined) return expressionError;
  const remainder = ast.args.slice(1);
  const hasDefault = remainder.length % 2 === 1;
  const pairCount = hasDefault ? (remainder.length - 1) / 2 : remainder.length / 2;
  for (let index = 0; index < pairCount; index += 1) {
    const value = evaluateNode(remainder[index * 2]!, context, visited);
    const valueError = firstError([value]);
    if (valueError !== undefined) return valueError;
    if (compareScalars(firstScalar(expression), firstScalar(value)) === 0)
      return evaluateNode(remainder[index * 2 + 1]!, context, visited);
  }
  return hasDefault ? evaluateNode(remainder.at(-1)!, context, visited) : "#N/A";
}

function evaluateAndOr(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
  isOr: boolean,
): FormulaValue {
  if (ast.args.length < 1) return "#VALUE!";
  let sawLogical = false;
  for (const argument of ast.args) {
    const value = evaluateNode(argument, context, visited);
    const error = firstError([value]);
    if (error !== undefined) return error;
    for (const scalar of flatten(value)) {
      if (scalar === null || typeof scalar === "string") continue;
      const result = formulaResult(() => scalarBoolean(scalar));
      if (isFormulaErrorValue(result)) return result;
      sawLogical = true;
      if (result === isOr) return result;
    }
  }
  return sawLogical ? !isOr : "#VALUE!";
}

function applyByDimension(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
  byColumn: boolean,
): FormulaValue {
  if (ast.args.length !== 2) return "#VALUE!";
  const input = evaluateNode(ast.args[0]!, context, visited);
  const handle = lambdaHandle(ast.args[1]!, context, visited);
  if (typeof handle !== "string" || !isMatrix(input)) return "#VALUE!";
  const lambda = lambdaValue(handle, context);
  const outer = byColumn ? (input[0] ?? []) : input;
  const output: FormulaScalar[] = outer.map((_, index) => {
    const slice: FormulaMatrix = byColumn
      ? input.map((row) => [row[index] ?? null])
      : [input[index] ?? []];
    const lambdaArguments: FormulaValue[] = [slice];
    return scalarLambdaResult(invokeLambda(lambda, lambdaArguments, context, visited));
  });
  const outputMatrix: FormulaMatrix = byColumn ? [output] : output.map((value) => [value]);
  return outputMatrix;
}

function applyMap(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (ast.args.length < 2) return "#VALUE!";
  const handle = lambdaHandle(ast.args.at(-1)!, context, visited);
  const matrices = ast.args.slice(0, -1).map((node) => {
    const value = evaluateNode(node, context, visited);
    return isMatrix(value) ? value : [[value]];
  });
  if (typeof handle !== "string") return "#VALUE!";
  const lambda = lambdaValue(handle, context);
  const first = matrices[0]!;
  if (
    matrices.some(
      (matrix) => matrix.length !== first.length || matrix[0]?.length !== first[0]?.length,
    )
  )
    return "#VALUE!";
  const output: FormulaMatrix = first.map((row, rowIndex) =>
    row.map((_, columnIndex) =>
      scalarLambdaResult(
        invokeLambda(
          lambda,
          matrices.map((matrix) => matrix[rowIndex]?.[columnIndex] ?? null),
          context,
          visited,
        ),
      ),
    ),
  );
  return output;
}

function applyMakeArray(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (ast.args.length !== 3) return "#VALUE!";
  const rows = toNumber(evaluateNode(ast.args[0]!, context, visited));
  const columns = toNumber(evaluateNode(ast.args[1]!, context, visited));
  if (!Number.isInteger(rows) || !Number.isInteger(columns) || rows < 1 || columns < 1)
    return "#VALUE!";
  try {
    assertFormulaCellCount(rows, columns);
  } catch {
    return "#VALUE!";
  }
  const handle = lambdaHandle(ast.args[2]!, context, visited);
  if (typeof handle !== "string") return "#VALUE!";
  const lambda = lambdaValue(handle, context);
  const output: FormulaMatrix = Array.from({ length: rows }, (_, row) =>
    Array.from({ length: columns }, (_, column) =>
      scalarLambdaResult(invokeLambda(lambda, [row + 1, column + 1], context, visited)),
    ),
  );
  return output;
}

function applyReduction(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
  scan: boolean,
): FormulaValue {
  if (ast.args.length !== 3) return "#VALUE!";
  let accumulator = evaluateNode(ast.args[0]!, context, visited);
  const handle = lambdaHandle(ast.args[2]!, context, visited);
  if (typeof handle !== "string") return "#VALUE!";
  const lambda = lambdaValue(handle, context);
  const input = evaluateNode(ast.args[1]!, context, visited);
  const values = isMatrix(input) ? input.flat() : [input];
  const output: FormulaScalar[] = values.map((value) => {
    accumulator = scalarLambdaResult(invokeLambda(lambda, [accumulator, value], context, visited));
    return scalarLambdaResult(accumulator);
  });
  const outputMatrix: FormulaMatrix = output.map((value) => [value]);
  return scan ? outputMatrix : accumulator;
}

function evaluateName(
  ast: Extract<FormulaAst, { kind: "name" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  const variable = context.variables?.get(ast.name.toUpperCase());
  if (variable !== undefined) return variable;
  const formula =
    context.workbook.getDefinedFormula(ast.name, context.sheet) ??
    context.workbook.getDefinedFormula(ast.name.toUpperCase(), context.sheet);
  if (!formula) {
    if (context.workbook.getTableName(ast.name)) {
      return evaluateTable({ table: ast.name, selectors: [] }, context, ast.implicitIntersection);
    }
    return evaluateReference(ast.raw, context);
  }
  if (visited.has(formula)) return "#VALUE!";
  const defined = parseFormula(formula, { dialect: context.dialect ?? "excel" });
  const evaluated =
    ast.implicitIntersection && defined.kind === "reference"
      ? {
          ...defined,
          reference: { ...defined.reference, implicitIntersection: true },
        }
      : defined;
  return evaluateNode(evaluated, context, new Set([...visited, formula]));
}

function evaluateReferenceAst(
  ast: Extract<FormulaAst, { kind: "reference" }>,
  context: FormulaEvaluationContext,
): FormulaValue {
  const reference = ast.reference;
  if (reference?.table)
    return evaluateTable(reference.table, context, reference.implicitIntersection);
  if (reference?.spill && reference.areas?.[0]) {
    const anchor = reference.areas[0];
    if (!anchor) return "#REF!";
    const sheet = areaEvaluationContext(anchor, context);
    if (sheet === undefined) return "#REF!";
    const anchorSheet = sheet.workbook.getSheetId(anchor.fromSheet ?? "") ?? context.sheet ?? 0;
    const start = edgeSpan(anchor.start, sheet, anchorSheet);
    if (!start || !sheet.workbook.getSpillRange) return "#REF!";
    const range = sheet.workbook.getSpillRange(anchorSheet, start.firstRow, start.firstColumn);
    if (!range) return "#REF!";
    return readArea(
      {
        ...(anchor.fromSheet === undefined ? {} : { fromSheet: anchor.fromSheet }),
        start: {
          row: range.firstRow,
          column: range.firstColumn,
          absoluteRow: true,
          absoluteColumn: true,
        },
        end: {
          row: range.lastRow,
          column: range.lastColumn,
          absoluteRow: true,
          absoluteColumn: true,
        },
      },
      sheet,
    );
  }
  if (reference?.areas?.length) {
    let totalCells = 0;
    for (const area of reference.areas) {
      const cellCount = areaCellCount(area, context);
      if (cellCount === undefined) return "#REF!";
      totalCells += cellCount;
      if (totalCells > maximumFormulaCells) throw new FormulaValueError("#VALUE!");
    }
    assertFormulaCellCount(totalCells, 1);
    if (reference.implicitIntersection) return intersectAreaAtCaller(reference.areas[0]!, context);
    const matrices = reference.areas.map((area) => readArea(area, context));
    return matrices.length === 1 ? matrices[0]! : matrices.reduce(unionMatrix);
  }
  if (reference?.start) {
    const start = reference.start;
    const end = reference.end ?? start;
    const sheet = context.sheet ?? 0;
    return readCells(sheet, start.row, start.column, end.row, end.column, context);
  }
  return evaluateReference(ast.raw, context);
}

interface FormulaTableSelection {
  sheetId: number;
  firstRow: number;
  lastRow: number;
  columns: [from: number, to: number][];
}

function resolveTableSelection(
  selector: FormulaTableSelector,
  context: FormulaEvaluationContext,
  implicitIntersection = false,
): { selection?: FormulaTableSelection; error?: FormulaScalar } {
  const table = context.workbook.getTableName(selector.table);
  if (!table) return { error: "#NAME?" };
  const item = selector.item ?? (implicitIntersection ? "this-row" : undefined);
  const sheetId = table.sheetId;
  const bodyFirst = table.range.firstRow + (table.hasHeaders ? 1 : 0);
  const bodyLast = table.range.lastRow - (table.hasTotals ? 1 : 0);
  if ((item === "headers" && !table.hasHeaders) || (item === "totals" && !table.hasTotals))
    return { error: "#REF!" };
  let firstRow: number;
  let lastRow: number;
  if (item === "headers") {
    firstRow = table.range.firstRow;
    lastRow = firstRow;
  } else if (item === "totals") {
    firstRow = table.range.lastRow;
    lastRow = firstRow;
  } else if (item === "this-row") {
    if (
      context.row === undefined ||
      (context.sheet ?? 0) !== sheetId ||
      context.row < bodyFirst ||
      context.row > bodyLast
    )
      return { error: "#VALUE!" };
    firstRow = context.row;
    lastRow = context.row;
  } else if (item === "all") {
    firstRow = table.range.firstRow;
    lastRow = table.range.lastRow;
  } else {
    firstRow = bodyFirst;
    lastRow = bodyLast;
    if (firstRow > lastRow) return { error: "#REF!" };
  }
  const resolveColumn = (column: string): number => {
    const index = table.columnNames?.findIndex(
      (name) => name.toLowerCase() === column.toLowerCase(),
    );
    return index === undefined || index < 0 ? 0 : table.range.firstColumn + index;
  };
  const selectedIndexes: [from: number, to: number][] = [];
  for (const selected of selector.selectors) {
    const indexes =
      selected.kind === "column"
        ? [resolveColumn(selected.column)]
        : [resolveColumn(selected.start), resolveColumn(selected.end)];
    if (indexes.some((index) => index === 0)) return { error: "#REF!" };
    selectedIndexes.push([Math.min(...indexes), Math.max(...indexes)]);
  }
  return {
    selection: {
      sheetId,
      firstRow,
      lastRow,
      columns: selectedIndexes.length
        ? selectedIndexes
        : [[table.range.firstColumn, table.range.lastColumn]],
    },
  };
}

function evaluateTable(
  selector: FormulaTableSelector,
  context: FormulaEvaluationContext,
  implicitIntersection = false,
): FormulaValue {
  const resolved = resolveTableSelection(selector, context, implicitIntersection);
  if (resolved.error !== undefined) return resolved.error;
  const selection = resolved.selection!;
  if (selection.columns.length > 1) {
    const matrices = selection.columns.map(([from, to]) =>
      readCells(selection.sheetId, selection.firstRow, from, selection.lastRow, to, context),
    );
    return matrices.reduce(unionTableMatrices);
  }
  const [firstColumn, lastColumn] = selection.columns[0]!;
  return readCells(
    selection.sheetId,
    selection.firstRow,
    firstColumn,
    selection.lastRow,
    lastColumn,
    context,
  );
}
function unionTableMatrices(left: FormulaValue, right: FormulaValue): FormulaValue {
  const leftMatrix = isMatrix(left) ? left : [[left]];
  const rightMatrix = isMatrix(right) ? right : [[right]];
  const cellCount = (matrix: FormulaMatrix) => matrix.reduce((total, row) => total + row.length, 0);
  assertFormulaCellCount(cellCount(leftMatrix) + cellCount(rightMatrix), 1);
  return leftMatrix.map((row, index) => [...row, ...(rightMatrix[index] ?? [])]);
}

function intersectAreaAtCaller(
  area: FormulaReferenceArea,
  context: FormulaEvaluationContext,
): FormulaValue {
  const row = context.row;
  const column = context.column;
  if (row === undefined || column === undefined) return formulaError("#VALUE!");
  const areaContext = areaEvaluationContext(area, context);
  if (areaContext === undefined) return formulaError("#REF!");
  if (area.fromSheet !== undefined) {
    const explicitSheet = areaContext.workbook.getSheetId(area.fromSheet);
    if (explicitSheet === undefined) return formulaError("#REF!");
    return intersectAreaWithSheet(area, areaContext, explicitSheet, context);
  }
  return intersectAreaWithSheet(
    area,
    areaContext,
    areaContext.workbook.getSheetId("") ?? context.sheet ?? 0,
    context,
  );
}

function intersectAreaWithSheet(
  area: FormulaReferenceArea,
  areaContext: FormulaEvaluationContext,
  sheet: number,
  context: FormulaEvaluationContext,
): FormulaValue {
  const row = context.row;
  const column = context.column;
  if (row === undefined || column === undefined) return formulaError("#VALUE!");
  const bounds = areaBounds(area, areaContext);
  if (!bounds) return formulaError("#REF!");
  if (bounds.firstColumn === bounds.lastColumn && row >= bounds.firstRow && row <= bounds.lastRow)
    return readCells(sheet, row, bounds.firstColumn, row, bounds.firstColumn, context);
  if (
    bounds.firstRow === bounds.lastRow &&
    column >= bounds.firstColumn &&
    column <= bounds.lastColumn
  )
    return readCells(sheet, bounds.firstRow, column, bounds.firstRow, column, context);
  return formulaError("#VALUE!");
}

function readArea(area: FormulaReferenceArea, context: FormulaEvaluationContext): FormulaValue {
  const evaluationContext = areaEvaluationContext(area, context);
  if (evaluationContext === undefined) return "#REF!";
  const sheetIds = areaSheetIds(area, evaluationContext);
  if (!sheetIds.length) return "#REF!";
  const cellCount = areaCellCount(area, context);
  if (cellCount === undefined) return "#REF!";
  assertFormulaCellCount(cellCount, 1);
  const matrices = sheetIds.map((sheet): FormulaValue => {
    const bounds = areaBounds(area, evaluationContext, sheet);
    if (!bounds) return formulaError("#REF!");
    return readCells(
      sheet,
      bounds.firstRow,
      bounds.firstColumn,
      bounds.lastRow,
      bounds.lastColumn,
      evaluationContext,
    );
  });
  return matrices.length === 1 ? matrices[0]! : matrices.reduce(unionMatrix);
}

function areaCellCount(
  area: FormulaReferenceArea,
  context: FormulaEvaluationContext,
): number | undefined {
  const evaluationContext = areaEvaluationContext(area, context);
  if (evaluationContext === undefined) return undefined;
  const sheetIds = areaSheetIds(area, evaluationContext);
  if (!sheetIds.length) return undefined;
  let total = 0;
  for (const sheet of sheetIds) {
    const bounds = areaBounds(area, evaluationContext, sheet);
    if (!bounds) return undefined;
    total +=
      (Math.abs(bounds.lastRow - bounds.firstRow) + 1) *
      (Math.abs(bounds.lastColumn - bounds.firstColumn) + 1);
  }
  return total;
}

function sparseAreaEntries(
  area: FormulaReferenceArea,
  context: FormulaEvaluationContext,
): Iterable<FormulaCellEntry> | undefined {
  const evaluationContext = areaEvaluationContext(area, context);
  if (evaluationContext === undefined) return undefined;
  if (evaluationContext.workbook.getCells === undefined) return undefined;
  return areaSheetIds(area, evaluationContext).flatMap((sheet) => {
    const bounds = areaBounds(area, evaluationContext, sheet);
    if (!bounds) return [];
    return [...(evaluationContext.workbook.getCells?.(sheet, bounds) ?? [])].filter(
      (cell) =>
        cell.row >= bounds.firstRow &&
        cell.row <= bounds.lastRow &&
        cell.column >= bounds.firstColumn &&
        cell.column <= bounds.lastColumn,
    );
  });
}

function unionReferenceAreas(ast: FormulaAst): FormulaReferenceArea[] | undefined {
  if (ast.kind === "group") return unionReferenceAreas(ast.operand);
  if (
    ast.kind === "binary" &&
    (ast.operator === "," || ast.operator === "~" || ast.operator === ":")
  ) {
    const left = unionReferenceAreas(ast.left);
    const right = unionReferenceAreas(ast.right);
    return left && right ? [...left, ...right] : undefined;
  }
  if (ast.kind !== "reference") return undefined;
  return ast.reference?.areas ?? undefined;
}

function evaluateCounting(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
  countNonBlank: boolean,
): FormulaValue {
  const matrixValues: FormulaValue[] = [];
  let sparseCount = 0;
  const isCounted = (value: FormulaScalar): boolean =>
    countNonBlank ? value !== null : typeof value === "number";
  for (const argument of ast.args) {
    if (argument.kind === "missing") continue;
    const areas = unionReferenceAreas(argument);
    const entries = areas?.map((area) => ({ area, cells: sparseAreaEntries(area, context) }));
    if (!entries || entries.some(({ cells }) => cells === undefined)) {
      matrixValues.push(evaluateNode(argument, context, visited));
      continue;
    }
    for (const { cells } of entries) {
      for (const cell of cells!) {
        const value =
          cell.error === true && isFormulaErrorToken(cell.value)
            ? formulaError(cell.value)
            : (cell.value ?? null);
        if (isCounted(value)) sparseCount += 1;
      }
    }
  }
  if (matrixValues.length === ast.args.length) {
    const values = matrixValues.flatMap((value) => (isMatrix(value) ? value.flat() : [value]));
    return values.filter(isCounted).length;
  }
  const localValues = matrixValues.flatMap((value) => (isMatrix(value) ? value.flat() : [value]));
  return sparseCount + localValues.filter(isCounted).length;
}

function areaEvaluationContext(
  area: FormulaReferenceArea,
  context: FormulaEvaluationContext,
): FormulaEvaluationContext | undefined {
  if (area.workbook === undefined) return context;
  const workbook = context.workbook.getWorkbook?.(area.workbook);
  return workbook === undefined ? undefined : { ...context, workbook };
}

function areaSheetIds(area: FormulaReferenceArea, context: FormulaEvaluationContext): number[] {
  if (area.fromSheet === undefined) return [context.sheet ?? 0];
  const from = context.workbook.getSheetId(area.fromSheet);
  if (from === undefined) return [];
  if (area.toSheet === undefined || area.toSheet === area.fromSheet) return [from];
  const to = context.workbook.getSheetId(area.toSheet);
  if (to === undefined) return [];
  const lower = Math.min(from, to);
  const upper = Math.max(from, to);
  const ids: number[] = [];
  for (let id = lower; id <= upper; id += 1) {
    if (context.workbook.getSheetName(id) !== undefined) ids.push(id);
  }
  return ids;
}

function areaBounds(
  area: FormulaReferenceArea,
  context: FormulaEvaluationContext,
  sheet = context.workbook.getSheetId(area.fromSheet ?? "") ?? context.sheet ?? 0,
): { firstRow: number; firstColumn: number; lastRow: number; lastColumn: number } | undefined {
  const startEdge = area.start;
  const endEdge = area.end ?? area.start;
  const start = edgeSpan(startEdge, context, sheet);
  const end = edgeSpan(endEdge, context, sheet);
  if (!start || !end) return undefined;
  return {
    firstRow: Math.min(start.firstRow, end.firstRow),
    firstColumn: Math.min(start.firstColumn, end.firstColumn),
    lastRow: Math.max(start.lastRow, end.lastRow),
    lastColumn: Math.max(start.lastColumn, end.lastColumn),
  };
}

function edgeSpan(
  edge: FormulaCellCoordinate | FormulaReferenceArea["start"],
  context: FormulaEvaluationContext,
  sheet: number,
): { firstRow: number; firstColumn: number; lastRow: number; lastColumn: number } | undefined {
  if (isWholeColumn(edge)) {
    return {
      firstRow: 1,
      firstColumn: edge.column,
      lastRow: context.workbook.getRowCount(sheet),
      lastColumn: edge.column,
    };
  }
  if (isWholeRow(edge)) {
    return {
      firstRow: edge.row,
      firstColumn: 1,
      lastRow: edge.row,
      lastColumn: context.workbook.getColumnCount(sheet),
    };
  }
  if (
    edge.row < 1 ||
    edge.column < 1 ||
    edge.row > context.workbook.getRowCount(sheet) ||
    edge.column > context.workbook.getColumnCount(sheet)
  )
    return undefined;
  return {
    firstRow: edge.row,
    firstColumn: edge.column,
    lastRow: edge.row,
    lastColumn: edge.column,
  };
}

function unionMatrix(left: FormulaValue, right: FormulaValue): FormulaValue {
  const leftMatrix = isMatrix(left) ? left : [[left]];
  const rightMatrix = isMatrix(right) ? right : [[right]];
  const cellCount = (matrix: FormulaMatrix) => matrix.reduce((total, row) => total + row.length, 0);
  assertFormulaCellCount(cellCount(leftMatrix) + cellCount(rightMatrix), 1);
  return [...flattenMatrix(leftMatrix), ...flattenMatrix(rightMatrix)];
}

function flattenMatrix(matrix: FormulaMatrix): FormulaMatrix {
  return matrix.flatMap((row) => row.map((value) => [value]));
}

function readCells(
  sheet: number,
  firstRowArgument: number,
  firstColumnArgument: number,
  lastRowArgument: number,
  lastColumnArgument: number,
  context: FormulaEvaluationContext,
): FormulaValue {
  const firstRow = Math.min(firstRowArgument, lastRowArgument);
  const lastRow = Math.max(firstRowArgument, lastRowArgument);
  const firstColumn = Math.min(firstColumnArgument, lastColumnArgument);
  const lastColumn = Math.max(firstColumnArgument, lastColumnArgument);
  if (lastRow > maximumRows || lastColumn > maximumColumns) return formulaError("#REF!");
  assertFormulaCellCount(lastRow - firstRow + 1, lastColumn - firstColumn + 1);
  const values: FormulaScalar[][] = [];
  for (let row = firstRow; row <= lastRow; row += 1) {
    const matrixRow: FormulaScalar[] = [];
    for (let column = firstColumn; column <= lastColumn; column += 1) {
      const cell = context.workbook.getCell(sheet, row, column);
      matrixRow.push(
        cell?.error === true && isFormulaErrorToken(cell.value)
          ? formulaError(cell.value)
          : (cell?.value ?? null),
      );
    }
    values.push(matrixRow);
  }
  return values.length === 1 && values[0]!.length === 1 ? values[0]![0]! : values;
}

function evaluateReference(raw: string, context: FormulaEvaluationContext): FormulaValue {
  const normalizedRaw = raw.replace(/^=/u, "");
  const sheetName = normalizedRaw.includes("!")
    ? normalizedRaw.split("!")[0]!.replace(/^'/u, "").replace(/'$/u, "")
    : undefined;
  const sheet = sheetName ? context.workbook.getSheetId(sheetName) : (context.sheet ?? 0);
  if (sheet === undefined) return "#REF!";
  const reference = normalizedRaw.includes("!")
    ? normalizedRaw.split("!")[1]!.replace(/#$/u, "")
    : normalizedRaw;
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/u.exec(reference);
  if (!match) return "#VALUE!";
  return readCells(
    sheet,
    Number(match[2]!),
    columnNumber(match[1]!),
    match[4] ? Number(match[4]) : Number(match[2]!),
    match[3] ? columnNumber(match[3]) : columnNumber(match[1]!),
    context,
  );
}

function unaryScalars(
  value: FormulaValue,
  operation: (value: FormulaScalar) => FormulaScalar,
): FormulaValue {
  const propagatedError = firstError([value]);
  if (propagatedError !== undefined) return propagatedError;
  if (!isMatrix(value)) return normalizeScalar(operation(value));
  return value.map((row) => row.map((scalar) => normalizeScalar(operation(scalar))));
}

function columnNumber(column: string): number {
  let value = 0;
  for (const character of column.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64;
  return value;
}

function evaluateBinary(
  operator: FormulaBinaryOperator,
  leftAst: FormulaAst,
  rightAst: FormulaAst,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  const left = evaluateNode(leftAst, context, visited);
  const right = evaluateNode(rightAst, context, visited);
  if (operator === "&")
    return binaryScalars(left, right, (first, second) => toText(first) + toText(second));
  switch (operator) {
    case "+":
      return binaryScalars(left, right, (first, second) => toNumber(first) + toNumber(second));
    case "-":
      return binaryScalars(left, right, (first, second) => toNumber(first) - toNumber(second));
    case "*":
      return binaryScalars(left, right, (first, second) => toNumber(first) * toNumber(second));
    case "/":
      return binaryScalars(left, right, (first, second) => {
        const divisor = toNumber(second);
        if (divisor === 0) return "#DIV/0!";
        return toNumber(first) / divisor;
      });
    case "^":
      return binaryScalars(left, right, (first, second) => toNumber(first) ** toNumber(second));
    case "=":
      return binaryScalars(left, right, (first, second) => compareScalars(first, second) === 0);
    case "<>":
      return binaryScalars(left, right, (first, second) => compareScalars(first, second) !== 0);
    case "<":
      return binaryScalars(left, right, (first, second) => compareScalars(first, second) < 0);
    case "<=":
      return binaryScalars(left, right, (first, second) => compareScalars(first, second) <= 0);
    case ">":
      return binaryScalars(left, right, (first, second) => compareScalars(first, second) > 0);
    case ">=":
      return binaryScalars(left, right, (first, second) => compareScalars(first, second) >= 0);
    case " ":
      return referenceIntersection(leftAst, rightAst, context);
    case "!":
      return referenceIntersection(leftAst, rightAst, context);
    case ":":
      return unionMatrix(left, right);
    case ",":
      return unionMatrix(left, right);
    case "~":
      return unionMatrix(left, right);
    default:
      return "#VALUE!";
  }
}

function binaryScalars(
  left: FormulaValue,
  right: FormulaValue,
  operation: (left: FormulaScalar, right: FormulaScalar) => FormulaScalar,
): FormulaValue {
  const normalizedOperation = (first: FormulaScalar, second: FormulaScalar): FormulaScalar =>
    normalizeScalar(operation(first, second));
  const propagatedError = firstError([left, right]);
  if (propagatedError !== undefined) return propagatedError;
  if (!isMatrix(left) && !isMatrix(right)) return normalizedOperation(left, right);
  const leftMatrix = isMatrix(left) ? left : [[left]];
  const rightMatrix = isMatrix(right) ? right : [[right]];
  const rows = Math.max(leftMatrix.length, rightMatrix.length);
  const columns = Math.max(
    Math.max(0, ...leftMatrix.map((row) => row.length)),
    Math.max(0, ...rightMatrix.map((row) => row.length)),
  );
  assertFormulaCellCount(rows, columns);
  const output: FormulaMatrix = [];
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const outputRow: FormulaScalar[] = [];
    for (let columnIndex = 0; columnIndex < columns; columnIndex += 1) {
      const first = broadcastCell(leftMatrix, rowIndex, columnIndex);
      const second = broadcastCell(rightMatrix, rowIndex, columnIndex);
      outputRow.push(
        first === "#N/A" || second === "#N/A" ? "#N/A" : normalizedOperation(first, second),
      );
    }
    output.push(outputRow);
  }
  return output;
}

function broadcastCell(matrix: FormulaMatrix, row: number, column: number): FormulaScalar {
  const sourceRow = matrix.length === 1 ? 0 : row;
  const sourceRowValues = matrix[sourceRow];
  if (sourceRowValues === undefined) return "#N/A";
  const sourceColumn = sourceRowValues.length === 1 ? 0 : column;
  return sourceRowValues[sourceColumn] ?? "#N/A";
}

function referenceSheetId(
  ast: Extract<FormulaAst, { kind: "reference" }>,
  context: FormulaEvaluationContext,
): number | undefined {
  const area = ast.reference?.areas?.[0];
  if (area?.fromSheet !== undefined) return context.workbook.getSheetId(area.fromSheet);
  const sheetName = ast.reference?.sheets?.names[0];
  if (sheetName === undefined) return context.sheet ?? 0;
  return context.workbook.getSheetId(sheetName);
}

function referenceIntersection(
  leftAst: FormulaAst,
  rightAst: FormulaAst,
  context: FormulaEvaluationContext,
): FormulaValue {
  if (leftAst.kind !== "reference" || rightAst.kind !== "reference") return "#NULL!";
  const leftSheet = referenceSheetId(leftAst, context);
  const rightSheet = referenceSheetId(rightAst, context);
  const leftBounds = referenceBounds(leftAst, context);
  const rightBounds = referenceBounds(rightAst, context);
  if (
    leftSheet === undefined ||
    rightSheet === undefined ||
    leftSheet !== rightSheet ||
    !leftBounds ||
    !rightBounds
  )
    return "#NULL!";
  const firstRow = Math.max(leftBounds.firstRow, rightBounds.firstRow);
  const firstColumn = Math.max(leftBounds.firstColumn, rightBounds.firstColumn);
  const lastRow = Math.min(leftBounds.lastRow, rightBounds.lastRow);
  const lastColumn = Math.min(leftBounds.lastColumn, rightBounds.lastColumn);
  if (firstRow > lastRow || firstColumn > lastColumn) return "#NULL!";
  return readCells(leftSheet, firstRow, firstColumn, lastRow, lastColumn, context);
}

function referenceAreaCount(ast: FormulaAst | undefined): FormulaValue {
  if (ast?.kind === "group") return referenceAreaCount(ast.operand);
  if (ast?.kind === "binary" && ast.operator === ",") {
    const left = referenceAreaCount(ast.left);
    const right = referenceAreaCount(ast.right);
    return typeof left === "number" && typeof right === "number" ? left + right : "#VALUE!";
  }
  if (ast?.kind !== "reference") return "#VALUE!";
  return ast.reference?.areas?.length ?? 1;
}

function referenceSheetCount(
  ast: FormulaAst | undefined,
  context: FormulaEvaluationContext,
): FormulaValue {
  if (ast?.kind === "group") return referenceSheetCount(ast.operand, context);
  if (ast !== undefined && ast.kind !== "reference") return "#VALUE!";
  const reference = ast?.reference;
  const area = reference?.areas?.[0];
  if (area?.fromSheet !== undefined) {
    const first = context.workbook.getSheetId(area.fromSheet);
    const last = context.workbook.getSheetId(area.toSheet ?? area.fromSheet);
    return first !== undefined && last !== undefined ? Math.abs(last - first) + 1 : "#REF!";
  }
  const names = reference?.sheets?.names;
  if (!names?.length) return "#VALUE!";
  if (!reference?.sheets?.range || names.length === 1) return names.length;
  const first = context.workbook.getSheetId(names[0]!);
  const last = context.workbook.getSheetId(names.at(-1)!);
  return first !== undefined && last !== undefined ? Math.abs(last - first) + 1 : "#REF!";
}

function sheetTotal(workbook: FormulaWorkbookContext): number {
  let count = 0;
  while (workbook.getSheetName(count) !== undefined) count += 1;
  return count;
}

function aggregateReferenceEntries(
  ast: Extract<FormulaAst, { kind: "reference" }>,
  context: FormulaEvaluationContext,
  ignoreHidden: boolean,
  ignoreNestedAggregates: boolean,
): FormulaScalar[] {
  const areas = ast.reference?.areas ?? [];
  if (!areas.length && ast.reference?.table) {
    const resolved = resolveTableSelection(
      ast.reference.table,
      context,
      ast.reference.implicitIntersection,
    );
    if (resolved.error !== undefined) return [resolved.error];
    const selection = resolved.selection!;
    const output: FormulaScalar[] = [];
    for (let row = selection.firstRow; row <= selection.lastRow; row += 1) {
      if (ignoreHidden && context.workbook.isRowHidden?.(selection.sheetId, row)) continue;
      for (const [from, to] of selection.columns) {
        for (let column = from; column <= to; column += 1) {
          const cell = context.workbook.getCell(selection.sheetId, row, column);
          output.push(
            ignoreNestedAggregates &&
              cell?.formula !== undefined &&
              isNestedAggregateFormula(cell.formula)
              ? null
              : cell?.error === true && isFormulaErrorToken(cell.value)
                ? formulaError(cell.value)
                : (cell?.value ?? null),
          );
        }
      }
    }
    return output;
  }
  return areas.flatMap((area) => {
    const areaContext = areaEvaluationContext(area, context);
    if (areaContext === undefined) return ["#REF!"];
    const cellCount = areaCellCount(area, context);
    if (cellCount === undefined) return ["#REF!"];
    assertFormulaCellCount(cellCount, 1);
    const output: FormulaScalar[] = [];
    for (const sheet of areaSheetIds(area, areaContext)) {
      const bounds = areaBounds(area, areaContext, sheet);
      if (!bounds) return ["#REF!"];
      for (let row = bounds.firstRow; row <= bounds.lastRow; row += 1) {
        if (ignoreHidden && areaContext.workbook.isRowHidden?.(sheet, row)) continue;
        for (let column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
          const cell = areaContext.workbook.getCell(sheet, row, column);
          output.push(
            ignoreNestedAggregates === true &&
              cell?.formula !== undefined &&
              isNestedAggregateFormula(cell.formula)
              ? null
              : cell?.error === true && isFormulaErrorToken(cell.value)
                ? formulaError(cell.value)
                : (cell?.value ?? null),
          );
        }
      }
    }
    return output;
  });
}

function isNestedAggregateFormula(formula: string): boolean {
  return /^=(?:_XLFN\.)?(?:SUBTOTAL|AGGREGATE)\(/iu.test(formula.trim());
}

function evaluateSubtotal(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (ast.args.length < 2) return "#VALUE!";
  const requested = toNumber(evaluateNode(ast.args[0]!, context, visited));
  const aggregateType = requested > 100 ? requested - 100 : requested;
  if (!Number.isInteger(requested) || requested < 1 || requested > 111 || aggregateType < 1)
    return "#VALUE!";
  const ignoreHidden = requested > 100;
  const entries = ast.args
    .slice(1)
    .flatMap((argument) => aggregateEntries(argument, context, visited, ignoreHidden, true));
  const mappedType =
    aggregateType === 3 ? 3 : aggregateType === 1 ? 1 : aggregateType === 2 ? 2 : aggregateType;
  return getFunction("AGGREGATE")!([mappedType, 4, [entries]], context);
}

function evaluateGrouping(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  const aggregatorIndex = ast.name.toUpperCase() === "GROUPBY" ? 2 : 3;
  const aggregator = ast.args[aggregatorIndex];
  const aggregatorName =
    aggregator?.kind === "name"
      ? aggregator.name
      : aggregator?.kind === "reference" &&
          aggregator.reference?.table === undefined &&
          aggregator.reference?.start === undefined &&
          aggregator.reference?.areas === undefined
        ? aggregator.raw
        : aggregator?.kind === "call" && aggregator.args.length === 0
          ? aggregator.name
          : undefined;
  const args = ast.args.map((argument, index) =>
    index === aggregatorIndex && aggregatorName !== undefined
      ? aggregatorName
      : evaluateNode(argument, context, visited),
  );
  return getFunction(ast.name)!(args, context);
}

function aggregateEntries(
  ast: FormulaAst,
  context: FormulaEvaluationContext,
  visited: Set<string>,
  ignoreHidden: boolean,
  ignoreNestedAggregates: boolean,
): FormulaScalar[] {
  if (ast.kind === "group") {
    return aggregateEntries(ast.operand, context, visited, ignoreHidden, ignoreNestedAggregates);
  }
  if (
    ast.kind === "call" &&
    ignoreNestedAggregates &&
    ["AGGREGATE", "SUBTOTAL"].includes(ast.name.toUpperCase())
  )
    return [];
  if (ast.kind === "binary" && (ast.operator === "," || ast.operator === "~")) {
    return [
      ...aggregateEntries(ast.left, context, visited, ignoreHidden, ignoreNestedAggregates),
      ...aggregateEntries(ast.right, context, visited, ignoreHidden, ignoreNestedAggregates),
    ];
  }
  if (ast.kind === "reference" && ast.reference) {
    return aggregateReferenceEntries(ast, context, ignoreHidden, ignoreNestedAggregates);
  }
  return flatten(evaluateNode(ast, context, visited));
}

function evaluateAggregate(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
): FormulaValue {
  if (ast.args.length < 3) return "#VALUE!";
  const type = toNumber(evaluateNode(ast.args[0]!, context, visited));
  const option = toNumber(evaluateNode(ast.args[1]!, context, visited));
  if (
    !Number.isInteger(type) ||
    type < 1 ||
    type > 19 ||
    !Number.isInteger(option) ||
    option < 0 ||
    option > 7
  )
    return "#VALUE!";
  if (type > 13 && ast.args.length !== 4) return "#VALUE!";
  const ignoreHidden = option === 1 || option === 3 || option === 5 || option === 7;
  const ignoreNestedAggregates = option >= 0 && option <= 3;
  const aggregateArguments = type <= 13 ? ast.args.slice(2) : [ast.args[2]!];
  const entries = aggregateArguments.flatMap((argument) =>
    aggregateEntries(argument, context, visited, ignoreHidden, ignoreNestedAggregates),
  );
  const values: FormulaMatrix = [entries];
  const extra =
    type > 13 ? ast.args.slice(3).map((argument) => evaluateNode(argument, context, visited)) : [];
  return getFunction("AGGREGATE")!([type, option, values, ...extra], context);
}

export function evaluateFormula(
  formula: string,
  workbook: FormulaWorkbookContext,
  options: FormulaEvaluationOptions = {},
): FormulaEvaluationResult {
  const ast = parseFormula(formula, { dialect: options.dialect ?? "excel" });
  const value = evaluateFormulaAst(ast, workbook, options);
  const dependencies = collectDependencies(ast).map((dependency) =>
    formatFormulaDependencyTarget(dependency.target),
  );
  return { value, dependencies: [...new Set(dependencies)] };
}

function referenceSnapshotError(
  ast: FormulaAst,
  context: FormulaEvaluationContext,
): FormulaErrorValue | null | undefined {
  if (ast.kind !== "reference") return undefined;
  const area = ast.reference?.areas?.[0];
  if (!area) return undefined;
  const areaContext = areaEvaluationContext(area, context);
  const sheet = areaSheetIds(area, areaContext ?? context)[0];
  const bounds = areaBounds(area, areaContext ?? context);
  if (!areaContext || sheet === undefined || !bounds) return undefined;
  if (bounds.firstRow !== bounds.lastRow || bounds.firstColumn !== bounds.lastColumn) return null;
  const cell = areaContext.workbook.getCell(sheet, bounds.firstRow, bounds.firstColumn);
  return cell?.error === true && isFormulaErrorToken(cell.value) ? cell.value : null;
}

function evaluateInformationError(
  ast: Extract<FormulaAst, { kind: "call" }>,
  context: FormulaEvaluationContext,
  visited: Set<string>,
  matches: (value: FormulaErrorValue) => boolean,
): boolean {
  const argument = ast.args[0];
  if (!argument) return false;
  const cellError = referenceSnapshotError(argument, context);
  if (cellError !== undefined) return cellError !== null && matches(cellError);
  const value = evaluateNode(argument, context, visited);
  const scalar = isMatrix(value) ? value.flat()[0] : value;
  return isFormulaErrorValue(scalar) && matches(scalar.value);
}

const astFunctionEvaluators: readonly (readonly [
  string,
  NonNullable<RegisteredFunction["evaluateAst"]>,
])[] = [
  ["IF", evaluateIf],
  ["IFERROR", (ast, context, visited) => evaluateIfError(ast, context, visited, false)],
  ["IFNA", (ast, context, visited) => evaluateIfError(ast, context, visited, true)],
  ["CHOOSE", evaluateChoose],
  ["IFS", evaluateIfs],
  ["SWITCH", evaluateSwitch],
  ["AND", (ast, context, visited) => evaluateAndOr(ast, context, visited, false)],
  ["OR", (ast, context, visited) => evaluateAndOr(ast, context, visited, true)],
  [
    "ISERR",
    (ast, context, visited) =>
      evaluateInformationError(ast, context, visited, (value) => value !== "#N/A"),
  ],
  [
    "ISERROR",
    (ast, context, visited) => evaluateInformationError(ast, context, visited, () => true),
  ],
  [
    "ISNA",
    (ast, context, visited) =>
      evaluateInformationError(ast, context, visited, (value) => value === "#N/A"),
  ],
  ["LAMBDA", createLambda],
  ["LET", evaluateLet],
  ["ISOMITTED", (ast) => ast.args.length === 1 && ast.args[0]?.kind === "missing"],
  ["BYROW", (ast, context, visited) => applyByDimension(ast, context, visited, false)],
  ["BYCOL", (ast, context, visited) => applyByDimension(ast, context, visited, true)],
  ["MAP", applyMap],
  ["MAKEARRAY", applyMakeArray],
  ["REDUCE", (ast, context, visited) => applyReduction(ast, context, visited, false)],
  ["SCAN", (ast, context, visited) => applyReduction(ast, context, visited, true)],
  ["ROW", (ast, context) => referencePosition(ast, context, "ROW")],
  ["COLUMN", (ast, context) => referencePosition(ast, context, "COLUMN")],
  ["COUNT", (ast, context, visited) => evaluateCounting(ast, context, visited, false)],
  ["COUNTA", (ast, context, visited) => evaluateCounting(ast, context, visited, true)],
  ["INDIRECT", evaluateIndirect],
  ["OFFSET", evaluateOffset],
  ["FORMULATEXT", (ast, context) => cellFormulaText(ast, context)],
  ["FORMULA", (ast, context) => cellFormulaText(ast, context)],
  ["AREAS", (ast) => referenceAreaCount(ast.args[0])],
  ["AGGREGATE", evaluateAggregate],
  ["SUBTOTAL", evaluateSubtotal],
  ["GROUPBY", evaluateGrouping],
  ["PIVOTBY", evaluateGrouping],
  [
    "SHEETS",
    (ast, context) =>
      ast.args[0] === undefined || ast.args[0].kind === "missing"
        ? sheetTotal(context.workbook)
        : referenceSheetCount(ast.args[0], context),
  ],
  [
    "SQL.REQUEST",
    (ast, context, visited) => {
      const args = ast.args.map((argument, index) => {
        if (argument.kind === "missing") return undefined;
        if (index === 1 && argument.kind === "reference") return argument.raw;
        return evaluateNode(argument, context, visited);
      });
      return sqlRequest(
        args,
        context,
        ast.args[1]?.kind === "reference" ? ast.args[1].raw : undefined,
      );
    },
  ],
  [
    "ISREF",
    (ast, context, visited) => {
      if (ast.args.length !== 1 || ast.args[0]?.kind === "missing") return "#VALUE!";
      const value = ast.args[0] === undefined ? null : evaluateNode(ast.args[0], context, visited);
      return firstError([value]) !== undefined ? true : ast.args[0]?.kind === "reference";
    },
  ],
];

for (const [name, evaluateAst] of astFunctionEvaluators) setFunctionAstEvaluator(name, evaluateAst);
