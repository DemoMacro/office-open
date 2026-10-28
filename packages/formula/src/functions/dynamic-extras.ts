import type { FormulaAst } from "../ast";
import type { FormulaEvaluationContext } from "../context";
import { FormulaValueError } from "../errors";
import type { FormulaMatrix, FormulaScalar, FormulaValue } from "../value";
import { isMatrix } from "../value";
import {
  fitLinearRegression,
  predictLinearValues,
  predictionRows,
  regressionObservations,
} from "./regression";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import {
  assertFormulaCellCount,
  define,
  firstError,
  scalarBoolean,
  scalarNumber,
  scalarText,
  values,
} from "./utils";

type Arguments = Parameters<FormulaFunctionImplementation>[0];

export function isFormulaDynamicContext(context: FormulaEvaluationContext): boolean {
  return context.workbook.resolveDynamic !== undefined;
}

function dynamicContext(context: FormulaEvaluationContext): FormulaEvaluationContext | undefined {
  return isFormulaDynamicContext(context) ? context : undefined;
}

function applyDynamicFunction(
  name: string,
  args: Arguments,
  context: FormulaEvaluationContext,
  nodes?: readonly FormulaAst[],
): FormulaValue {
  const host = dynamicContext(context)?.workbook;
  if (!host?.resolveDynamic) return "#CALC!";
  return (
    host.resolveDynamic({
      name,
      values: args,
      ...(nodes === undefined ? {} : { nodes }),
      caller: { sheet: context.sheet, row: context.row, column: context.column },
    }) ?? "#CALC!"
  );
}

function matrix(value: FormulaValue): FormulaMatrix {
  return isMatrix(value) ? value : [[value]];
}

function width(input: FormulaMatrix): number {
  return Math.max(0, ...input.map((row) => row.length));
}

function transpose(input: FormulaMatrix): FormulaMatrix {
  return Array.from({ length: width(input) }, (_, column) =>
    Array.from({ length: input.length }, (_, row) => input[row]?.[column] ?? null),
  );
}

function scalarValue(value: FormulaValue): FormulaScalar {
  return isMatrix(value) ? (value.flat()[0] ?? null) : value;
}

function requireInteger(value: number): number {
  const integer = Math.trunc(value);
  if (integer !== value) throw new FormulaValueError("#VALUE!");
  return integer;
}

function selectedRows(source: FormulaMatrix, indexes: readonly number[]): FormulaMatrix {
  return indexes.map((index) => {
    const effective = index < 0 ? source.length + index : index;
    if (effective < 0 || effective >= source.length) throw new FormulaValueError("#VALUE!");
    return [...(source[effective] ?? [])];
  });
}

function selectedColumns(source: FormulaMatrix, indexes: readonly number[]): FormulaMatrix {
  const total = width(source);
  return indexes.map((index) => {
    const effective = index < 0 ? total + index : index;
    if (effective < 0 || effective >= total) throw new FormulaValueError("#VALUE!");
    return source.map((row) => row[effective] ?? null);
  });
}

function argumentIndexes(args: readonly FormulaValue[]): number[] {
  return args.flatMap((value) => values(value)).map((value) => requireInteger(scalarNumber(value)));
}

function growth(args: Arguments): FormulaValue {
  const knownY = args[0]!;
  const knownX: FormulaValue = args[1] ?? [
    Array.from({ length: isMatrix(knownY) ? knownY.flat().length : 1 }, (_, index) => index + 1),
  ];
  const useConstant = args[3] === undefined ? true : scalarBoolean(args[3]);
  const observations = regressionObservations(knownY, knownX);
  if (observations === undefined || observations.response.length < 2)
    throw new FormulaValueError("#REF!");
  if (observations.response.some((value) => value <= 0)) throw new FormulaValueError("#NUM!");
  const logarithms = observations.response.map((value) => Math.log(value));
  const statistics = fitLinearRegression(logarithms, observations.predictors, useConstant);
  const newX: FormulaValue = args[2] ?? knownX;
  const predictions = predictionRows(newX, observations.predictors[0]?.length ?? 0);
  if (predictions === undefined) throw new FormulaValueError("#REF!");
  return predictLinearValues(statistics, predictions, true);
}

export const dynamicExtraFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  [
    "ANCHORARRAY",
    define({ minArgs: 1, maxArgs: 1 }, (args, context) =>
      applyDynamicFunction("ANCHORARRAY", args, context),
    ),
  ],
  [
    "BYCOL",
    define({ minArgs: 2, maxArgs: 2 }, (args, context) =>
      applyDynamicFunction("BYCOL", args, context),
    ),
  ],
  [
    "BYROW",
    define({ minArgs: 2, maxArgs: 2 }, (args, context) =>
      applyDynamicFunction("BYROW", args, context),
    ),
  ],
  [
    "CHOOSECOLS",
    define({ minArgs: 2, maxArgs: Number.POSITIVE_INFINITY }, (args) => {
      const error = firstError(args);
      if (error !== undefined) return error;
      return selectedColumns(matrix(args[0]!), argumentIndexes(args.slice(1)));
    }),
  ],
  [
    "CHOOSEROWS",
    define({ minArgs: 2, maxArgs: Number.POSITIVE_INFINITY }, (args) => {
      const error = firstError(args);
      if (error !== undefined) return error;
      return selectedRows(matrix(args[0]!), argumentIndexes(args.slice(1)));
    }),
  ],
  [
    "CURRENT",
    define({ minArgs: 0, maxArgs: 0 }, (_args, context) =>
      applyDynamicFunction("CURRENT", [], context),
    ),
  ],
  [
    "GROWTH",
    define({ minArgs: 1, maxArgs: Number.POSITIVE_INFINITY }, (args) => {
      const error = firstError(args);
      return error === undefined ? growth(args) : error;
    }),
  ],
  [
    "IFS",
    define({ minArgs: 2, maxArgs: Number.POSITIVE_INFINITY }, (args) => {
      const error = firstError(args);
      if (error !== undefined) return error;
      for (let index = 0; index + 1 < args.length; index += 2)
        if (scalarBoolean(args[index]!)) return args[index + 1] ?? null;
      return "#N/A";
    }),
  ],
  [
    "ISOMITTED",
    define({ minArgs: 1, maxArgs: 1 }, (args, context) =>
      applyDynamicFunction("ISOMITTED", args, context),
    ),
  ],
  [
    "LAMBDA",
    define({ minArgs: 1, maxArgs: Number.POSITIVE_INFINITY }, (args, context) =>
      applyDynamicFunction("LAMBDA", args, context),
    ),
  ],
  [
    "LET",
    define({ minArgs: 3, maxArgs: Number.POSITIVE_INFINITY }, (args, context) =>
      applyDynamicFunction("LET", args, context),
    ),
  ],
  [
    "MAKEARRAY",
    define({ minArgs: 3, maxArgs: 3 }, (args, context) =>
      applyDynamicFunction("MAKEARRAY", args, context),
    ),
  ],
  [
    "MAP",
    define({ minArgs: 0, maxArgs: 255 }, (args, context) =>
      applyDynamicFunction("MAP", args, context),
    ),
  ],
  [
    "MULTIPLEOPERATIONS",
    define({ minArgs: 4, maxArgs: 5 }, (args, context) =>
      applyDynamicFunction("MULTIPLEOPERATIONS", args, context),
    ),
  ],
  [
    "MULTIPLE.OPERATIONS",
    define({ minArgs: 4, maxArgs: 5 }, (args, context) =>
      applyDynamicFunction("MULTIPLEOPERATIONS", args, context),
    ),
  ],
  [
    "RAND.NV",
    define({ minArgs: 0, maxArgs: 0 }, (_args, context) =>
      applyDynamicFunction("RAND.NV", [], context),
    ),
  ],
  [
    "RANDBETWEEN.NV",
    define({ minArgs: 2, maxArgs: 2 }, (args, context) =>
      applyDynamicFunction("RANDBETWEEN.NV", args, context),
    ),
  ],
  [
    "REDUCE",
    define({ minArgs: 2, maxArgs: 3 }, (args, context) =>
      applyDynamicFunction("REDUCE", args, context),
    ),
  ],
  [
    "SCAN",
    define({ minArgs: 2, maxArgs: 3 }, (args, context) =>
      applyDynamicFunction("SCAN", args, context),
    ),
  ],
  [
    "SWITCH",
    define({ minArgs: 3, maxArgs: Number.POSITIVE_INFINITY }, (args) => {
      const error = firstError(args);
      if (error !== undefined) return error;
      const expression = scalarText(args[0]!).toUpperCase();
      for (let index = 1; index + 1 < args.length; index += 2)
        if (scalarText(args[index]!).toUpperCase() === expression) return args[index + 1] ?? null;
      return args.length % 2 === 0 ? (args[args.length - 1] ?? null) : "#N/A";
    }),
  ],
  ["TRANSPOSE", define({ minArgs: 1, maxArgs: 1 }, ([value]) => transpose(matrix(value)))],
  [
    "VALUES",
    define({ minArgs: 0, maxArgs: 255 }, (args): FormulaValue => {
      if (args.length === 0) return null;
      const groups = args.map((value) => values(value));
      let total = 0;
      for (const group of groups) {
        total += group.length;
        assertFormulaCellCount(Math.max(total, 1), 1);
      }
      const flattened = groups.flat();
      if (flattened.length === 1) return flattened[0] ?? null;
      const column: FormulaMatrix = flattened.map((value) => [scalarValue(value)]);
      return column;
    }),
  ],
]);
