import type { FormulaAst } from "../ast";
import type { FormulaEvaluationContext } from "../context";
import type { FormulaMatrix, FormulaValue } from "../value";
import { isMatrix } from "../value";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { assertFormulaCellCount, define, firstError, scalarNumber, scalarText } from "./utils";

type Arguments = Parameters<FormulaFunctionImplementation>[0];

export function isFormulaLookupContext(context: FormulaEvaluationContext): boolean {
  return context.workbook.resolveLookup !== undefined;
}

function lookupContext(context: FormulaEvaluationContext): FormulaEvaluationContext | undefined {
  return isFormulaLookupContext(context) ? context : undefined;
}

function applyLookupFunction(
  name: string,
  args: Arguments,
  context: FormulaEvaluationContext,
  nodes?: readonly FormulaAst[],
): FormulaValue {
  const host = lookupContext(context)?.workbook;
  if (!host?.resolveLookup) return "#REF!";
  return (
    host.resolveLookup({
      name,
      values: args,
      ...(nodes === undefined ? {} : { nodes }),
      caller: { sheet: context.sheet, row: context.row, column: context.column },
    }) ?? "#REF!"
  );
}

function sum(values: readonly FormulaValue[]): number {
  return values
    .flatMap((value) => (isMatrix(value) ? value.flat() : [value]))
    .reduce<number>((total, value) => total + scalarNumber(value), 0);
}

function percentOf(args: Arguments): FormulaValue {
  const denominatorSource = args.length > 1 ? args.slice(1) : args;
  const denominator = sum(denominatorSource);
  if (denominator === 0) return "#DIV/0!";
  const numerator = args[0]!;
  if (args.length === 1 && isMatrix(numerator))
    return numerator.map((row) => row.map((value) => scalarNumber(value) / denominator));
  return sum([numerator]) / denominator;
}

function identity(size: number): FormulaMatrix {
  return Array.from({ length: size }, (_, row) =>
    Array.from({ length: size }, (_, column) => (row === column ? 1 : 0)),
  );
}

function munit(value: FormulaValue): FormulaValue {
  const size = Math.trunc(scalarNumber(value));
  if (size <= 0 || size !== scalarNumber(value)) return "#VALUE!";
  assertFormulaCellCount(size, size);
  return identity(size);
}

function hyperlink(location: FormulaValue, display?: FormulaValue): FormulaValue {
  const target = scalarText(location).trim();
  if (target === "") return "#VALUE!";
  return display === undefined ? target : display;
}

export const lookupExtraFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  [
    "CELL",
    define({ minArgs: 1, maxArgs: 2, volatile: true }, (args, context) =>
      applyLookupFunction("CELL", args, context),
    ),
  ],
  [
    "COLUMN",
    define({ minArgs: 0, maxArgs: 1 }, (args, context) => {
      if (args.length === 0) return context.column === undefined ? "#VALUE!" : context.column;
      return applyLookupFunction("COLUMN", args, context);
    }),
  ],
  [
    "FORMULA",
    define({ minArgs: 1, maxArgs: 1 }, (args, context) =>
      applyLookupFunction("FORMULA", args, context),
    ),
  ],
  [
    "FORMULATEXT",
    define({ minArgs: 1, maxArgs: 1 }, (args, context) =>
      applyLookupFunction("FORMULATEXT", args, context),
    ),
  ],
  [
    "HYPERLINK",
    define({ minArgs: 1, maxArgs: 2 }, (args) => {
      const error = firstError(args);
      if (error !== undefined) return error;
      return hyperlink(args[0]!, args[1]);
    }),
  ],
  [
    "IMPORTRANGE",
    define({ minArgs: 2, maxArgs: 2 }, (args, context) =>
      applyLookupFunction("IMPORTRANGE", args, context),
    ),
  ],
  [
    "INDIRECT",
    define({ minArgs: 1, maxArgs: 2, volatile: true }, (args, context) =>
      applyLookupFunction("INDIRECT", args, context),
    ),
  ],
  [
    "INFO",
    define({ minArgs: 1, maxArgs: 1 }, (args, context) =>
      applyLookupFunction("INFO", args, context),
    ),
  ],
  [
    "ISFORMULA",
    define({ minArgs: 1, maxArgs: 1 }, (args, context) =>
      applyLookupFunction("ISFORMULA", args, context),
    ),
  ],
  [
    "ISREF",
    define({ minArgs: 1, maxArgs: 1 }, (args, context) =>
      applyLookupFunction("ISREF", args, context),
    ),
  ],
  ["MUNIT", define({ minArgs: 1, maxArgs: 1 }, ([value]) => munit(value))],
  [
    "OFFSET",
    define({ minArgs: 3, maxArgs: 5, volatile: true }, (args, context) =>
      applyLookupFunction("OFFSET", args, context),
    ),
  ],
  [
    "PERCENTOF",
    define({ minArgs: 0, maxArgs: 255 }, (args) => {
      const error = firstError(args);
      return error === undefined ? percentOf(args) : error;
    }),
  ],
  [
    "ROW",
    define({ minArgs: 0, maxArgs: 1 }, (args, context) => {
      if (args.length === 0) return context.row === undefined ? "#VALUE!" : context.row;
      return applyLookupFunction("ROW", args, context);
    }),
  ],
  [
    "STYLE",
    define({ minArgs: 1, maxArgs: 3 }, (args, context) =>
      applyLookupFunction("STYLE", args, context),
    ),
  ],
  [
    "SUBTOTAL",
    define({ minArgs: 2, maxArgs: Number.POSITIVE_INFINITY }, (args, context) =>
      applyLookupFunction("SUBTOTAL", args, context),
    ),
  ],
]);
