import type { FormulaEvaluationContext } from "../context";
import { FormulaValueError } from "../errors";
import type { FormulaScalar, FormulaValue } from "../value";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { firstError, numbers, scalarNumber, scalarText, scalarBoolean } from "./utils";

type Arguments = FormulaValue[];
type OptionalArguments = ReadonlyArray<FormulaValue | undefined>;

function firstScalar(value: FormulaValue): FormulaScalar {
  return Array.isArray(value) ? (value.flat()[0] ?? null) : value;
}

function textArgument(value: FormulaValue | undefined, error: FormulaValue): string | FormulaValue {
  if (value === undefined) return error;
  const text = scalarText(firstScalar(value)).trim();
  return text === "" ? error : text;
}

function optionalNumber(value: FormulaValue | undefined): number | undefined {
  return value === undefined ? undefined : scalarNumber(firstScalar(value));
}

function optionalBoolean(value: FormulaValue | undefined): boolean | undefined {
  return value === undefined ? undefined : scalarNumber(firstScalar(value)) !== 0;
}

export function getPivotData(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  const error = firstError(args);
  if (error !== undefined) return error;
  if (args.length < 2 || (args.length - 2) % 2 !== 0) return "#VALUE!";
  const workbook = context?.workbook;
  if (!workbook) return "#NAME?";
  if (!workbook.getPivotValue) return "#REF!";
  const dataField = textArgument(args[0], "#VALUE!");
  const pivotTable = textArgument(args[1], "#VALUE!");
  if (typeof dataField !== "string" || typeof pivotTable !== "string") return "#VALUE!";
  return (
    workbook.getPivotValue({
      dataField,
      pivotTable,
      selectors: Array.from({ length: (args.length - 2) / 2 }, (_, index) => {
        const field = scalarText(firstScalar(args[index * 2 + 2]!));
        const item = scalarText(firstScalar(args[index * 2 + 3]!));
        return { field, item };
      }),
    }) ?? "#REF!"
  );
}

export function realTimeData(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  const error = firstError(args);
  if (error !== undefined) return error;
  const workbook = context?.workbook;
  if (!workbook) return "#NAME?";
  if (!workbook.getRealTimeData) return "#REF!";
  const programId = textArgument(args[0], "#VALUE!");
  if (typeof programId !== "string") return "#VALUE!";
  return (
    workbook.getRealTimeData({
      programId,
      server: scalarText(firstScalar(args[1] ?? null)).trim(),
      topics: args.slice(2).map((value) => scalarText(firstScalar(value))),
    }) ?? "#REF!"
  );
}

export function image(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  const error = firstError(args);
  if (error !== undefined) return error;
  const workbook = context?.workbook;
  if (!workbook?.getImage) return "#REF!";
  const source = textArgument(args[0], "#VALUE!");
  if (typeof source !== "string") return "#VALUE!";
  return (
    workbook.getImage({
      source,
      ...(args[1] === undefined ? {} : { altText: scalarText(firstScalar(args[1])) }),
      ...(args[2] === undefined ? {} : { sizing: scalarNumber(firstScalar(args[2])) }),
      ...(args[3] === undefined ? {} : { height: scalarNumber(firstScalar(args[3])) }),
      ...(args[4] === undefined ? {} : { width: scalarNumber(firstScalar(args[4])) }),
    }) ?? "#REF!"
  );
}

export function stockHistory(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  const error = firstError(args);
  if (error !== undefined) return error;
  const workbook = context?.workbook;
  if (!workbook?.getStockHistory) return "#REF!";
  const symbol = textArgument(args[0], "#VALUE!");
  if (typeof symbol !== "string" || args[1] === undefined) return "#VALUE!";
  return (
    workbook.getStockHistory({
      symbol,
      startDate: firstScalar(args[1]),
      ...(args[2] === undefined ? {} : { endDate: firstScalar(args[2]) }),
      ...(args[3] === undefined ? {} : { interval: scalarNumber(firstScalar(args[3])) }),
      ...(args[4] === undefined ? {} : { headers: scalarBoolean(firstScalar(args[4])) }),
      ...(args[5] === undefined
        ? {}
        : {
            properties: numbers(args[5]).map((property) => {
              if (!Number.isInteger(property) || property < 0 || property > 5)
                throw new FormulaValueError("#VALUE!");
              return property;
            }),
          }),
      ...(args[6] === undefined ? {} : { identifier: scalarText(firstScalar(args[6])) }),
    }) ?? "#REF!"
  );
}

export function translateText(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  const error = firstError(args);
  if (error !== undefined) return error;
  const workbook = context?.workbook;
  if (!workbook?.translateText || args[0] === undefined) return "#REF!";
  return (
    workbook.translateText({
      text: firstScalar(args[0]),
      ...(args[1] === undefined ? {} : { sourceLanguage: scalarText(firstScalar(args[1])) }),
      ...(args[2] === undefined ? {} : { targetLanguage: scalarText(firstScalar(args[2])) }),
    }) ?? "#REF!"
  );
}

export function detectLanguage(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  const error = firstError(args);
  if (error !== undefined) return error;
  const workbook = context?.workbook;
  if (!workbook?.detectLanguage || args[0] === undefined) return "#REF!";
  return workbook.detectLanguage({ text: [firstScalar(args[0])] }) ?? "#REF!";
}

export function fieldValue(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  const error = firstError(args);
  if (error !== undefined) return error;
  if (!context?.workbook) return "#NAME?";
  const workbook = context?.workbook;
  if (!workbook.getFieldValue) return "#REF!";
  const field = textArgument(args[1], "#VALUE!");
  if (typeof field !== "string" || args[0] === undefined) return "#VALUE!";
  return workbook.getFieldValue({ value: firstScalar(args[0]), field }) ?? "#REF!";
}

export function webService(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  const error = firstError(args);
  if (error !== undefined) return error;
  const workbook = context?.workbook;
  if (!workbook) return "#NAME?";
  if (!workbook.fetchWebService) return "#REF!";
  const url = textArgument(args[0], "#VALUE!");
  if (typeof url !== "string") return "#VALUE!";
  return workbook.fetchWebService(url) ?? "#REF!";
}

function macroScalars(values: readonly FormulaValue[]): FormulaScalar[] {
  return values.map((value) => firstScalar(value));
}

export function call(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  const error = firstError(args);
  if (error !== undefined) return error;
  if (args.length < 2) return "#VALUE!";
  const workbook = context?.workbook;
  if (!workbook) return "#NAME?";
  if (!workbook.callMacro) return "#REF!";
  const library = textArgument(args[0], "#VALUE!");
  const procedure = textArgument(args[1], "#VALUE!");
  if (typeof library !== "string" || typeof procedure !== "string") return "#VALUE!";
  return (
    workbook.callMacro({
      library,
      procedure,
      typeText: scalarText(firstScalar(args[2]!)),
      arguments: macroScalars(args.slice(3)),
    }) ?? "#REF!"
  );
}

export function registerId(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  const error = firstError(args);
  if (error !== undefined) return error;
  if (args.length < 2 || args.length > 3) return "#VALUE!";
  const workbook = context?.workbook;
  if (!workbook) return "#NAME?";
  if (!workbook.registerMacro) return "#REF!";
  const library = textArgument(args[0], "#VALUE!");
  const procedure = textArgument(args[1], "#VALUE!");
  if (typeof library !== "string" || typeof procedure !== "string") return "#VALUE!";
  return (
    workbook.registerMacro({
      library,
      procedure,
      ...(args[2] === undefined ? {} : { typeText: scalarText(firstScalar(args[2]!)) }),
    }) ?? "#REF!"
  );
}

export function sqlRequest(
  args: OptionalArguments,
  context?: FormulaEvaluationContext,
  outputReferenceOverride?: string,
): FormulaValue {
  const error = firstError(args.filter((argument) => argument !== undefined));
  if (error !== undefined) return error;
  if (args.length < 1) return "#VALUE!";
  const workbook = context?.workbook;
  if (!workbook) return "#NAME?";
  if (!workbook.executeSql) return "#REF!";
  const connectionString = textArgument(args[0], "#VALUE!");
  const outputReference =
    outputReferenceOverride ??
    (args[1] === undefined ? undefined : textArgument(args[1], "#VALUE!"));
  const query = args[3] === undefined ? undefined : textArgument(args[3], "#VALUE!");
  if (typeof connectionString !== "string" || typeof query !== "string") return "#VALUE!";
  return (
    workbook.executeSql({
      connectionString,
      ...(query === undefined ? {} : { query }),
      ...(typeof outputReference === "string" ? { outputReference } : {}),
      ...(args[2] === undefined ? {} : { driverPrompt: optionalNumber(args[2]) }),
      ...(args[4] === undefined ? {} : { columnNames: optionalBoolean(args[4]) }),
      ...(args[5] === undefined ? {} : { rowNames: optionalBoolean(args[5]) }),
      ...(args[6] === undefined ? {} : { maxRows: optionalNumber(args[6]) }),
      ...(args[7] === undefined ? {} : { maxChanges: optionalNumber(args[7]) }),
      ...(args[8] === undefined ? {} : { timeout: optionalNumber(args[8]) }),
    }) ?? "#REF!"
  );
}

export const webHostFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  ["GETPIVOTDATA", (args, context) => getPivotData(args, context)],
  ["RTD", (args, context) => realTimeData(args, context)],
  ["IMAGE", (args, context) => image(args, context)],
  ["STOCKHISTORY", (args, context) => stockHistory(args, context)],
  ["TRANSLATE", (args, context) => translateText(args, context)],
  ["DETECTLANGUAGE", (args, context) => detectLanguage(args, context)],
  ["FIELDVALUE", (args, context) => fieldValue(args, context)],
  ["WEBSERVICE", (args, context) => webService(args, context)],
  ["CALL", (args, context) => call(args, context)],
  ["REGISTER.ID", (args, context) => registerId(args, context)],
  ["SQL.REQUEST", (args, context) => sqlRequest(args, context)],
]);
