import type { FormulaEvaluationContext } from "../context";
import type { FormulaValue } from "../value";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { scalarNumber, scalarText } from "./utils";

type Arguments = Parameters<FormulaFunctionImplementation>[0];

function text(value: FormulaValue | undefined): string {
  return scalarText(Array.isArray(value) ? (value.flat()[0] ?? null) : (value ?? null));
}

function optionalText(value: FormulaValue | undefined): string | undefined {
  return value === undefined ? undefined : text(value);
}

function optionalNumber(value: FormulaValue | undefined): number | undefined {
  return value === undefined
    ? undefined
    : scalarNumber(Array.isArray(value) ? (value.flat()[0] ?? null) : value);
}

function color(value: Arguments): FormulaValue {
  const channels = value.slice(0, 3).map((channel) => scalarNumber(channel));
  const alpha = value[3] === undefined ? 255 : scalarNumber(value[3]);
  if (channels.some((channel) => channel < 0 || channel > 255) || alpha < 0 || alpha > 255)
    return "#NUM!";
  return (
    (alpha * 2 ** 24 + channels[0]! * 2 ** 16 + channels[1]! * 2 ** 8 + channels[2]!) % 2 ** 31
  );
}

export function cubeKpiMember(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  if (args.length < 3 || args.length > 4 || !context) return "#VALUE!";
  return (
    context.workbook.getCubeKpiMember?.({
      connection: text(args[0]),
      expression: text(args[1]),
      property: text(args[2]),
      caption: optionalText(args[3]),
    }) ?? "#N/A"
  );
}

export function cubeMember(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  if (args.length < 2 || args.length > 3 || !context) return "#VALUE!";
  return (
    context.workbook.getCubeMember?.({
      connection: text(args[0]),
      expression: text(args[1]),
      caption: optionalText(args[2]),
    }) ?? "#N/A"
  );
}

export function cubeMemberProperty(
  args: Arguments,
  context?: FormulaEvaluationContext,
): FormulaValue {
  if (args.length !== 3 || !context) return "#VALUE!";
  return (
    context.workbook.getCubeMemberProperty?.({
      connection: text(args[0]),
      expression: text(args[1]),
      property: text(args[2]),
    }) ?? "#N/A"
  );
}

export function cubeRankedMember(
  args: Arguments,
  context?: FormulaEvaluationContext,
): FormulaValue {
  if (args.length < 3 || args.length > 4 || !context) return "#VALUE!";
  return (
    context.workbook.getCubeRankedMember?.({
      connection: text(args[0]),
      setExpression: text(args[1]),
      rank: scalarNumber(args[2]),
      caption: optionalText(args[3]),
    }) ?? "#N/A"
  );
}

export function cubeSet(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  if (args.length < 2 || args.length > 5 || !context) return "#VALUE!";
  return (
    context.workbook.getCubeSet?.({
      connection: text(args[0]),
      setExpression: text(args[1]),
      caption: optionalText(args[2]),
      sortOrder: optionalNumber(args[3]),
      sortBy: optionalText(args[4]),
    }) ?? "#N/A"
  );
}

export function cubeSetCount(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  if (args.length !== 1 || !context) return "#VALUE!";
  return context.workbook.getCubeSetCount?.({ set: args[0]! }) ?? "#N/A";
}

export function cubeValue(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  if (args.length < 2 || !context) return "#VALUE!";
  return (
    context.workbook.getCubeValue?.({
      connection: text(args[0]),
      expressions: args.slice(1).map(text),
    }) ?? "#N/A"
  );
}

export function dde(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  if (args.length < 3 || args.length > 4 || !context) return "#VALUE!";
  return (
    context.workbook.getDdeData?.({
      server: text(args[0]),
      file: text(args[1]),
      item: text(args[2]),
      mode: optionalText(args[3]),
    }) ?? "#N/A"
  );
}

export function filterXml(args: Arguments, context?: FormulaEvaluationContext): FormulaValue {
  if (args.length !== 2 || !context) return "#VALUE!";
  return (
    context.workbook.evaluateXmlXPath?.({ xml: text(args[0]), expression: text(args[1]) }) ??
    "#VALUE!"
  );
}

export function colorValue(args: Arguments): FormulaValue {
  return args.length < 3 || args.length > 4 ? "#VALUE!" : color(args);
}

export const externalExtraFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  ["COLOR", (args) => colorValue(args)],
  ["DDE", (args, context) => dde(args, context)],
  ["FILTERXML", (args, context) => filterXml(args, context)],
  ["CUBEKPIMEMBER", (args, context) => cubeKpiMember(args, context)],
  ["CUBEMEMBER", (args, context) => cubeMember(args, context)],
  ["CUBEMEMBERPROPERTY", (args, context) => cubeMemberProperty(args, context)],
  ["CUBERANKEDMEMBER", (args, context) => cubeRankedMember(args, context)],
  ["CUBESET", (args, context) => cubeSet(args, context)],
  ["CUBESETCOUNT", (args, context) => cubeSetCount(args, context)],
  ["CUBEVALUE", (args, context) => cubeValue(args, context)],
]);
