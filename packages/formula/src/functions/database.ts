import { FormulaValueError } from "../errors";
import type { FormulaMatrix, FormulaScalar, FormulaValue } from "../value";
import { isMatrix } from "../value";
import { matchesCriterion } from "./criterion";
import type { FormulaFunctionTable } from "./utils";
import { firstError, scalarText } from "./utils";

type Arguments = FormulaValue[];

type DatabaseSelection = {
  fieldIndex: number | undefined;
  records: FormulaScalar[][];
};

function matrix(value: FormulaValue): FormulaMatrix {
  return isMatrix(value) ? value : [[value]];
}

function numericValue(value: FormulaScalar): number | undefined {
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value)))
    return Number(value);
  return undefined;
}

function isNumeric(value: FormulaScalar): boolean {
  return numericValue(value) !== undefined;
}

function fieldIndex(headers: FormulaScalar[], field: FormulaValue): number | undefined {
  const scalar = Array.isArray(field) ? field.flat()[0] : field;
  if (scalar === undefined || scalar === null || scalarText(scalar).trim() === "") return undefined;
  if (typeof scalar === "number" && Number.isInteger(scalar)) {
    if (scalar < 1 || scalar > headers.length) throw new FormulaValueError("#VALUE!");
    return scalar - 1;
  }
  const name = scalarText(scalar).trim().toUpperCase();
  const index = headers.findIndex((header) => scalarText(header).trim().toUpperCase() === name);
  if (index < 0) throw new FormulaValueError("#VALUE!");
  return index;
}

function selection(
  database: FormulaValue,
  field: FormulaValue,
  criteria: FormulaValue,
): DatabaseSelection {
  const rows = matrix(database);
  const criterionRows = matrix(criteria);
  if (rows.length === 0 || criterionRows.length === 0) throw new FormulaValueError("#VALUE!");
  const headers = rows[0]!;
  const criterionHeaders = criterionRows[0]!;
  const indexes = criterionHeaders.map((header) => {
    const index = fieldIndex(headers, header);
    if (index === undefined) throw new FormulaValueError("#VALUE!");
    return index;
  });
  if (indexes.length === 0) throw new FormulaValueError("#VALUE!");
  const records = rows.slice(1).filter((record) =>
    criterionRows.slice(1).some((row) =>
      row.every((criterion, position) => {
        if (criterion === null || criterion === "") return true;
        return matchesCriterion(record[indexes[position]!] ?? null, criterion);
      }),
    ),
  );
  return { fieldIndex: fieldIndex(headers, field), records };
}

function selectedScalars(selection: DatabaseSelection): FormulaScalar[] {
  if (selection.fieldIndex === undefined)
    return selection.records.map((record) => record.find((entry) => entry !== null) ?? null);
  return selection.records.map((record) => record[selection.fieldIndex!] ?? null);
}

function selectedNumbers(selection: DatabaseSelection): number[] {
  if (selection.fieldIndex === undefined)
    return selection.records
      .filter((record) => record.some((entry) => isNumeric(entry)))
      .map((record) => record.find((entry) => isNumeric(entry))!)
      .map((entry) => numericValue(entry)!);
  return selectedScalars(selection)
    .filter((entry) => isNumeric(entry))
    .map((entry) => numericValue(entry)!);
}

function sum(values: number[]): number {
  return values.reduce((total, value) => total + value, 0);
}

function variance(values: number[], sample: boolean): number {
  const divisor = values.length - (sample ? 1 : 0);
  const center = sum(values) / values.length;
  return values.reduce((total, value) => total + (value - center) ** 2, 0) / divisor;
}

function maximum(values: number[]): number {
  let result = values[0] ?? Number.NEGATIVE_INFINITY;
  for (const value of values) result = value > result ? value : result;
  return result;
}

function minimum(values: number[]): number {
  let result = values[0] ?? Number.POSITIVE_INFINITY;
  for (const value of values) result = value < result ? value : result;
  return result;
}

function aggregate(name: string, args: Arguments): FormulaValue {
  const error = firstError(args);
  if (error !== undefined) return error;
  const selected = selection(args[0]!, args[1]!, args[2]!);
  const values = selectedNumbers(selected);
  switch (name) {
    case "DCOUNT":
      return selected.fieldIndex === undefined
        ? selected.records.filter((record) => record.some((entry) => isNumeric(entry))).length
        : values.length;
    case "DCOUNTA":
      return selected.fieldIndex === undefined
        ? selected.records.filter((record) => record.some((entry) => entry !== null)).length
        : selectedScalars(selected).filter((entry) => entry !== null).length;
    case "DAVERAGE":
      return values.length === 0 ? "#DIV/0!" : sum(values) / values.length;
    case "DMAX":
      return values.length === 0 ? 0 : maximum(values);
    case "DMIN":
      return values.length === 0 ? 0 : minimum(values);
    case "DPRODUCT":
      return values.length === 0 ? 0 : values.reduce((product, value) => product * value, 1);
    case "DSTDEV":
      return values.length < 2 ? "#DIV/0!" : Math.sqrt(variance(values, true));
    case "DSTDEVP":
      return values.length === 0 ? "#DIV/0!" : Math.sqrt(variance(values, false));
    case "DSUM":
      return sum(values);
    case "DVAR":
      return values.length < 2 ? "#DIV/0!" : variance(values, true);
    case "DVARP":
      return values.length === 0 ? "#DIV/0!" : variance(values, false);
    case "DGET": {
      if (selected.fieldIndex === undefined) return "#VALUE!";
      const entries = selectedScalars(selected);
      if (entries.length === 0) return "#VALUE!";
      if (entries.length > 1) return "#NUM!";
      return entries[0] ?? null;
    }
    default:
      return "#VALUE!";
  }
}

function databaseFunction(name: string): (args: Arguments) => FormulaValue {
  return (args) => aggregate(name, args);
}

export const databaseFunctions: FormulaFunctionTable = new Map(
  [
    "DAVERAGE",
    "DCOUNT",
    "DCOUNTA",
    "DGET",
    "DMAX",
    "DMIN",
    "DPRODUCT",
    "DSTDEV",
    "DSTDEVP",
    "DSUM",
    "DVAR",
    "DVARP",
  ].map((name) => [name, databaseFunction(name)]),
);
