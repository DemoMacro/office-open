import { FormulaValueError } from "../errors";
import type { FormulaErrorValue, FormulaMatrix, FormulaScalar, FormulaValue } from "../value";
import { isFormulaErrorValue, isMatrix } from "../value";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { scalarBoolean, scalarNumber, scalarText } from "./utils";

function matrix(value: FormulaValue): FormulaMatrix {
  return isMatrix(value) ? value : [[value]];
}

function scalarArray(value: FormulaValue): FormulaScalar[] {
  return isMatrix(value) ? value.flat() : [value];
}

function matrixColumn(table: FormulaMatrix, column: number): FormulaMatrix {
  const output: FormulaMatrix = [];
  for (const row of table) output.push([row[column - 1] ?? "#REF!"]);
  return output;
}

function sheetPrefix(sheetText: FormulaValue): string {
  const input = scalarText(sheetText);
  if (input === "") return "";
  const quoted = /^'((?:[^']|'')*)'$/u.exec(input);
  const name = quoted ? (quoted[1] ?? "").replace(/''/gu, "'") : input;
  const simple = /^[A-Za-z_][A-Za-z0-9_]*$/u.test(name);
  return simple ? `${name}!` : `'${name.replace(/'/gu, "''")}'!`;
}

function columnLetter(value: number): string {
  let remaining = value;
  let output = "";
  while (remaining > 0) {
    const modulo = (remaining - 1) % 26;
    output = String.fromCharCode(65 + modulo) + output;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return output;
}

function compare(left: FormulaScalar, right: FormulaScalar): number {
  const rank = (value: FormulaScalar): number =>
    typeof value === "number"
      ? 0
      : typeof value === "string"
        ? 1
        : typeof value === "boolean"
          ? 2
          : 3;
  const rankDifference = rank(left) - rank(right);
  if (rankDifference !== 0) return rankDifference;
  if (typeof left === "number" && typeof right === "number") return left - right;
  if (typeof left === "boolean" && typeof right === "boolean") return Number(left) - Number(right);
  return scalarText(left).localeCompare(scalarText(right), undefined, { sensitivity: "base" });
}

function lookupCompare(left: FormulaScalar, right: FormulaScalar): number | undefined {
  if (typeof left === "number" && typeof right === "number") return left - right;
  if (typeof left === "string" && typeof right === "string")
    return left.localeCompare(right, undefined, { sensitivity: "base" });
  if (typeof left === "boolean" && typeof right === "boolean") return Number(left) - Number(right);
  return undefined;
}

function scalarKey(value: FormulaScalar): string {
  const text = typeof value === "string" ? value.toLowerCase() : scalarText(value).toLowerCase();
  return `${typeof value}:${text}`;
}

export function aggregateValues(args: FormulaValue[], ignoreErrors: boolean): number[] {
  const entries = args.flatMap(scalarArray);
  const error = entries.find((value): value is FormulaErrorValue => isFormulaErrorValue(value));
  if (error !== undefined && !ignoreErrors) throw new FormulaValueError(error);
  return entries.filter((value): value is number => typeof value === "number");
}

function variance(input: number[], population: boolean): number {
  if (input.length < (population ? 1 : 2)) throw new FormulaValueError("#DIV/0!");
  const mean = input.reduce((sum, value) => sum + value, 0) / input.length;
  const squared = input.reduce((sum, value) => sum + (value - mean) ** 2, 0);
  return squared / (population ? input.length : input.length - 1);
}

function maximum(input: number[]): number {
  let result = input[0] ?? Number.NEGATIVE_INFINITY;
  for (const value of input) result = value > result ? value : result;
  return result;
}

function minimum(input: number[]): number {
  let result = input[0] ?? Number.POSITIVE_INFINITY;
  for (const value of input) result = value < result ? value : result;
  return result;
}

function percentile(input: number[], value: number, exclusive: boolean): number | string {
  if (exclusive ? value <= 0 || value >= 1 : value < 0 || value > 1) return "#NUM!";
  const sorted = [...input].sort((left, right) => left - right);
  const rank = exclusive ? (sorted.length + 1) * value : (sorted.length - 1) * value;
  const lower = Math.floor(rank);
  const upper = Math.ceil(rank);
  if (lower === upper) return sorted[lower] ?? "#NUM!";
  return (sorted[lower] ?? 0) + ((sorted[upper] ?? 0) - (sorted[lower] ?? 0)) * (rank - lower);
}

function mode(input: number[]): number | string {
  const counts = new Map<number, { count: number; index: number }>();
  input.forEach((value, index) => {
    const current = counts.get(value);
    counts.set(value, { count: (current?.count ?? 0) + 1, index: current?.index ?? index });
  });
  const winner = [...counts.entries()]
    .filter(([, detail]) => detail.count > 1)
    .sort((left, right) => right[1].count - left[1].count || left[1].index - right[1].index)
    .at(0);
  return winner?.[0] ?? "#N/A";
}

export const lookupReferenceFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  [
    "INDEX",
    ([source, row, column]): FormulaValue => {
      const table = matrix(source);
      const requestedRow = row === undefined || row === null ? undefined : scalarNumber(row);
      const requestedColumn =
        column === undefined || column === null ? undefined : scalarNumber(column);
      const selectAll =
        (requestedRow === undefined || requestedRow === 0) &&
        (requestedColumn === undefined || requestedColumn === 0);
      if (selectAll) return table;
      if (requestedColumn === undefined && table.length === 1) {
        const selected = table[0]?.[(requestedRow ?? 1) - 1];
        return selected === undefined ? "#REF!" : selected;
      }
      if (requestedColumn === undefined && table.every((tableRow) => tableRow.length === 1)) {
        const selected = table[(requestedRow ?? 1) - 1]?.[0];
        return selected === undefined ? "#REF!" : selected;
      }
      if ((requestedRow ?? 0) === 0) return matrixColumn(table, requestedColumn ?? 1);
      if ((requestedColumn ?? 0) === 0) {
        const selected = table[(requestedRow ?? 1) - 1];
        return selected ? [selected] : "#REF!";
      }
      return table[(requestedRow ?? 1) - 1]?.[(requestedColumn ?? 1) - 1] ?? "#REF!";
    },
  ],
  [
    "ADDRESS",
    ([row, columnValue, absolute, a1Style, sheetText]) => {
      const rowNumber = scalarNumber(row);
      const column = scalarNumber(columnValue);
      const mode = absolute === undefined ? 1 : scalarNumber(absolute);
      if (
        mode < 1 ||
        mode > 4 ||
        !Number.isInteger(rowNumber) ||
        rowNumber < 1 ||
        rowNumber > 1_048_576 ||
        !Number.isInteger(column) ||
        column < 1 ||
        column > 16_384
      )
        throw new FormulaValueError("#VALUE!");
      const useA1 = a1Style === undefined || scalarBoolean(a1Style);
      const absoluteRow = mode === 1 || mode === 2;
      const absoluteColumn = mode === 1 || mode === 3;
      const rowPart = useA1
        ? `${absoluteRow ? "$" : ""}${rowNumber}`
        : `R${absoluteRow ? rowNumber : `[${rowNumber}]`}`;
      const columnPart = useA1
        ? `${absoluteColumn ? "$" : ""}${columnLetter(column)}`
        : `${absoluteColumn ? `C${column}` : `C[${column}]`}`;
      const prefix = sheetText === undefined ? "" : sheetPrefix(sheetText);
      return useA1 ? `${prefix}${columnPart}${rowPart}` : `${prefix}${rowPart}${columnPart}`;
    },
  ],
  [
    "AGGREGATE",
    (args) => {
      const type = scalarNumber(args[0]!);
      const options = scalarNumber(args[1]!);
      if (
        !Number.isInteger(type) ||
        type < 1 ||
        type > 19 ||
        !Number.isInteger(options) ||
        options < 0 ||
        options > 7
      )
        return "#VALUE!";
      const ignoreErrors = options === 2 || options === 3 || options === 6 || options === 7;
      if (type === 3) {
        const entries = Array.isArray(args[2]) ? args[2].flat() : [args[2]!];
        return entries.filter(
          (value) =>
            value !== null && value !== "" && (ignoreErrors ? !isFormulaErrorValue(value) : true),
        ).length;
      }
      const input = aggregateValues([args[2]!], ignoreErrors);
      switch (type) {
        case 1:
          return input.length
            ? input.reduce((sum: number, value) => sum + value, 0) / input.length
            : "#DIV/0!";
        case 2:
        case 3:
          return input.length;
        case 4:
          return input.length ? maximum(input) : "#NUM!";
        case 5:
          return input.length ? minimum(input) : "#NUM!";
        case 6:
          return input.reduce((product, value) => product * value, 1);
        case 7:
          return Math.sqrt(variance(input, false));
        case 8:
          return Math.sqrt(variance(input, true));
        case 9:
          return input.reduce((sum, value) => sum + value, 0);
        case 10:
          return variance(input, false);
        case 11:
          return variance(input, true);
        case 12: {
          const sorted = [...input].sort((left, right) => left - right);
          const middle = Math.floor(sorted.length / 2);
          if (sorted.length === 0) return "#NUM!";
          return sorted.length % 2 === 1
            ? sorted[middle]!
            : (sorted[middle - 1]! + sorted[middle]!) / 2;
        }
        case 13:
          return mode(input);
        case 14:
        case 15: {
          if (args.length !== 4) return "#VALUE!";
          const rank = scalarNumber(args[3]!);
          const sorted = [...input].sort((left, right) =>
            type === 14 ? right - left : left - right,
          );
          return Number.isInteger(rank) && rank >= 1 && rank <= input.length
            ? (sorted[rank - 1] ?? "#NUM!")
            : "#NUM!";
        }
        case 16:
        case 18:
          if (args.length !== 4) return "#VALUE!";
          return percentile(input, scalarNumber(args[3]!), type === 18);
        case 17:
        case 19:
          if (args.length !== 4) return "#VALUE!";
          return percentile(input, (scalarNumber(args[3]!) - 1) / 4, type === 19);
        default:
          return "#VALUE!";
      }
    },
  ],
  ["AREAS", (args) => args.length],
  ["CHOOSE", ([index, ...choices]) => choices[scalarNumber(index) - 1] ?? "#VALUE!"],
  [
    "COLUMNS",
    ([range]) => matrix(range).reduce((maximum, row) => Math.max(maximum, row.length), 0),
  ],
  ["ROWS", ([range]) => matrix(range).length],
  [
    "HLOOKUP",
    ([needle, table, index, approximate]) =>
      lookupVector(
        scalarArray(needle)[0] ?? null,
        matrix(table),
        scalarNumber(index),
        approximate === undefined || scalarNumber(approximate) !== 0,
        false,
      ),
  ],
  [
    "LOOKUP",
    ([needle, lookup, returns]) => {
      const keys = scalarArray(lookup);
      const source = returns === undefined ? keys : scalarArray(returns);
      const target = scalarArray(needle)[0] ?? null;
      const candidates = keys
        .map((key, index) => [key, index] as const)
        .filter(([key]) => (lookupCompare(key, target) ?? Number.NaN) <= 0);
      const last = candidates.at(-1);
      return last ? (source[last[1]] ?? "#N/A") : "#N/A";
    },
  ],
  [
    "MATCH",
    ([needle, lookup, type]) => {
      const entries = scalarArray(lookup);
      const targetScalar = scalarArray(needle)[0] ?? null;
      const matchType = type === undefined ? 1 : scalarNumber(type);
      if (matchType === 0)
        return (
          entries.findIndex((entry) => scalarKey(entry) === scalarKey(targetScalar)) + 1 || "#N/A"
        );
      let result: number | string = "#N/A";
      entries.forEach((entry, index) => {
        const comparison = lookupCompare(entry, targetScalar);
        if (comparison === undefined) return;
        const direction = matchType > 0 ? comparison <= 0 : comparison >= 0;
        if (direction) result = index + 1;
      });
      return result;
    },
  ],
  [
    "SORT",
    ([range, index, order, byColumn]) => {
      const source = matrix(range);
      const position = index === undefined ? 1 : scalarNumber(index);
      const direction = order === undefined ? 1 : scalarNumber(order);
      const horizontal = byColumn !== undefined && scalarText(byColumn).toUpperCase() === "TRUE";
      const rows = horizontal ? transpose(source) : source;
      const sorted = [...rows].sort(
        (left, right) =>
          compare(left[position - 1] ?? null, right[position - 1] ?? null) * direction,
      );
      return horizontal ? transpose(sorted) : sorted;
    },
  ],
  [
    "VLOOKUP",
    ([needle, table, index, approximate]) =>
      lookupVector(
        scalarArray(needle)[0] ?? null,
        matrix(table),
        scalarNumber(index),
        approximate === undefined || scalarNumber(approximate) !== 0,
        true,
      ),
  ],
]);

function lookupVector(
  target: FormulaScalar,
  table: FormulaMatrix,
  index: number,
  approximate: boolean,
  vertical: boolean,
): FormulaValue {
  const rows = vertical ? table : transpose(table);
  const first = rows.map((row) => row[0] ?? null);
  const row = approximate
    ? first
        .map((entry, position) => [entry, position] as const)
        .filter(([entry]) => (lookupCompare(entry, target) ?? Number.NaN) <= 0)
        .at(-1)?.[1]
    : first.findIndex((entry) => scalarKey(entry) === scalarKey(target));
  if (row === -1 || row === undefined) return "#N/A";
  return rows[row]?.[index - 1] ?? "#REF!";
}

function transpose(input: FormulaMatrix): FormulaMatrix {
  const width = Math.max(0, ...input.map((row) => row.length));
  return Array.from({ length: width }, (_, column) =>
    Array.from({ length: input.length }, (_, row) => input[row]?.[column] ?? null),
  );
}
