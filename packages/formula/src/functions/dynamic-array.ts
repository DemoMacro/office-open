import type { FormulaEvaluationContext } from "../context";
import { FormulaValueError } from "../errors";
import type { FormulaMatrix, FormulaScalar, FormulaValue } from "../value";
import { formulaError, isFormulaErrorValue, isMatrix } from "../value";
import { matchesCriterion, scalarCriterion } from "./criterion";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import {
  assertFormulaCellCount,
  isError,
  numbers,
  scalarBoolean,
  scalarNumber,
  scalarText,
} from "./utils";

type Arguments = Parameters<FormulaFunctionImplementation>[0];

function matrix(value: FormulaValue): FormulaMatrix {
  return isMatrix(value) ? value : [[value]];
}

function height(input: FormulaMatrix): number {
  return input.length;
}

function width(input: FormulaMatrix): number {
  let maximum = 0;
  for (const row of input) maximum = Math.max(maximum, row.length);
  return maximum;
}

function cell(input: FormulaMatrix, row: number, column: number): FormulaScalar {
  return input[row]?.[column] ?? null;
}

function transpose(input: FormulaMatrix): FormulaMatrix {
  return Array.from({ length: width(input) }, (_, column) =>
    Array.from({ length: height(input) }, (_, row) => cell(input, row, column)),
  );
}

function flattenMatrix(input: FormulaValue): FormulaScalar[] {
  return scalars(input);
}

function flattenOrdered(input: FormulaValue, byColumn: boolean): FormulaScalar[] {
  const source = matrix(input);
  return (byColumn ? transpose(source) : source).flat();
}

function scalars(input: FormulaValue): FormulaScalar[] {
  return isMatrix(input) ? input.flat() : [input];
}

function arrayLength(value: FormulaValue, minimum: number): number {
  const length = scalarNumber(value);
  if (!Number.isInteger(length) || length < minimum || length > Number.MAX_SAFE_INTEGER)
    throw new FormulaValueError("#VALUE!");
  return length;
}

function shape(rows: number, columns: number, fill: FormulaScalar = null): FormulaMatrix {
  return Array.from({ length: Math.max(0, Math.trunc(rows)) }, () =>
    Array.from({ length: Math.max(0, Math.trunc(columns)) }, () => fill),
  );
}

function stack(input: Arguments, horizontal: boolean): FormulaMatrix {
  const matrices = input.map(matrix);
  const columnCount = horizontal
    ? matrices.reduce((total, value) => total + width(value), 0)
    : Math.max(0, ...matrices.map(width));
  const rowCount = horizontal
    ? Math.max(0, ...matrices.map(height))
    : matrices.reduce((total, value) => total + height(value), 0);
  assertFormulaCellCount(Math.max(rowCount, 1), Math.max(columnCount, 1));
  if (!horizontal)
    return matrices.flatMap((value) =>
      Array.from({ length: height(value) }, (_, row) =>
        Array.from(
          { length: columnCount },
          (_, column) => cell(value, row, column) ?? formulaError("#N/A"),
        ),
      ),
    );
  return Array.from({ length: rowCount }, (_, row) =>
    matrices.flatMap((value) =>
      Array.from(
        { length: width(value) },
        (_, column) => cell(value, row, column) ?? formulaError("#N/A"),
      ),
    ),
  );
}

function takeDrop(
  input: FormulaValue,
  count: number,
  take: boolean,
  horizontal: boolean,
): FormulaMatrix {
  const source = matrix(input);
  const total = horizontal ? width(source) : height(source);
  const selected = take
    ? Math.min(Math.max(0, count), total)
    : Math.max(0, total - Math.abs(count < 0 ? Math.max(count, -total) : count));
  const offset = take ? 0 : count < 0 ? 0 : Math.abs(count);
  return horizontal
    ? source.map((row) => row.slice(offset, offset + selected))
    : source.slice(offset, offset + selected);
}

function wrap(
  input: FormulaValue,
  count: number,
  byColumn: boolean,
  pad: FormulaScalar = formulaError("#N/A"),
): FormulaMatrix {
  const entries = flattenMatrix(input);
  const size = Math.max(1, Math.trunc(count));
  const rows = byColumn ? size : Math.ceil(entries.length / size);
  const columns = byColumn ? Math.ceil(entries.length / size) : size;
  assertFormulaCellCount(rows, columns);
  if (byColumn) {
    return Array.from({ length: size }, (_, row) =>
      Array.from({ length: columns }, (_, column) => entries[column * size + row] ?? pad),
    );
  }
  return Array.from({ length: rows }, (_, row) =>
    Array.from({ length: size }, (_, column) => entries[row * size + column] ?? pad),
  );
}

function pivot(input: Arguments, context: FormulaEvaluationContext): FormulaMatrix {
  const filter = input[9];
  const sourceRows = matrix(input[0]!).length;
  const columnRows = matrix(input[1]!).length;
  const valueRows = scalars(input[2]!).length;
  if (columnRows !== sourceRows || valueRows !== sourceRows) throw new FormulaValueError("#VALUE!");
  const selectedRows = (filter === undefined ? [] : scalars(filter)).map(scalarBoolean);
  if (filter !== undefined && selectedRows.length !== sourceRows)
    throw new FormulaValueError("#VALUE!");
  const rowSource =
    selectedRows.length === sourceRows
      ? matrix(input[0]!).filter((_, index) => selectedRows[index])
      : matrix(input[0]!);
  const valueSource =
    selectedRows.length === sourceRows
      ? scalars(input[2]!).filter((_, index) => selectedRows[index])
      : scalars(input[2]!);
  const values = valueSource;
  const columnSource =
    selectedRows.length === sourceRows
      ? matrix(input[1]!).filter((_, index) => selectedRows[index])
      : matrix(input[1]!);
  const includeHeaders = input[4] !== undefined && scalarNumber(input[4]) >= 2;
  const includeRowTotal = input[5] !== undefined && scalarNumber(input[5]) >= 1;
  const includeColumnTotal = input[7] !== undefined && scalarNumber(input[7]) >= 1;
  const rowKeyWidth = width(rowSource);
  const rowGroups = groupRows(rowSource);
  const columnGroups = groupRows(columnSource);
  const outputRows = rowGroups.length + (includeHeaders ? 1 : 0) + (includeRowTotal ? 1 : 0);
  const outputColumns = rowKeyWidth + columnGroups.length + (includeColumnTotal ? 1 : 0);
  assertFormulaCellCount(outputRows, outputColumns);
  const output: FormulaMatrix = includeHeaders
    ? [
        [
          "Group",
          ...columnGroups.map((group) => group.key.map((value) => scalarText(value)).join(" / ")),
          ...(includeColumnTotal ? ["Total"] : []),
        ],
      ]
    : [];
  output.push(
    ...rowGroups.map((rowGroup) => [
      ...rowGroup.key,
      ...columnGroups.map((columnGroup) =>
        aggregateIndexes(
          values,
          intersectionIndexes(rowGroup.indexes, columnGroup.indexes),
          input[3]!,
          context,
        ),
      ),
      ...(includeColumnTotal
        ? [aggregateIndexes(values, rowGroup.indexes, input[3]!, context)]
        : []),
    ]),
  );
  if (includeRowTotal)
    output.push([
      "Total",
      ...columnGroups.map((columnGroup) =>
        aggregateIndexes(values, columnGroup.indexes, input[3]!, context),
      ),
      ...(includeColumnTotal
        ? [
            aggregateIndexes(
              values,
              values.map((_, index) => index),
              input[3]!,
              context,
            ),
          ]
        : []),
    ]);
  if (includeHeaders && output[0])
    output[0] = [
      ...Array.from({ length: rowKeyWidth }, () => "Group"),
      ...output[0].slice(1, 1 + columnGroups.length),
      ...(includeColumnTotal ? ["Total"] : []),
    ];
  return output;
}

function blank(value: FormulaScalar): boolean {
  return value === null || value === "";
}

function flattenIgnoring(input: FormulaValue, ignore: number, byColumn: boolean): FormulaScalar[] {
  return flattenOrdered(input, byColumn).filter((value) => {
    if ((ignore & 1) !== 0 && blank(value)) return false;
    return !((ignore & 2) !== 0 && isError(value));
  });
}

function trimBounds(length: number, mode: number, filled: (index: number) => boolean) {
  const indexes = Array.from({ length }, (_, index) => index).filter(filled);
  const first = indexes[0] ?? 0;
  const last = (indexes.at(-1) ?? -1) + 1;
  return {
    start: (mode & 2) === 0 ? 0 : first,
    end: (mode & 1) === 0 ? length : last,
  };
}

function trimMatrix(input: FormulaMatrix, rowMode: number, columnMode: number): FormulaMatrix {
  const rows = trimBounds(input.length, rowMode, (row) =>
    input[row]!.some((value) => !blank(value)),
  );
  const columns = trimBounds(width(input), columnMode, (column) =>
    input.some((row) => !blank(row[column] ?? null)),
  );
  return input.slice(rows.start, rows.end).map((row) => row.slice(columns.start, columns.end));
}

function scalarKey(value: FormulaScalar): string {
  const text = typeof value === "string" ? value.toLowerCase() : scalarText(value);
  return `${typeof value}:${text}`;
}

function uniqueRows(rows: FormulaMatrix, exactlyOnce: boolean): FormulaMatrix {
  const counts = new Map<string, number>();
  const keys = rows.map((row) => row.map(scalarKey).join("\0"));
  for (const key of keys) counts.set(key, (counts.get(key) ?? 0) + 1);
  const seen = new Set<string>();
  return rows.filter((_, index) => {
    const key = keys[index]!;
    if (seen.has(key)) return false;
    seen.add(key);
    return !exactlyOnce || counts.get(key) === 1;
  });
}

function comparable(left: FormulaScalar, right: FormulaScalar): number | undefined {
  if (typeof left === "number" && typeof right === "number") return left - right;
  if (typeof left === "boolean" && typeof right === "boolean") return Number(left) - Number(right);
  if (typeof left === "string" && typeof right === "string")
    return left.localeCompare(right, undefined, { sensitivity: "base" });
  return undefined;
}

function exactMatch(key: FormulaScalar, target: FormulaScalar, mode: number): boolean {
  if (mode !== 2 || typeof target !== "string") return scalarKey(key) === scalarKey(target);
  return matchesCriterion(key, `=${target}`);
}

function lookupIndex(
  keys: FormulaScalar[],
  target: FormulaScalar,
  matchMode: number,
  searchMode: number,
): number {
  const ordered = searchMode === -1 ? [...keys].reverse() : keys;
  const offset = searchMode === -1 ? keys.length - 1 : 0;
  const exact = ordered.findIndex((key) => exactMatch(key, target, matchMode));
  if (exact >= 0) return searchMode === -1 ? offset - exact : exact;
  if (matchMode === 0 || matchMode === 2) return -1;

  let nearest = -1;
  ordered.forEach((key, index) => {
    const comparison = comparable(key, target);
    if (comparison === undefined) return;
    const matchesDirection = matchMode === -1 ? comparison <= 0 : comparison >= 0;
    if (!matchesDirection) return;
    const nearestComparison = nearest < 0 ? undefined : comparable(ordered[nearest]!, target);
    if (nearestComparison === undefined || Math.abs(comparison) < Math.abs(nearestComparison))
      nearest = index;
  });
  if (nearest < 0) return -1;
  return searchMode === -1 ? offset - nearest : nearest;
}

function lookupArguments(
  lookup: FormulaValue,
  returns: FormulaValue,
): {
  keys: FormulaScalar[];
  results: FormulaScalar[];
} {
  const keys = scalars(lookup);
  const results = scalars(returns);
  if (keys.length !== results.length) throw new FormulaValueError("#VALUE!");
  return { keys, results };
}

function lookupMode(
  value: FormulaValue | undefined,
  allowed: readonly number[],
  fallback: number,
): number | undefined {
  if (value === undefined) return fallback;
  const mode = scalarNumber(value);
  return allowed.includes(mode) ? mode : undefined;
}

function groupRows(rowFields: FormulaValue): Array<{ key: FormulaScalar[]; indexes: number[] }> {
  const rows = matrix(rowFields);
  const keyLength = Math.max(1, width(rows));
  const groups = new Map<
    string | number | boolean,
    {
      key: FormulaScalar[];
      indexes: number[];
    }
  >();
  rows.forEach((row, index) => {
    if (keyLength === 1) {
      const value = row[0] ?? null;
      const identity =
        typeof value === "string"
          ? `s:${value.toLowerCase()}`
          : typeof value === "number" || typeof value === "boolean"
            ? value
            : value === null
              ? "null"
              : `e:${value.value}`;
      const group = groups.get(identity);
      if (group) group.indexes.push(index);
      else groups.set(identity, { key: [value], indexes: [index] });
      return;
    }
    const key = row.slice(0, keyLength);
    const identity = key.map(scalarKey).join("\0");
    const group = groups.get(identity);
    if (group) group.indexes.push(index);
    else groups.set(identity, { key, indexes: [index] });
  });
  return [...groups.values()];
}

function intersectionIndexes(left: number[], right: number[]): number[] {
  const [outer, inner] = left.length <= right.length ? [left, right] : [right, left];
  const membership = new Set(inner);
  return outer.filter((index) => membership.has(index));
}

function aggregateIndexes(
  values: FormulaScalar[],
  indexes: number[],
  aggregator: FormulaValue,
  context: FormulaEvaluationContext,
): FormulaScalar {
  if (typeof aggregator === "string" && isFormulaErrorValue(aggregator)) return aggregator;
  if (typeof aggregator === "string") {
    const direct = directAggregate(values, indexes, aggregator);
    if (direct !== undefined) return direct;
  } else {
    return "#VALUE!";
  }
  if (!context.invokeLambdaValue) return "#VALUE!";
  try {
    const selected: FormulaMatrix = [indexes.map((index) => values[index] ?? null)];
    const result = context.invokeLambdaValue(aggregator, [selected]);
    const scalar = isMatrix(result) ? (result.flat()[0] ?? null) : result;
    if (typeof scalar === "string" && scalar.startsWith("__FORMULA_LAMBDA__")) return "#VALUE!";
    return scalar;
  } catch (error) {
    return error instanceof FormulaValueError ? error.value : "#VALUE!";
  }
}

function directAggregate(
  values: FormulaScalar[],
  indexes: number[],
  name: string,
): FormulaScalar | undefined {
  const statistic = name.toUpperCase();
  const needsNumericSamples =
    statistic === "MEDIAN" || statistic === "STDEV" || statistic === "VAR";
  let numericCount = 0;
  let nonEmptyCount = 0;
  let sum = 0;
  let product = 1;
  let minimum = Number.POSITIVE_INFINITY;
  let maximum = Number.NEGATIVE_INFINITY;
  const samples: number[] = [];
  for (const index of indexes) {
    const value = values[index] ?? null;
    if (value !== null && value !== "") nonEmptyCount += 1;
    if (typeof value !== "number" || !Number.isFinite(value)) continue;
    numericCount += 1;
    sum += value;
    product *= value;
    minimum = Math.min(minimum, value);
    maximum = Math.max(maximum, value);
    if (needsNumericSamples) samples.push(value);
  }
  switch (statistic) {
    case "SUM":
      return sum;
    case "AVERAGE":
      return numericCount ? sum / numericCount : "#DIV/0!";
    case "COUNT":
      return numericCount;
    case "COUNTA":
      return nonEmptyCount;
    case "MAX":
      return numericCount ? maximum : 0;
    case "MIN":
      return numericCount ? minimum : 0;
    case "MEDIAN": {
      const ordered = [...samples].sort((left, right) => left - right);
      const middle = Math.floor(ordered.length / 2);
      return ordered.length === 0
        ? "#NUM!"
        : ordered.length % 2
          ? ordered[middle]!
          : (ordered[middle - 1]! + ordered[middle]!) / 2;
    }
    case "PRODUCT":
      return product;
    case "STDEV":
      return numericCount > 1 ? sampleStandardDeviation(samples) : "#DIV/0!";
    case "VAR":
      return numericCount > 1 ? sampleVariance(samples) : "#DIV/0!";
    default:
      return undefined;
  }
}

function sampleVariance(input: number[]): number {
  const mean = input.reduce((sum, value) => sum + value, 0) / input.length;
  return input.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (input.length - 1);
}

function sampleStandardDeviation(input: number[]): number {
  return Math.sqrt(sampleVariance(input));
}

function groupBy(input: Arguments, context: FormulaEvaluationContext): FormulaMatrix {
  const source = matrix(input[0]!);
  const filter = input[6];
  const selectedRows = (filter === undefined ? [] : scalars(filter)).map(scalarBoolean);
  const rawValues = scalars(input[1]!);
  if (rawValues.length !== source.length) throw new FormulaValueError("#VALUE!");
  if (filter !== undefined && selectedRows.length !== source.length)
    throw new FormulaValueError("#VALUE!");
  const filteredSource =
    selectedRows.length === source.length
      ? source.filter((_, index) => selectedRows[index])
      : source;
  const values =
    selectedRows.length === source.length
      ? rawValues.filter((_, index) => selectedRows[index])
      : rawValues;
  const includeHeader = input[3] !== undefined && scalarNumber(input[3]) >= 2;
  const includeTotal = input[4] !== undefined && scalarNumber(input[4]) >= 1;
  const groups = groupRows(filteredSource);
  const sortOrder = input[5] === undefined ? 0 : scalarNumber(input[5]);
  const sortField = Math.abs(sortOrder) - 1;
  const sorted =
    sortOrder === 0
      ? groups
      : [...groups].sort((left, right) => {
          const comparison = comparable(left.key[sortField] ?? null, right.key[sortField] ?? null);
          if (comparison !== undefined) return sortOrder > 0 ? comparison : -comparison;
          return (
            scalarKey(left.key[sortField] ?? null).localeCompare(
              scalarKey(right.key[sortField] ?? null),
            ) * (sortOrder > 0 ? 1 : -1)
          );
        });
  assertFormulaCellCount(
    sorted.length + (includeHeader ? 1 : 0) + (includeTotal ? 1 : 0),
    width(filteredSource) + 1,
  );
  const output: FormulaMatrix = includeHeader
    ? [
        Array.from({ length: width(matrix(input[0]!)) + 1 }, (_, column) =>
          column === 0 ? "Group" : "Result",
        ),
      ]
    : [];
  for (const group of sorted)
    output.push([...group.key, aggregateIndexes(values, group.indexes, input[2]!, context)]);
  if (includeTotal)
    output.push([
      "Total",
      aggregateIndexes(
        values,
        values.map((_, index) => index),
        input[2]!,
        context,
      ),
    ]);
  return output;
}

export const dynamicArrayFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  [
    "DROP",
    ([value, rows, columns]) => {
      let output = matrix(value);
      if (rows !== undefined) output = takeDrop(output, scalarNumber(rows), false, false);
      if (columns !== undefined) output = takeDrop(output, scalarNumber(columns), false, true);
      return output;
    },
  ],
  [
    "EXPAND",
    ([value, rows, columns, pad]) => {
      const source = matrix(value);
      const targetRows = rows === undefined ? height(source) : arrayLength(rows, height(source));
      const targetColumns =
        columns === undefined ? width(source) : arrayLength(columns, width(source));
      assertFormulaCellCount(targetRows, targetColumns);
      if (targetRows < height(source) || targetColumns < width(source))
        throw new FormulaValueError("#VALUE!");
      const filler = pad !== undefined && !isMatrix(pad) ? pad : formulaError("#N/A");
      return Array.from({ length: targetRows }, (_, row) =>
        Array.from({ length: targetColumns }, (_, column) => cell(source, row, column) ?? filler),
      );
    },
  ],
  [
    "FILTER",
    ([range, include, empty]) => {
      const source = matrix(range);
      const criteria = matrix(include);
      const vertical = height(source) === 1 && width(source) > 1;
      const expectedShape = vertical ? [1, width(source)] : [height(source), 1];
      const actualShape = [height(criteria), width(criteria)];
      if (actualShape[0] !== expectedShape[0] || actualShape[1] !== expectedShape[1])
        throw new FormulaValueError("#VALUE!");
      const selected = vertical
        ? source.map((row) =>
            row.filter((_, column) => scalarBoolean(criteria[0]![column] ?? false)),
          )
        : source.filter((_, row) => scalarBoolean(criteria[row]![0] ?? false));
      return selected.length ? selected : (empty ?? "#CALC!");
    },
  ],
  ["GROUPBY", groupBy],
  ["HSTACK", (args) => stack(args, true)],
  ["PIVOTBY", (args, context) => pivot(args, context)],
  [
    "RANDARRAY",
    ([rows, columns, minimum, maximum, whole], context) => {
      const rowCount = rows === undefined ? 1 : arrayLength(rows, 1);
      const columnCount = columns === undefined ? 1 : arrayLength(columns, 1);
      assertFormulaCellCount(rowCount, columnCount);
      const low = minimum === undefined ? 0 : scalarNumber(minimum);
      const high = maximum === undefined ? 1 : scalarNumber(maximum);
      const integer = whole !== undefined && scalarBoolean(whole);
      return shape(rowCount, columnCount).map((row) =>
        row.map(() =>
          integer
            ? Math.floor((context.workbook.random?.() ?? Math.random()) * (high - low + 1)) + low
            : (context.workbook.random?.() ?? Math.random()) * (high - low) + low,
        ),
      );
    },
  ],
  [
    "SEQUENCE",
    ([rows, columns, start, step]) => {
      const rowCount = arrayLength(rows, 1);
      const columnCount = columns === undefined ? 1 : arrayLength(columns, 1);
      assertFormulaCellCount(rowCount, columnCount);
      const first = start === undefined ? 1 : scalarNumber(start);
      const increment = step === undefined ? 1 : scalarNumber(step);
      let serial = first;
      return Array.from({ length: rowCount }, () =>
        Array.from({ length: columnCount }, () => {
          const value = serial;
          serial += increment;
          return value;
        }),
      );
    },
  ],
  [
    "SORTBY",
    (args) => {
      const [range] = args;
      const source = matrix(range);
      const byColumn = height(source) === 1 && width(source) > 1;
      const total = byColumn ? width(source) : height(source);
      let indexes = Array.from({ length: total }, (_, index) => index);
      for (let argument = 1; argument < args.length; argument += 2) {
        const byArray = args[argument] ?? range;
        const order = args[argument + 1];
        const criteria = numbers(byArray);
        if (criteria.length !== total) throw new FormulaValueError("#VALUE!");
        const direction = order === undefined || order === null ? 1 : scalarNumber(order);
        if (direction !== 1 && direction !== -1) throw new FormulaValueError("#VALUE!");
        indexes = [...indexes].sort(
          (left, right) => (criteria[left]! - criteria[right]!) * direction || left - right,
        );
      }
      return byColumn
        ? [indexes.map((index) => source[0]?.[index] ?? null)]
        : indexes.map((index) => source[index] ?? []);
    },
  ],
  [
    "TAKE",
    ([value, rows, columns]) => {
      let output = matrix(value);
      if (rows !== undefined) output = takeDrop(output, scalarNumber(rows), true, false);
      if (columns !== undefined) output = takeDrop(output, scalarNumber(columns), true, true);
      return output;
    },
  ],
  [
    "TOCOL",
    (args) => {
      if (args.length > 3) throw new FormulaValueError("#VALUE!");
      const ignore = args[1] === undefined ? 0 : scalarNumber(args[1]);
      if (![0, 1, 2, 3].includes(ignore)) throw new FormulaValueError("#VALUE!");
      const byColumn = args[2] !== undefined && scalarBoolean(args[2]);
      const entries = flattenIgnoring(args[0]!, ignore, byColumn);
      return entries.map((entry) => [entry]);
    },
  ],
  [
    "TOROW",
    (args) => {
      if (args.length > 3) throw new FormulaValueError("#VALUE!");
      const ignore = args[1] === undefined ? 0 : scalarNumber(args[1]);
      if (![0, 1, 2, 3].includes(ignore)) throw new FormulaValueError("#VALUE!");
      const byColumn = args[2] !== undefined && scalarBoolean(args[2]);
      const entries = flattenIgnoring(args[0]!, ignore, byColumn);
      return [entries];
    },
  ],
  [
    "TRIMRANGE",
    (args) => {
      if (args.length > 3) throw new FormulaValueError("#VALUE!");
      const source = matrix(args[0] ?? null);
      const rowMode = args[1] === undefined ? 3 : scalarNumber(args[1]);
      const columnMode = args[2] === undefined ? 3 : scalarNumber(args[2]);
      if (![0, 1, 2, 3].includes(rowMode) || ![0, 1, 2, 3].includes(columnMode))
        throw new FormulaValueError("#VALUE!");
      return trimMatrix(source, rowMode, columnMode);
    },
  ],
  [
    "UNIQUE",
    (args) => {
      if (args.length > 3) throw new FormulaValueError("#VALUE!");
      const source = matrix(args[0] ?? null);
      const byColumn = args[1];
      const exactlyOnce = args[2];
      if (args.length > 3) throw new FormulaValueError("#VALUE!");
      const horizontal = byColumn !== undefined && scalarBoolean(byColumn);
      const unique = uniqueRows(
        horizontal ? transpose(source) : source,
        exactlyOnce !== undefined && scalarBoolean(exactlyOnce),
      );
      if (unique.length === 0) return "#CALC!";
      return horizontal ? transpose(unique) : unique;
    },
  ],
  ["VSTACK", (args) => stack(args, false)],
  [
    "WRAPCOLS",
    ([value, count, pad]) =>
      wrap(
        value,
        arrayLength(count, 1),
        true,
        pad !== undefined && !isMatrix(pad) ? pad : formulaError("#N/A"),
      ),
  ],
  [
    "WRAPROWS",
    ([value, count, pad]) =>
      wrap(
        value,
        arrayLength(count, 1),
        false,
        pad !== undefined && !isMatrix(pad) ? pad : formulaError("#N/A"),
      ),
  ],
  [
    "XLOOKUP",
    (args) => {
      if (args.length > 6) throw new FormulaValueError("#VALUE!");
      const [needle, lookup, returns, missing] = args;
      const { keys, results } = lookupArguments(lookup!, returns!);
      const mode = lookupMode(args[4], [-1, 0, 1, 2], 0);
      const search = lookupMode(args[5], [-2, -1, 1, 2], 1);
      if (mode === undefined || search === undefined) throw new FormulaValueError("#VALUE!");
      const target = scalarCriterion(needle ?? null);
      const index = lookupIndex(keys, target, mode, search);
      if (index >= 0) return results[index] ?? null;
      return missing ?? "#N/A";
    },
  ],
  [
    "XMATCH",
    (args) => {
      if (args.length > 4) throw new FormulaValueError("#VALUE!");
      const mode = lookupMode(args[2], [-1, 0, 1, 2], 0);
      const search = lookupMode(args[3], [-2, -1, 1, 2], 1);
      if (mode === undefined || search === undefined) throw new FormulaValueError("#VALUE!");
      const { keys } = lookupArguments(args[1]!, args[1]!);
      const index = lookupIndex(keys, scalarCriterion(args[0] ?? null), mode, search);
      return index >= 0 ? index + 1 : "#N/A";
    },
  ],
]);
