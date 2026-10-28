import { FormulaValueError } from "../errors";
import type { FormulaMatrix, FormulaScalar, FormulaValue } from "../value";
import { isMatrix } from "../value";
import { scalarNumber } from "./utils";
import { numericSeries } from "./utils";

export interface LinearRegressionStatistics {
  coefficients: number[];
  includeConstant: boolean;
  standardErrors: number[];
  rSquared: number;
  standardErrorY: number;
  fStatistic: number;
  degreesOfFreedom: number;
  regressionSumSquares: number;
  residualSumSquares: number;
}

interface QrSolution {
  values: number[];
  pivotColumns: number[];
}

interface QrCovarianceSolution extends QrSolution {
  covariance: number[][];
}

interface QrFactorization extends QrSolution {
  rank: number;
  reduced: number[][];
}

export function observationPredictors(
  value: FormulaValue,
  observationCount: number,
): number[][] | undefined {
  if (!isMatrix(value))
    return Array.from({ length: observationCount }, () => [scalarNumber(value)]);

  const rows: Array<number[] | undefined> = [];
  for (const row of value) rows.push(numericRow(row ?? []));
  if (rows.some((row) => row === undefined)) return undefined;
  const resolvedRows = rows.filter((row): row is number[] => row !== undefined);

  if (resolvedRows.length === observationCount) return rectangularRows(resolvedRows);
  if (resolvedRows.length === 1) {
    const [row] = resolvedRows;
    if (row?.length === observationCount) return row.map((entry) => [entry]);
    if (observationCount === 1) return [row ?? []];
  }
  return undefined;
}

export function regressionObservations(
  responseValue: FormulaValue,
  predictorValue: FormulaValue,
): { response: number[]; predictors: number[][] } | undefined {
  const response = numericSeries(responseValue);
  const allPredictors = optionalObservationPredictors(predictorValue, response.length);
  if (allPredictors === undefined) return undefined;

  const pairedResponse: number[] = [];
  const pairedPredictors: number[][] = [];
  for (const [index, value] of response.entries()) {
    const predictors = allPredictors[index];
    if (value !== undefined && predictors !== undefined) {
      pairedResponse.push(value);
      pairedPredictors.push(predictors);
    }
  }
  return pairedResponse.length === 0
    ? undefined
    : { response: pairedResponse, predictors: pairedPredictors };
}

function optionalObservationPredictors(
  value: FormulaValue,
  observationCount: number,
): Array<number[] | undefined> | undefined {
  if (!isMatrix(value))
    return Array.from({ length: observationCount }, () => [scalarNumber(value)]);

  const rows = value.map((row) => optionalNumericRow(row ?? []));
  if (rows.length === observationCount) return rectangularOptionalRows(rows);
  if (rows.length === 1) {
    const [row] = rows;
    if (row?.length === observationCount)
      return row.map((entry) => (entry === undefined ? undefined : [entry]));
    if (observationCount === 1 && row?.every((entry) => entry !== undefined)) return [row];
  }
  return undefined;
}

export function predictionRows(
  value: FormulaValue,
  predictorCount: number,
): number[][] | undefined {
  if (!isMatrix(value)) {
    if (predictorCount !== 1) return undefined;
    return [[scalarNumber(value)]];
  }

  const rows: Array<number[] | undefined> = [];
  for (const row of value) rows.push(numericRow(row ?? []));
  if (rows.some((row) => row === undefined)) return undefined;
  const resolvedRows = rows.filter((row): row is number[] => row !== undefined);

  if (resolvedRows.length === 1) {
    const [row] = resolvedRows;
    if (row === undefined || row.length !== predictorCount) return undefined;
    return predictorCount === 1 ? row.map((entry) => [entry]) : [row];
  }
  return rectangularRows(resolvedRows)?.every((row) => row.length === predictorCount)
    ? resolvedRows
    : undefined;
}

export function fitLinearRegression(
  response: readonly number[],
  predictors: readonly number[][],
  includeConstant: boolean,
  returnStatistics = false,
): LinearRegressionStatistics {
  const columns = predictors[0]?.length ?? 0;
  if (
    columns === 0 ||
    predictors.length !== response.length ||
    predictors.some((row) => row.length !== columns) ||
    response.some((value) => !Number.isFinite(value))
  )
    throw new FormulaValueError("#REF!");

  const columnMeans = Array.from(
    { length: columns },
    (_, column) =>
      predictors.reduce((total, row) => total + (row[column] ?? 0), 0) / response.length,
  );
  const responseMean = response.reduce((total, value) => total + value, 0) / response.length;
  const centeredPredictors = predictors.map((row) =>
    row.map((value, column) => (includeConstant ? value - (columnMeans[column] ?? 0) : value)),
  );
  const centeredResponse = includeConstant
    ? response.map((value) => value - responseMean)
    : [...response];
  let solution: QrSolution;
  let covariance: number[][];
  if (returnStatistics) {
    const solved = solveWithQrCovariance(centeredPredictors, centeredResponse);
    solution = solved;
    covariance = solved.covariance;
  } else {
    solution = solveWithQr(centeredPredictors, centeredResponse);
    covariance = [];
  }
  const slopes = solution.values;
  const intercept = includeConstant
    ? responseMean -
      slopes.reduce((total, slope, column) => total + slope * (columnMeans[column] ?? 0), 0)
    : 0;
  const coefficients = includeConstant ? [intercept, ...slopes] : slopes;
  const rank = solution.pivotColumns.length;
  if (rank === 0 && !includeConstant) throw new FormulaValueError("#NUM!");
  const fitted = predictors.map((row) =>
    slopes.reduce(
      (total, slope, column) => total + slope * (row[column] ?? 0),
      includeConstant ? intercept : 0,
    ),
  );
  const residualSumSquares = sumSquares(response, fitted);
  const meanResponse = response.reduce((total, value) => total + value, 0) / response.length;
  const totalSumSquares = includeConstant
    ? response.reduce((total, value) => total + (value - meanResponse) ** 2, 0)
    : response.reduce((total, value) => total + value ** 2, 0);
  const regressionSumSquares = Math.max(0, totalSumSquares - residualSumSquares);
  const degreesOfFreedom = response.length - rank - (includeConstant ? 1 : 0);
  const errorVariance = degreesOfFreedom === 0 ? 0 : residualSumSquares / degreesOfFreedom;
  const standardErrors = Array.from({ length: columns + (includeConstant ? 1 : 0) }, () => 0);

  if (returnStatistics) {
    solution.pivotColumns.forEach((column, activeIndex) => {
      const coefficientIndex = includeConstant ? column + 1 : column;
      standardErrors[coefficientIndex] = Math.sqrt(
        Math.max(0, covariance[activeIndex]?.[activeIndex] ?? 0) * errorVariance,
      );
    });
    if (includeConstant) {
      const meanQuadratic = solution.pivotColumns.reduce(
        (total, column, row) =>
          total +
          (columnMeans[column] ?? 0) *
            covariance[row]!.reduce(
              (rowTotal, covariance, covarianceColumn) =>
                rowTotal +
                covariance * (columnMeans[solution.pivotColumns[covarianceColumn]!] ?? 0),
              0,
            ),
        0,
      );
      standardErrors[0] = Math.sqrt(
        Math.max(0, 1 / response.length + meanQuadratic) * errorVariance,
      );
    }
  }

  return {
    coefficients,
    includeConstant,
    standardErrors,
    rSquared: totalSumSquares === 0 ? 1 : Math.max(0, 1 - residualSumSquares / totalSumSquares),
    standardErrorY: Math.sqrt(errorVariance),
    fStatistic:
      residualSumSquares === 0
        ? Number.POSITIVE_INFINITY
        : regressionSumSquares / Math.max(1, rank) / (errorVariance || 1),
    degreesOfFreedom,
    regressionSumSquares,
    residualSumSquares,
  };
}

export function predictLinearValues(
  statistics: LinearRegressionStatistics,
  predictors: readonly number[][],
  exponentiate = false,
): FormulaMatrix {
  const { coefficients, includeConstant } = statistics;
  const intercept = includeConstant ? (coefficients[0] ?? 0) : 0;
  const slopes = includeConstant ? coefficients.slice(1) : coefficients;
  return predictors.map((row) => {
    const fitted = slopes.reduce(
      (total, coefficient, index) => total + coefficient * (row[index] ?? 0),
      intercept,
    );
    const value = exponentiate ? Math.exp(fitted) : fitted;
    if (!Number.isFinite(value)) throw new FormulaValueError("#NUM!");
    return [value];
  });
}

function numericRow(row: readonly FormulaScalar[]): number[] | undefined {
  const values: number[] = [];
  for (const value of row) {
    if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
    values.push(value);
  }
  return values;
}

function optionalNumericRow(row: readonly FormulaScalar[]): Array<number | undefined> {
  return row.map((value) => {
    if (value === null) return undefined;
    if (typeof value === "string" && (value.trim() === "" || !Number.isFinite(Number(value))))
      return undefined;
    return scalarNumber(value);
  });
}

function rectangularOptionalRows(
  rows: ReadonlyArray<Array<number | undefined> | undefined>,
): Array<number[] | undefined> | undefined {
  const columns = rows[0]?.length ?? 0;
  return rows.every((row): row is number[] => row !== undefined && row.length === columns)
    ? [...rows]
    : undefined;
}

function rectangularRows(rows: readonly number[][]): number[][] | undefined {
  const columns = rows[0]?.length ?? 0;
  return rows.every((row) => row.length === columns) ? [...rows] : undefined;
}

function sumSquares(observed: readonly number[], fitted: readonly number[]): number {
  return observed.reduce((total, value, index) => total + (value - (fitted[index] ?? 0)) ** 2, 0);
}

function euclideanNorm(values: readonly number[]): number {
  const scale = values.reduce((maximum, value) => Math.max(maximum, Math.abs(value)), 0);
  if (scale === 0) return 0;
  return Math.sqrt(values.reduce((total, value) => total + (value / scale) ** 2, 0)) * scale;
}

function solveWithQr(
  design: readonly (readonly number[])[],
  response: readonly number[],
): QrSolution {
  const { values, pivotColumns } = solveQr(design, response);
  return { values, pivotColumns };
}

function solveWithQrCovariance(
  design: readonly (readonly number[])[],
  response: readonly number[],
): QrCovarianceSolution {
  const { values, pivotColumns, rank, reduced } = solveQr(design, response);
  const inverseR = Array.from({ length: rank }, () => Array.from({ length: rank }, () => 0));
  for (let target = 0; target < rank; target += 1) {
    for (let activeIndex = rank - 1; activeIndex >= 0; activeIndex -= 1) {
      let residual = activeIndex === target ? 1 : 0;
      for (let otherIndex = activeIndex + 1; otherIndex < rank; otherIndex += 1)
        residual -=
          reduced[activeIndex]![pivotColumns[otherIndex]!]! * inverseR[otherIndex]![target]!;
      inverseR[activeIndex]![target] =
        residual / reduced[activeIndex]![pivotColumns[activeIndex]!]!;
    }
  }
  const covariance = Array.from({ length: rank }, (_, left) =>
    Array.from({ length: rank }, (_, right) =>
      inverseR[left]!.reduce((total, value, index) => total + value * inverseR[right]![index], 0),
    ),
  );
  if (!covariance.flat().every(Number.isFinite)) throw new FormulaValueError("#NUM!");
  return { values, pivotColumns, covariance };
}

function solveQr(
  design: readonly (readonly number[])[],
  response: readonly number[],
): QrFactorization {
  const rows = design.length;
  const columns = design[0]?.length ?? 0;
  const reduced = design.map((row) => [...row]);
  const transformed = [...response];
  const columnScales = Array.from({ length: columns }, (_, column) =>
    euclideanNorm(design.map((row) => row[column] ?? 0)),
  );
  const scale = Math.max(0, ...columnScales);
  const tolerance = Number.EPSILON * scale * Math.max(1, rows, columns);
  const pivotColumns: number[] = [];

  for (let column = 0; column < Math.min(rows, columns); column += 1) {
    const pivotRow = pivotColumns.length;
    const norm = euclideanNorm(
      reduced.flatMap((row, index) => (index >= pivotRow ? [row[column]!] : [])),
    );
    if (norm <= tolerance) continue;

    const reflectorScale = reduced.reduce(
      (maximum, row, index) =>
        index >= pivotRow ? Math.max(maximum, Math.abs(row[column] ?? 0)) : maximum,
      0,
    );
    const reflector = reduced.map((row, index) =>
      index >= pivotRow ? (row[column] ?? 0) / reflectorScale : 0,
    );
    const diagonal = reflector[pivotRow]!;
    const signedNorm = diagonal >= 0 ? -norm : norm;
    reflector[pivotRow] = diagonal - signedNorm / reflectorScale;
    const divisor = reflector.reduce((total, value) => total + value ** 2, 0);

    for (let right = column + 1; right < columns; right += 1) {
      const projection = reduced.reduce(
        (total, row, index) => total + (index >= pivotRow ? row[right]! * reflector[index]! : 0),
        0,
      );
      const factor = (2 * projection) / divisor;
      reduced.forEach((row, index) => {
        if (index >= pivotRow) row[right]! -= factor * reflector[index]!;
      });
    }
    const responseProjection = transformed.reduce(
      (total, value, index) => total + (index >= pivotRow ? value * reflector[index]! : 0),
      0,
    );
    const responseFactor = (2 * responseProjection) / divisor;
    transformed.forEach((value, index) => {
      if (index >= pivotRow) transformed[index] = value - responseFactor * reflector[index]!;
    });

    reduced[pivotRow]![column] = signedNorm;
    for (let index = pivotRow + 1; index < rows; index += 1) reduced[index]![column] = 0;
    pivotColumns.push(column);
    if (pivotColumns.length === columns) break;
  }

  const values = Array.from({ length: columns }, () => 0);
  const rank = pivotColumns.length;
  for (let activeIndex = rank - 1; activeIndex >= 0; activeIndex -= 1) {
    const column = pivotColumns[activeIndex]!;
    let residual = transformed[activeIndex]!;
    for (let otherIndex = activeIndex + 1; otherIndex < rank; otherIndex += 1)
      residual -=
        reduced[activeIndex]![pivotColumns[otherIndex]!]! * values[pivotColumns[otherIndex]!]!;
    values[column] = residual / reduced[activeIndex]![column]!;
  }

  if (!values.every(Number.isFinite)) throw new FormulaValueError("#NUM!");
  return { values, pivotColumns, rank, reduced };
}
