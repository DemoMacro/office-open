import { FormulaValueError } from "../errors";
import type { FormulaScalar } from "../value";
import { statisticalFunctions } from "./statistical";
import {
  betaFunction,
  combination,
  factorial,
  gamma,
  gammaLn,
  incompleteBeta,
  inverseIncompleteBeta,
  lowerGamma,
  normalCdf,
  normalInverse,
  requireNonNegative,
  requirePositive,
  requireProbability,
  sumSeries,
  studentTailProbability,
} from "./statistical-numeric";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { numbers, scalarBoolean, scalarNumber } from "./utils";

type FormulaFunction = FormulaFunctionImplementation;
type FunctionArguments = Parameters<FormulaFunctionImplementation>[0];

function existingFunction(name: string): FormulaFunction {
  const implementation = statisticalFunctions.get(name);
  if (!implementation) throw new Error(`Statistical function is unavailable: ${name}`);
  return implementation;
}

function alias(source: string, target: string): [string, FormulaFunction] {
  const implementation = existingFunction(target);
  return [source, (args, context) => implementation(args, context)];
}

function aliases(entries: Array<[string, string]>): Array<[string, FormulaFunction]> {
  return entries.map(([source, target]) => alias(source, target));
}

function flatNumbers(args: FunctionArguments): number[] {
  return args.flatMap(numbers);
}

function pairedNumbers(
  left: FunctionArguments[number],
  right: FunctionArguments[number],
): {
  x: number[];
  y: number[];
} {
  const x = numbers(left);
  const y = numbers(right);
  const length = Math.min(x.length, y.length);
  return { x: x.slice(0, length), y: y.slice(0, length) };
}

function matrixRows(value: FunctionArguments[number]): FormulaScalar[][] {
  return Array.isArray(value) ? value : [[value]];
}

function columnCount(rows: FormulaScalar[][]): number {
  return Math.max(1, ...rows.map((row) => row.length));
}

function mean(input: number[]): number {
  return input.reduce((sum, value) => sum + value, 0) / input.length;
}

function variance(input: number[], sample: boolean): number {
  if (input.length < (sample ? 2 : 1)) throw new FormulaValueError("#DIV/0!");
  const center = mean(input);
  return (
    input.reduce((sum, value) => sum + (value - center) ** 2, 0) / (input.length - (sample ? 1 : 0))
  );
}

function betaDistribution(
  x: number,
  alpha: number,
  beta: number,
  lower = 0,
  upper = 1,
  cumulative = true,
): number {
  if (lower >= upper || x < lower || x > upper) throw new FormulaValueError("#NUM!");
  const standardized = requireProbability((x - lower) / (upper - lower));
  return cumulative
    ? incompleteBeta(standardized, alpha, beta)
    : (standardized ** (alpha - 1) * (1 - standardized) ** (beta - 1)) /
        ((upper - lower) * betaFunction(alpha, beta));
}

function betaInverse(
  probability: number,
  alpha: number,
  beta: number,
  lower = 0,
  upper = 1,
): number {
  if (lower >= upper) throw new FormulaValueError("#NUM!");
  const standardized = inverseIncompleteBeta(requireProbability(probability), alpha, beta);
  return lower + standardized * (upper - lower);
}

function binomialMass(successes: number, trials: number, probability: number): number {
  const trialsInteger = Math.trunc(trials);
  const successesInteger = Math.trunc(successes);
  if (successesInteger < 0 || successesInteger > trialsInteger) return 0;
  return (
    combination(trialsInteger, successesInteger) *
    probability ** successesInteger *
    (1 - probability) ** (trialsInteger - successesInteger)
  );
}

function binomialDistribution(
  successes: number,
  trials: number,
  probability: number,
  cumulative: boolean,
): number {
  requireProbability(probability);
  const trialsInteger = Math.trunc(trials);
  if (trialsInteger < 0 || trials !== trialsInteger) throw new FormulaValueError("#NUM!");
  const successesInteger = Math.trunc(successes);
  if (successesInteger < 0) throw new FormulaValueError("#NUM!");
  return cumulative
    ? sumSeries(0, successesInteger, (index) => binomialMass(index, trialsInteger, probability))
    : binomialMass(successesInteger, trialsInteger, probability);
}

function negativeBinomialDistribution(
  failures: number,
  successes: number,
  probability: number,
  cumulative: boolean,
): number {
  requireProbability(probability);
  const successesInteger = Math.trunc(successes);
  const failuresInteger = Math.trunc(failures);
  if (successesInteger < 1 || failuresInteger < 0) throw new FormulaValueError("#NUM!");
  const mass = (index: number) =>
    combination(index + successesInteger - 1, index) *
    probability ** successesInteger *
    (1 - probability) ** index;
  return cumulative ? sumSeries(0, failuresInteger, mass) : mass(failuresInteger);
}

function chiSquareDistribution(value: number, degrees: number, cumulative: boolean): number {
  requireNonNegative(value);
  if (degrees <= 0) throw new FormulaValueError("#NUM!");
  if (!cumulative)
    return (
      (Math.exp(-value / 2) * value ** (degrees / 2 - 1)) /
      (2 ** (degrees / 2) * gamma(degrees / 2))
    );
  return lowerGamma(degrees / 2, value / 2) / gamma(degrees / 2);
}

function chiSquareInverse(probability: number, degrees: number): number {
  const target = requireProbability(probability);
  if (degrees <= 0) throw new FormulaValueError("#NUM!");
  let low = 0;
  let high = Math.max(degrees * 5, 10);
  for (let index = 0; index < 120; index += 1) {
    const middle = (low + high) / 2;
    if (chiSquareDistribution(middle, degrees, true) < target) low = middle;
    else high = middle;
  }
  return (low + high) / 2;
}

function fDistribution(
  value: number,
  numeratorDegrees: number,
  denominatorDegrees: number,
  cumulative: boolean,
): number {
  if (value < 0 || numeratorDegrees <= 0 || denominatorDegrees <= 0)
    throw new FormulaValueError("#NUM!");
  const transformed = (numeratorDegrees * value) / (numeratorDegrees * value + denominatorDegrees);
  const lower = incompleteBeta(transformed, numeratorDegrees / 2, denominatorDegrees / 2);
  const densityScale = numeratorDegrees / denominatorDegrees;
  if (!cumulative) {
    return (
      (densityScale ** (numeratorDegrees / 2) *
        value ** (numeratorDegrees / 2 - 1) *
        (1 + densityScale * value) ** (-(numeratorDegrees + denominatorDegrees) / 2)) /
      betaFunction(numeratorDegrees / 2, denominatorDegrees / 2)
    );
  }
  return lower;
}

function tDistribution(
  value: number,
  degrees: number,
  cumulative: boolean,
  twoTailed = false,
): number {
  if (degrees <= 0 || !Number.isFinite(value)) throw new FormulaValueError("#NUM!");
  if (!cumulative)
    return (
      (gamma((degrees + 1) / 2) / (Math.sqrt(degrees * Math.PI) * gamma(degrees / 2))) *
      (1 + (value * value) / degrees) ** (-(degrees + 1) / 2)
    );
  if (twoTailed) return studentTailProbability(value, degrees, 2);
  const halfTail = studentTailProbability(value, degrees, 1);
  return value >= 0 ? 1 - halfTail : halfTail;
}

function fInverseLeft(probability: number, numeratorDegrees: number, denominatorDegrees: number) {
  const transformed = inverseIncompleteBeta(
    requireProbability(probability),
    numeratorDegrees / 2,
    denominatorDegrees / 2,
  );
  return (denominatorDegrees * transformed) / (numeratorDegrees * (1 - transformed));
}

function fInverseRight(probability: number, numeratorDegrees: number, denominatorDegrees: number) {
  return fInverseLeft(1 - requireProbability(probability), numeratorDegrees, denominatorDegrees);
}

function tInverseLeft(probability: number, degrees: number): number {
  const target = requireProbability(probability);
  if (degrees <= 0) throw new FormulaValueError("#NUM!");
  if (target <= 0) throw new FormulaValueError("#NUM!");
  let low = -1;
  let high = 1;
  while (target < 0.5 && tDistribution(low, degrees, true) > target) low *= 2;
  while (target > 0.5 && tDistribution(high, degrees, true) < target) high *= 2;
  for (let index = 0; index < 120; index += 1) {
    const middle = (low + high) / 2;
    if (tDistribution(middle, degrees, true) < target) low = middle;
    else high = middle;
  }
  return (low + high) / 2;
}

function tInverseTwoTailed(probability: number, degrees: number): number {
  const target = requireProbability(probability);
  if (degrees <= 0) throw new FormulaValueError("#NUM!");
  if (target <= 0) throw new FormulaValueError("#NUM!");
  let low = 0;
  let high = 1;
  while (tDistribution(high, degrees, true, true) > target) high *= 2;
  for (let index = 0; index < 120; index += 1) {
    const middle = (low + high) / 2;
    if (tDistribution(middle, degrees, true, true) > target) low = middle;
    else high = middle;
  }
  return (low + high) / 2;
}

function gammaDistribution(x: number, alpha: number, beta: number, cumulative: boolean): number {
  requireNonNegative(x);
  if (alpha <= 0 || beta <= 0) throw new FormulaValueError("#NUM!");
  const standardized = x / beta;
  if (!cumulative)
    return (Math.exp(-standardized) * standardized ** (alpha - 1)) / (beta * gamma(alpha));
  return lowerGamma(alpha, standardized) / gamma(alpha);
}

function gammaInverse(probability: number, alpha: number, beta: number): number {
  const target = requireProbability(probability);
  if (alpha <= 0 || beta <= 0) throw new FormulaValueError("#NUM!");
  let low = 0;
  let high = alpha * beta + 10 * Math.sqrt(alpha) * beta + 10;
  for (let index = 0; index < 120; index += 1) {
    const middle = (low + high) / 2;
    if (gammaDistribution(middle, alpha, beta, true) < target) low = middle;
    else high = middle;
  }
  return (low + high) / 2;
}

function logNormalDistribution(
  x: number,
  meanValue: number,
  standardDeviation: number,
  cumulative: boolean,
): number {
  if (x <= 0 || standardDeviation <= 0) throw new FormulaValueError("#NUM!");
  const z = (Math.log(x) - meanValue) / standardDeviation;
  if (!cumulative) return Math.exp((-z * z) / 2) / (x * standardDeviation * Math.sqrt(2 * Math.PI));
  return normalCdf(z);
}

function poissonDistribution(value: number, expected: number, cumulative: boolean): number {
  requireNonNegative(value);
  requireNonNegative(expected);
  const integer = Math.trunc(value);
  if (!cumulative) return (Math.exp(-expected) * expected ** integer) / factorial(integer);
  return sumSeries(
    0,
    integer,
    (index) => (Math.exp(-expected) * expected ** index) / factorial(index),
  );
}

function weibullDistribution(x: number, shape: number, scale: number, cumulative: boolean): number {
  if (x < 0 || shape <= 0 || scale <= 0) throw new FormulaValueError("#NUM!");
  const transformed = (x / scale) ** shape;
  return cumulative
    ? 1 - Math.exp(-transformed)
    : (shape / scale ** shape) * x ** (shape - 1) * Math.exp(-transformed);
}

function percentileValue(input: number[], probability: number, inclusive: boolean): number {
  if (input.length === 0) throw new FormulaValueError("#NUM!");
  requireProbability(probability);
  const sorted = [...input].sort((left, right) => left - right);
  const position = inclusive
    ? probability * (sorted.length - 1)
    : probability * (sorted.length + 1) - 1;
  if (position < 0 || position > sorted.length - 1) throw new FormulaValueError("#NUM!");
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  return sorted[lower]! + (position - lower) * (sorted[upper]! - sorted[lower]!);
}

function percentileRank(
  input: number[],
  value: number,
  significance: number | undefined,
  inclusive: boolean,
): number {
  if (input.length === 0 || (inclusive && input.length === 1)) throw new FormulaValueError("#NUM!");
  const sorted = [...input].sort((left, right) => left - right);
  const lastIndex = sorted.length - 1;
  const exactIndex = sorted.indexOf(value);
  const lowerIndex = sorted.findLastIndex((entry) => entry < value);
  const upperIndex = sorted.findIndex((entry) => entry > value);
  const raw = inclusive
    ? exactIndex !== -1
      ? exactIndex / lastIndex
      : lowerIndex === -1 || upperIndex === -1
        ? Number.NaN
        : lowerIndex / lastIndex +
          ((value - sorted[lowerIndex]!) / (sorted[upperIndex]! - sorted[lowerIndex]!)) *
            ((upperIndex - lowerIndex) / lastIndex)
    : exactIndex !== -1
      ? (exactIndex + 1) / (sorted.length + 1)
      : lowerIndex === -1 || upperIndex === -1
        ? Number.NaN
        : (lowerIndex + 1) / (sorted.length + 1) +
          ((value - sorted[lowerIndex]!) / (sorted[upperIndex]! - sorted[lowerIndex]!)) *
            ((upperIndex - lowerIndex) / (sorted.length + 1));
  if (!Number.isFinite(raw)) throw new FormulaValueError("#N/A");
  if (significance !== undefined && (significance <= 0 || !Number.isInteger(significance)))
    throw new FormulaValueError("#NUM!");
  const digits = significance ?? 3;
  return Number(raw.toFixed(digits));
}

function quartileValue(input: number[], quart: number, inclusive: boolean): number {
  if (quart < 0 || quart > 4 || !Number.isInteger(quart)) throw new FormulaValueError("#NUM!");
  return percentileValue(input, quart / 4, inclusive);
}

function rankValue(value: number, input: number[], descending: boolean, averageTies: boolean) {
  if (!input.includes(value)) throw new FormulaValueError("#N/A");
  const relevant = input.filter((entry) => (descending ? entry >= value : entry <= value));
  if (!averageTies) return relevant.length;
  const tied = input.filter((entry) => entry === value).length;
  return relevant.length - (tied - 1) / 2;
}

function pooledVariance(left: number[], right: number[]): number {
  const leftVariance = variance(left, true);
  const rightVariance = variance(right, true);
  return (
    ((left.length - 1) * leftVariance + (right.length - 1) * rightVariance) /
    (left.length + right.length - 2)
  );
}

function welchDegreesOfFreedom(left: number[], right: number[]): number {
  const leftVariance = variance(left, true) / left.length;
  const rightVariance = variance(right, true) / right.length;
  const denominator =
    leftVariance ** 2 / (left.length - 1) + rightVariance ** 2 / (right.length - 1);
  return (leftVariance + rightVariance) ** 2 / denominator;
}

function pairedDifferences(left: number[], right: number[]): number[] {
  if (left.length !== right.length) throw new FormulaValueError("#N/A");
  return left.map((value, index) => value - right[index]!);
}

function tTestProbability(left: number[], right: number[], tails: number, type: number): number {
  if (tails !== 1 && tails !== 2) throw new FormulaValueError("#NUM!");
  if (type < 1 || type > 3 || !Number.isInteger(type)) throw new FormulaValueError("#NUM!");
  let t: number;
  let degrees: number;
  if (type === 1) {
    const differences = pairedDifferences(left, right);
    t = mean(differences) / Math.sqrt(variance(differences, true) / differences.length);
    degrees = differences.length - 1;
  } else if (type === 2) {
    const pooled = pooledVariance(left, right);
    t = (mean(left) - mean(right)) / Math.sqrt(pooled * (1 / left.length + 1 / right.length));
    degrees = left.length + right.length - 2;
  } else {
    t =
      (mean(left) - mean(right)) /
      Math.sqrt(variance(left, true) / left.length + variance(right, true) / right.length);
    degrees = welchDegreesOfFreedom(left, right);
  }
  const probability = tDistribution(t, degrees, false, true);
  return tails === 1 ? probability / 2 : probability;
}

function fTestProbability(left: number[], right: number[]): number {
  const f = variance(left, true) / variance(right, true);
  const probability =
    2 *
    Math.min(
      fDistribution(f, left.length - 1, right.length - 1, true),
      1 - fDistribution(f, left.length - 1, right.length - 1, true),
    );
  return Math.min(1, probability);
}

function standardErrorY(x: number[], y: number[]): number {
  const centerX = mean(x);
  const centerY = mean(y);
  const slope =
    x.reduce((sum, value, index) => sum + (value - centerX) * (y[index]! - centerY), 0) /
    x.reduce((sum, value) => sum + (value - centerX) ** 2, 0);
  const intercept = centerY - slope * centerX;
  const squaredError = x.reduce(
    (sum, value, index) => sum + (y[index]! - (intercept + slope * value)) ** 2,
    0,
  );
  return Math.sqrt(squaredError / (x.length - 2));
}

function confidenceNormal(alpha: number, standardDeviation: number, size: number): number {
  if (alpha <= 0 || alpha >= 1 || standardDeviation <= 0 || size < 1)
    throw new FormulaValueError("#NUM!");
  return normalInverse(1 - alpha / 2) * (standardDeviation / Math.sqrt(size));
}

function confidenceStudent(alpha: number, standardDeviation: number, size: number): number {
  if (alpha <= 0 || alpha >= 1 || standardDeviation <= 0 || size < 2)
    throw new FormulaValueError("#NUM!");
  return tInverseTwoTailed(alpha, size - 1) * (standardDeviation / Math.sqrt(size));
}

function chiSquareTest(
  actual: FunctionArguments[number],
  expected: FunctionArguments[number],
): number {
  const actualRows = matrixRows(actual);
  const expectedRows = matrixRows(expected);
  const rows = Math.min(actualRows.length, expectedRows.length);
  const columns = Math.min(columnCount(actualRows), columnCount(expectedRows));
  let statistic = 0;
  for (let row = 0; row < rows; row += 1)
    for (let column = 0; column < columns; column += 1) {
      const observed = scalarNumber(actualRows[row]![column] ?? null);
      const probability = scalarNumber(expectedRows[row]![column] ?? null);
      if (probability <= 0) throw new FormulaValueError("#DIV/0!");
      statistic += (observed - probability) ** 2 / probability;
    }
  const degrees =
    rows === 1 || columns === 1 ? Math.max(1, rows * columns - 1) : (rows - 1) * (columns - 1);
  return chiSquareDistribution(statistic, degrees, false);
}

function cumulativeArgument(value: FunctionArguments[number]): boolean {
  return scalarBoolean(value);
}

function integerArgument(value: FunctionArguments[number]): number {
  const number = Math.trunc(scalarNumber(value));
  if (!Number.isFinite(number)) throw new FormulaValueError("#NUM!");
  return number;
}

function binomialRange(
  trials: FunctionArguments[number],
  probability: FunctionArguments[number],
  successes: FunctionArguments[number],
  successes2?: FunctionArguments[number],
): number {
  const count = requireNonNegative(Math.trunc(scalarNumber(trials)));
  const probabilityValue = requireProbability(scalarNumber(probability));
  const from = requireNonNegative(Math.trunc(scalarNumber(successes)));
  const to = successes2 === undefined ? from : Math.trunc(scalarNumber(successes2));
  return sumSeries(from, to, (index) => binomialMass(index, count, probabilityValue));
}

function binomialInverse(
  trials: FunctionArguments[number],
  probability: FunctionArguments[number],
  alpha: FunctionArguments[number],
): number {
  const count = Math.trunc(scalarNumber(trials));
  const probabilityValue = requireProbability(scalarNumber(probability));
  const target = requireProbability(scalarNumber(alpha));
  for (let index = 0; index <= count; index += 1)
    if (binomialDistribution(index, count, probabilityValue, true) >= target) return index;
  return count;
}

function exponentialDistribution(x: number, lambda: number, cumulative: boolean): number {
  requireNonNegative(x);
  const rate = requirePositive(lambda);
  return cumulative ? 1 - Math.exp(-rate * x) : rate * Math.exp(-rate * x);
}

function hypergeometricDistribution(
  successes: FunctionArguments[number],
  sampleSize: FunctionArguments[number],
  populationSuccesses: FunctionArguments[number],
  populationSize: FunctionArguments[number],
  cumulative?: FunctionArguments[number],
): number {
  const sampleSuccesses = Math.trunc(scalarNumber(successes));
  const sample = Math.trunc(scalarNumber(sampleSize));
  const populationSuccess = Math.trunc(scalarNumber(populationSuccesses));
  const population = Math.trunc(scalarNumber(populationSize));
  if (
    sampleSuccesses < 0 ||
    sample < 0 ||
    populationSuccess < 0 ||
    population < 0 ||
    sample > population ||
    populationSuccess > population
  )
    throw new FormulaValueError("#NUM!");
  const mass = (index: number) =>
    (combination(populationSuccess, index) *
      combination(population - populationSuccess, sample - index)) /
    combination(population, sample);
  const wantCumulative = cumulative === undefined ? false : cumulativeArgument(cumulative);
  return wantCumulative
    ? sumSeries(Math.max(0, sample + populationSuccess - population), sampleSuccesses, mass)
    : mass(sampleSuccesses);
}

function testSamples(
  left: FunctionArguments[number],
  right: FunctionArguments[number],
): [number[], number[]] {
  return [numbers(left), numbers(right)];
}

function zTest(
  input: FunctionArguments[number],
  value: FunctionArguments[number],
  standardDeviation?: FunctionArguments[number],
): number {
  const sample = flatNumbers([input]);
  const center = mean(sample);
  const sigma =
    standardDeviation === undefined
      ? Math.sqrt(variance(sample, true))
      : requirePositive(scalarNumber(standardDeviation));
  return 1 - normalCdf((center - scalarNumber(value)) / (sigma / Math.sqrt(sample.length)));
}

export const statisticalExtraFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  [
    "B",
    ([x, trials, probability, cumulative]) =>
      binomialDistribution(
        scalarNumber(x),
        scalarNumber(trials),
        scalarNumber(probability),
        cumulative === undefined ? false : cumulativeArgument(cumulative),
      ),
  ],
  [
    "BETADIST",
    ([x, alpha, beta, lower, upper]) =>
      betaDistribution(
        scalarNumber(x),
        requirePositive(scalarNumber(alpha)),
        requirePositive(scalarNumber(beta)),
        lower === undefined ? 0 : scalarNumber(lower),
        upper === undefined ? 1 : scalarNumber(upper),
      ),
  ],
  [
    "BETADISTCOMPAT",
    ([x, alpha, beta, lower, upper]) =>
      betaDistribution(
        scalarNumber(x),
        requirePositive(scalarNumber(alpha)),
        requirePositive(scalarNumber(beta)),
        lower === undefined ? 0 : scalarNumber(lower),
        upper === undefined ? 1 : scalarNumber(upper),
      ),
  ],
  [
    "BINOM.DIST",
    ([x, trials, probability, cumulative]) =>
      binomialDistribution(
        scalarNumber(x),
        scalarNumber(trials),
        scalarNumber(probability),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "BETAINV",
    ([probability, alpha, beta, lower, upper]) =>
      betaInverse(
        scalarNumber(probability),
        requirePositive(scalarNumber(alpha)),
        requirePositive(scalarNumber(beta)),
        lower === undefined ? 0 : scalarNumber(lower),
        upper === undefined ? 1 : scalarNumber(upper),
      ),
  ],
  [
    "CHIDIST",
    ([x, degrees]) => chiSquareDistribution(scalarNumber(x), scalarNumber(degrees), false),
  ],
  [
    "CHIINV",
    ([probability, degrees]) => chiSquareInverse(scalarNumber(probability), scalarNumber(degrees)),
  ],
  [
    "CHISQ.DIST",
    ([x, degrees, cumulative]) =>
      chiSquareDistribution(scalarNumber(x), scalarNumber(degrees), cumulativeArgument(cumulative)),
  ],
  [
    "BETAINVCOMPAT",
    ([probability, alpha, beta, lower, upper]) =>
      betaInverse(
        scalarNumber(probability),
        requirePositive(scalarNumber(alpha)),
        requirePositive(scalarNumber(beta)),
        lower === undefined ? 0 : scalarNumber(lower),
        upper === undefined ? 1 : scalarNumber(upper),
      ),
  ],
  [
    "CHISQ.INV",
    ([probability, degrees]) => chiSquareInverse(scalarNumber(probability), scalarNumber(degrees)),
  ],
  [
    "BINOM.DIST.RANGE",
    ([trials, probability, successes, successes2]) => {
      const count = requireNonNegative(Math.trunc(scalarNumber(trials)));
      const probabilityValue = requireProbability(scalarNumber(probability));
      const from = requireNonNegative(Math.trunc(scalarNumber(successes)));
      const to = successes2 === undefined ? from : Math.trunc(scalarNumber(successes2));
      return sumSeries(from, to, (index) => binomialMass(index, count, probabilityValue));
    },
  ],
  [
    "BINOM.INV",
    ([trials, probability, alpha]) => {
      const count = Math.trunc(scalarNumber(trials));
      const probabilityValue = requireProbability(scalarNumber(probability));
      const target = requireProbability(scalarNumber(alpha));
      for (let index = 0; index <= count; index += 1)
        if (binomialDistribution(index, count, probabilityValue, true) >= target) return index;
      return count;
    },
  ],
  [
    "BINOMDISTCOMPAT",
    ([x, trials, probability, cumulative]) =>
      binomialDistribution(
        scalarNumber(x),
        scalarNumber(trials),
        scalarNumber(probability),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "BINOMDISTRANGE",
    ([trials, probability, successes, successes2]) =>
      binomialRange(trials, probability, successes, successes2),
  ],
  ["BINOMINV", ([trials, probability, alpha]) => binomialInverse(trials, probability, alpha)],
  [
    "CHIDISTCOMPAT",
    ([x, degrees]) => chiSquareDistribution(scalarNumber(x), scalarNumber(degrees), false),
  ],
  [
    "CHIINVCOMPAT",
    ([probability, degrees]) => chiSquareInverse(scalarNumber(probability), scalarNumber(degrees)),
  ],
  [
    "CHISQ.DIST.RT",
    ([x, degrees]) => 1 - chiSquareDistribution(scalarNumber(x), scalarNumber(degrees), true),
  ],
  [
    "CHISQ.INV.RT",
    ([probability, degrees]) => chiSquareInverse(scalarNumber(probability), scalarNumber(degrees)),
  ],
  ["CHISQ.TEST", ([actual, expected]) => chiSquareTest(actual, expected)],
  [
    "CHISQDIST",
    ([x, degrees, cumulative]) =>
      chiSquareDistribution(scalarNumber(x), scalarNumber(degrees), cumulativeArgument(cumulative)),
  ],
  [
    "CHISQDISTRT",
    ([x, degrees]) => 1 - chiSquareDistribution(scalarNumber(x), scalarNumber(degrees), true),
  ],
  [
    "CHISQINV",
    ([probability, degrees]) => chiSquareInverse(scalarNumber(probability), scalarNumber(degrees)),
  ],
  [
    "CHISQINVRT",
    ([probability, degrees]) => chiSquareInverse(scalarNumber(probability), scalarNumber(degrees)),
  ],
  ["CHISQTEST", ([actual, expected]) => chiSquareTest(actual, expected)],
  ["CHITEST", ([actual, expected]) => chiSquareTest(actual, expected)],
  ["CHITESTCOMPAT", ([actual, expected]) => chiSquareTest(actual, expected)],
  [
    "CONFIDENCE",
    ([alpha, standardDeviation, size]) =>
      confidenceNormal(scalarNumber(alpha), scalarNumber(standardDeviation), scalarNumber(size)),
  ],
  [
    "CONFIDENCE.NORM",
    ([alpha, standardDeviation, size]) =>
      confidenceNormal(scalarNumber(alpha), scalarNumber(standardDeviation), scalarNumber(size)),
  ],
  [
    "CONFIDENCE.T",
    ([alpha, standardDeviation, size]) =>
      confidenceStudent(scalarNumber(alpha), scalarNumber(standardDeviation), scalarNumber(size)),
  ],
  [
    "CONFIDENCECOMPAT",
    ([alpha, standardDeviation, size]) =>
      confidenceNormal(scalarNumber(alpha), scalarNumber(standardDeviation), scalarNumber(size)),
  ],
  [
    "CONFIDENCENORM",
    ([alpha, standardDeviation, size]) =>
      confidenceNormal(scalarNumber(alpha), scalarNumber(standardDeviation), scalarNumber(size)),
  ],
  [
    "CONFIDENCET",
    ([alpha, standardDeviation, size]) =>
      confidenceStudent(scalarNumber(alpha), scalarNumber(standardDeviation), scalarNumber(size)),
  ],
  ...aliases([
    ["COVARCOMPAT", "COVAR"],
    ["COVARIANCEP", "COVARIANCE.P"],
    ["COVARIANCES", "COVARIANCE.S"],
  ]),
  ["CRITBINOM", ([trials, probability, alpha]) => binomialInverse(trials, probability, alpha)],
  [
    "CRITBINOMCOMPAT",
    ([trials, probability, alpha]) => binomialInverse(trials, probability, alpha),
  ],
  [
    "EXPON.DIST",
    ([x, lambda, cumulative]) =>
      exponentialDistribution(
        scalarNumber(x),
        scalarNumber(lambda),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "EXPONDISTCOMPAT",
    ([x, lambda, cumulative]) =>
      exponentialDistribution(
        scalarNumber(x),
        scalarNumber(lambda),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "F.DIST",
    ([x, numeratorDegrees, denominatorDegrees, cumulative]) =>
      fDistribution(
        scalarNumber(x),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "F.DIST.RT",
    ([x, numeratorDegrees, denominatorDegrees]) =>
      1 -
      fDistribution(
        scalarNumber(x),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
        true,
      ),
  ],
  [
    "F.INV.RT",
    ([probability, numeratorDegrees, denominatorDegrees]) =>
      fInverseRight(
        scalarNumber(probability),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
      ),
  ],
  [
    "F.INV",
    ([probability, numeratorDegrees, denominatorDegrees]) =>
      fInverseLeft(
        scalarNumber(probability),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
      ),
  ],
  ["F.TEST", ([left, right]) => fTestProbability(...testSamples(left, right))],
  [
    "FDIST",
    ([x, numeratorDegrees, denominatorDegrees]) =>
      1 -
      fDistribution(
        scalarNumber(x),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
        true,
      ),
  ],
  [
    "FDISTCOMPAT",
    ([x, numeratorDegrees, denominatorDegrees]) =>
      1 -
      fDistribution(
        scalarNumber(x),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
        true,
      ),
  ],
  [
    "FINV",
    ([probability, numeratorDegrees, denominatorDegrees]) =>
      fInverseRight(
        scalarNumber(probability),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
      ),
  ],
  [
    "FDISTRT",
    ([x, numeratorDegrees, denominatorDegrees]) =>
      1 -
      fDistribution(
        scalarNumber(x),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
        true,
      ),
  ],
  [
    "FINVCOMPAT",
    ([probability, numeratorDegrees, denominatorDegrees]) =>
      fInverseRight(
        scalarNumber(probability),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
      ),
  ],
  [
    "FINVRT",
    ([probability, numeratorDegrees, denominatorDegrees]) =>
      fInverseRight(
        scalarNumber(probability),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
      ),
  ],
  ["FTEST", ([left, right]) => fTestProbability(...testSamples(left, right))],
  ["FTESTCOMPAT", ([left, right]) => fTestProbability(...testSamples(left, right))],
  [
    "GAMMA.DIST",
    ([x, alpha, beta, cumulative]) =>
      gammaDistribution(
        scalarNumber(x),
        scalarNumber(alpha),
        scalarNumber(beta),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "GAMMADIST",
    ([x, alpha, beta, cumulative]) =>
      gammaDistribution(
        scalarNumber(x),
        scalarNumber(alpha),
        scalarNumber(beta),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "GAMMADISTCOMPAT",
    ([x, alpha, beta, cumulative]) =>
      gammaDistribution(
        scalarNumber(x),
        scalarNumber(alpha),
        scalarNumber(beta),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "GAMMA.INV",
    ([probability, alpha, beta]) =>
      gammaInverse(scalarNumber(probability), scalarNumber(alpha), scalarNumber(beta)),
  ],
  [
    "GAMMAINV",
    ([probability, alpha, beta]) =>
      gammaInverse(scalarNumber(probability), scalarNumber(alpha), scalarNumber(beta)),
  ],
  [
    "GAMMAINVCOMPAT",
    ([probability, alpha, beta]) =>
      gammaInverse(scalarNumber(probability), scalarNumber(alpha), scalarNumber(beta)),
  ],
  ["GAMMALN.PRECISE", ([x]) => gammaLn(scalarNumber(x))],
  ["GAMMALNPRECISE", ([x]) => gammaLn(scalarNumber(x))],
  ["GAMMA", ([x]) => gamma(scalarNumber(x))],
  [
    "HYPGEOMDIST",
    ([successes, sampleSize, populationSuccesses, populationSize, cumulative]) => {
      const sampleSuccesses = Math.trunc(scalarNumber(successes));
      const sample = Math.trunc(scalarNumber(sampleSize));
      const populationSuccess = Math.trunc(scalarNumber(populationSuccesses));
      const population = Math.trunc(scalarNumber(populationSize));
      if (
        sampleSuccesses < 0 ||
        sample < 0 ||
        populationSuccess < 0 ||
        population < 0 ||
        sample > population ||
        populationSuccess > population
      )
        throw new FormulaValueError("#NUM!");
      const mass = (index: number) =>
        (combination(populationSuccess, index) *
          combination(population - populationSuccess, sample - index)) /
        combination(population, sample);
      const wantCumulative = cumulative === undefined ? false : cumulativeArgument(cumulative);
      return wantCumulative
        ? sumSeries(Math.max(0, sample + populationSuccess - population), sampleSuccesses, mass)
        : mass(sampleSuccesses);
    },
  ],
  [
    "HYPGEOMDISTCOMPAT",
    ([successes, sampleSize, populationSuccesses, populationSize]) =>
      hypergeometricDistribution(successes, sampleSize, populationSuccesses, populationSize),
  ],
  [
    "LEGACYCHIDIST",
    ([x, degrees]) => chiSquareDistribution(scalarNumber(x), scalarNumber(degrees), false),
  ],
  [
    "LEGACYCHIINV",
    ([probability, degrees]) => chiSquareInverse(scalarNumber(probability), scalarNumber(degrees)),
  ],
  ["LEGACYCHITEST", ([actual, expected]) => chiSquareTest(actual, expected)],
  [
    "LEGACYFDIST",
    ([x, numeratorDegrees, denominatorDegrees]) =>
      1 -
      fDistribution(
        scalarNumber(x),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
        true,
      ),
  ],
  [
    "LEGACYFINV",
    ([probability, numeratorDegrees, denominatorDegrees]) =>
      fInverseRight(
        scalarNumber(probability),
        scalarNumber(numeratorDegrees),
        scalarNumber(denominatorDegrees),
      ),
  ],
  ["LEGACYNORMSDIST", ([x]) => normalCdf(scalarNumber(x))],
  ["LEGACYNORMSINV", ([probability]) => normalInverse(scalarNumber(probability))],
  [
    "LOGNORM.DIST",
    ([x, mean, standardDeviation, cumulative]) =>
      logNormalDistribution(
        scalarNumber(x),
        scalarNumber(mean),
        scalarNumber(standardDeviation),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "LOGNORM.INV",
    ([probability, mean, standardDeviation]) =>
      Math.exp(
        normalInverse(scalarNumber(probability)) * scalarNumber(standardDeviation) +
          scalarNumber(mean),
      ),
  ],
  [
    "LOGNORMDIST",
    ([x, mean, standardDeviation]) =>
      logNormalDistribution(
        scalarNumber(x),
        scalarNumber(mean),
        scalarNumber(standardDeviation),
        true,
      ),
  ],
  [
    "LOGINV",
    ([probability, mean, standardDeviation]) =>
      Math.exp(
        normalInverse(scalarNumber(probability)) * scalarNumber(standardDeviation) +
          scalarNumber(mean),
      ),
  ],
  [
    "LOGNORMDISTCOMPAT",
    ([x, mean, standardDeviation]) =>
      logNormalDistribution(
        scalarNumber(x),
        scalarNumber(mean),
        scalarNumber(standardDeviation),
        true,
      ),
  ],
  [
    "LOGINVCOMPAT",
    ([probability, mean, standardDeviation]) =>
      Math.exp(
        normalInverse(scalarNumber(probability)) * scalarNumber(standardDeviation) +
          scalarNumber(mean),
      ),
  ],
  [
    "NEGBINOM.DIST",
    ([failures, successes, probability, cumulative]) =>
      negativeBinomialDistribution(
        scalarNumber(failures),
        scalarNumber(successes),
        scalarNumber(probability),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "NEGBINOMDIST",
    ([failures, successes, probability, cumulative]) =>
      negativeBinomialDistribution(
        scalarNumber(failures),
        scalarNumber(successes),
        scalarNumber(probability),
        cumulative === undefined ? false : cumulativeArgument(cumulative),
      ),
  ],
  [
    "NEGBINOMDISTCOMPAT",
    ([failures, successes, probability, cumulative]) =>
      negativeBinomialDistribution(
        scalarNumber(failures),
        scalarNumber(successes),
        scalarNumber(probability),
        cumulative === undefined ? false : cumulativeArgument(cumulative),
      ),
  ],
  [
    "NORM.DIST",
    ([x, mean, standardDeviation, cumulative]) => {
      const standardDeviationValue = requirePositive(scalarNumber(standardDeviation));
      const meanValue = scalarNumber(mean);
      const xValue = scalarNumber(x);
      const z = (xValue - meanValue) / standardDeviationValue;
      return cumulativeArgument(cumulative)
        ? normalCdf(z)
        : Math.exp((-z * z) / 2) / (standardDeviationValue * Math.sqrt(2 * Math.PI));
    },
  ],
  [
    "NORM.INV",
    ([probability, mean, standardDeviation]) =>
      normalInverse(scalarNumber(probability)) * scalarNumber(standardDeviation) +
      scalarNumber(mean),
  ],
  [
    "NORM.S.DIST",
    ([z, cumulative]) =>
      cumulativeArgument(cumulative)
        ? normalCdf(scalarNumber(z))
        : Math.exp(-(scalarNumber(z) ** 2) / 2) / Math.sqrt(2 * Math.PI),
  ],
  ["NORM.S.INV", ([probability]) => normalInverse(scalarNumber(probability))],
  ["NORMSINV", ([probability]) => normalInverse(scalarNumber(probability))],
  [
    "NORMDIST",
    ([x, mean, standardDeviation, cumulative]) => {
      const standardDeviationValue = requirePositive(scalarNumber(standardDeviation));
      const z = (scalarNumber(x) - scalarNumber(mean)) / standardDeviationValue;
      return cumulativeArgument(cumulative)
        ? normalCdf(z)
        : Math.exp((-z * z) / 2) / (standardDeviationValue * Math.sqrt(2 * Math.PI));
    },
  ],
  [
    "NORMINV",
    ([probability, mean, standardDeviation]) =>
      normalInverse(scalarNumber(probability)) * scalarNumber(standardDeviation) +
      scalarNumber(mean),
  ],
  [
    "NORMDISTCOMPAT",
    ([x, mean, standardDeviation, cumulative]) => {
      const standardDeviationValue = requirePositive(scalarNumber(standardDeviation));
      const z = (scalarNumber(x) - scalarNumber(mean)) / standardDeviationValue;
      return cumulativeArgument(cumulative)
        ? normalCdf(z)
        : Math.exp((-z * z) / 2) / (standardDeviationValue * Math.sqrt(2 * Math.PI));
    },
  ],
  [
    "NORMINVCOMPAT",
    ([probability, mean, standardDeviation]) =>
      normalInverse(scalarNumber(probability)) * scalarNumber(standardDeviation) +
      scalarNumber(mean),
  ],
  ["NORMSDISTCOMPAT", ([value]) => normalCdf(scalarNumber(value))],
  ["NORMSINVCOMPAT", ([probability]) => normalInverse(scalarNumber(probability))],
  [
    "PERCENTILE",
    ([input, probability]) =>
      percentileValue(flatNumbers([input]), scalarNumber(probability), true),
  ],
  [
    "PERCENTILE.INC",
    ([input, probability]) =>
      percentileValue(flatNumbers([input]), scalarNumber(probability), true),
  ],
  [
    "PERCENTILE.EXC",
    ([input, probability]) =>
      percentileValue(flatNumbers([input]), scalarNumber(probability), false),
  ],
  [
    "PERCENTILECOMPAT",
    ([input, probability]) =>
      percentileValue(flatNumbers([input]), scalarNumber(probability), true),
  ],
  [
    "PERCENTRANKCOMPAT",
    ([input, value, significance]) =>
      percentileRank(
        flatNumbers([input]),
        scalarNumber(value),
        significance === undefined ? undefined : scalarNumber(significance),
        true,
      ),
  ],
  [
    "PERCENTRANK",
    ([input, value, significance]) =>
      percentileRank(
        flatNumbers([input]),
        scalarNumber(value),
        significance === undefined ? undefined : scalarNumber(significance),
        true,
      ),
  ],
  [
    "PERCENTRANK.INC",
    ([input, value, significance]) =>
      percentileRank(
        flatNumbers([input]),
        scalarNumber(value),
        significance === undefined ? undefined : scalarNumber(significance),
        true,
      ),
  ],
  [
    "PERCENTRANK.EXC",
    ([input, value, significance]) =>
      percentileRank(
        flatNumbers([input]),
        scalarNumber(value),
        significance === undefined ? undefined : scalarNumber(significance),
        false,
      ),
  ],
  ...aliases([["POISSONCOMPAT", "POISSON"]]),
  [
    "POISSON.DIST",
    ([x, expected, cumulative]) =>
      poissonDistribution(scalarNumber(x), scalarNumber(expected), cumulativeArgument(cumulative)),
  ],
  [
    "PROB",
    ([xRange, probabilityRange, lower, upper]) => {
      const x = flatNumbers([xRange]);
      const probabilities = flatNumbers([probabilityRange]);
      if (x.length !== probabilities.length) throw new FormulaValueError("#N/A");
      const lowerValue = scalarNumber(lower);
      const upperValue = upper === undefined ? lowerValue : scalarNumber(upper);
      return x.reduce(
        (sum, value, index) =>
          value >= lowerValue && value <= upperValue ? sum + probabilities[index]! : sum,
        0,
      );
    },
  ],
  [
    "SKEWP",
    (args) => {
      const input = flatNumbers(args);
      const center = mean(input);
      const second = input.reduce((sum, value) => sum + (value - center) ** 2, 0) / input.length;
      const third = input.reduce((sum, value) => sum + (value - center) ** 3, 0) / input.length;
      return third / second ** 1.5;
    },
  ],
  ["STDEVPCOMPAT", ([input]) => Math.sqrt(variance(flatNumbers([input]), false))],
  ["STDEVS", ([input]) => Math.sqrt(variance(flatNumbers([input]), true))],
  [
    "STEYX",
    ([yRange, xRange]) => {
      const paired = pairedNumbers(xRange, yRange);
      return standardErrorY(paired.x, paired.y);
    },
  ],
  [
    "T.DIST",
    ([x, degrees, cumulative]) =>
      tDistribution(scalarNumber(x), scalarNumber(degrees), cumulativeArgument(cumulative)),
  ],
  [
    "T.DIST.2T",
    ([x, degrees]) => {
      const value = scalarNumber(x);
      if (value < 0) throw new FormulaValueError("#NUM!");
      return tDistribution(value, scalarNumber(degrees), true, true);
    },
  ],
  [
    "T.DIST.RT",
    ([x, degrees]) => {
      const value = scalarNumber(x);
      const tailCount = scalarNumber(degrees);
      if (value < 0) throw new FormulaValueError("#NUM!");
      return studentTailProbability(value, tailCount, 1);
    },
  ],
  [
    "T.INV",
    ([probability, degrees]) => tInverseLeft(scalarNumber(probability), scalarNumber(degrees)),
  ],
  [
    "T.INV.2T",
    ([probability, degrees]) => tInverseTwoTailed(scalarNumber(probability), scalarNumber(degrees)),
  ],
  [
    "T.TEST",
    ([left, right, tails, type]) => {
      const samples = testSamples(left, right);
      return tTestProbability(
        samples[0],
        samples[1],
        integerArgument(tails),
        integerArgument(type),
      );
    },
  ],
  [
    "TDIST",
    ([x, degrees, tails]) => {
      const value = scalarNumber(x);
      if (value < 0) throw new FormulaValueError("#NUM!");
      const tailCount = integerArgument(tails);
      if (tailCount === 1) return tDistribution(value, scalarNumber(degrees), false);
      if (tailCount === 2) return tDistribution(value, scalarNumber(degrees), false, true);
      throw new FormulaValueError("#NUM!");
    },
  ],
  [
    "TINV",
    ([probability, degrees]) => tInverseTwoTailed(scalarNumber(probability), scalarNumber(degrees)),
  ],
  [
    "TTEST",
    ([left, right, tails, type]) =>
      tTestProbability(...testSamples(left, right), integerArgument(tails), integerArgument(type)),
  ],
  [
    "TRIMMEAN",
    ([input, percent]) => {
      const valuesInput = flatNumbers([input]);
      const percentage = requireProbability(scalarNumber(percent));
      const trimCount = Math.floor((valuesInput.length * percentage) / 2);
      if (valuesInput.length - trimCount * 2 < 1) throw new FormulaValueError("#NUM!");
      const sorted = [...valuesInput].sort((left, right) => left - right);
      return mean(sorted.slice(trimCount, sorted.length - trimCount));
    },
  ],
  ...aliases([
    ["VARCOMPAT", "VAR"],
    ["VARPCOMPAT", "VARP"],
  ]),
  [
    "WEIBULL",
    ([x, shape, scale, cumulative]) =>
      weibullDistribution(
        scalarNumber(x),
        scalarNumber(shape),
        scalarNumber(scale),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "WEIBULL.DIST",
    ([x, shape, scale, cumulative]) =>
      weibullDistribution(
        scalarNumber(x),
        scalarNumber(shape),
        scalarNumber(scale),
        cumulativeArgument(cumulative),
      ),
  ],
  [
    "Z.TEST",
    ([input, value, standardDeviation]) => {
      const sample = flatNumbers([input]);
      const center = mean(sample);
      const sigma =
        standardDeviation === undefined
          ? Math.sqrt(variance(sample, true))
          : requirePositive(scalarNumber(standardDeviation));
      return 1 - normalCdf((center - scalarNumber(value)) / (sigma / Math.sqrt(sample.length)));
    },
  ],
  ["ZTEST", ([input, value, standardDeviation]) => zTest(input, value, standardDeviation)],
  ["QUARTILE", ([input, quart]) => quartileValue(flatNumbers([input]), scalarNumber(quart), true)],
  [
    "QUARTILE.INC",
    ([input, quart]) => quartileValue(flatNumbers([input]), scalarNumber(quart), true),
  ],
  [
    "QUARTILE.EXC",
    ([input, quart]) => quartileValue(flatNumbers([input]), scalarNumber(quart), false),
  ],
  [
    "QUARTILECOMPAT",
    ([input, quart]) => quartileValue(flatNumbers([input]), scalarNumber(quart), true),
  ],
  [
    "RANK",
    ([value, input, order]) =>
      rankValue(
        scalarNumber(value),
        flatNumbers([input]),
        order === undefined ? false : scalarNumber(order) !== 0,
        false,
      ),
  ],
  [
    "RANK.EQ",
    ([value, input, order]) =>
      rankValue(
        scalarNumber(value),
        flatNumbers([input]),
        order === undefined ? false : scalarNumber(order) !== 0,
        false,
      ),
  ],
  [
    "RANK.AVG",
    ([value, input, order]) =>
      rankValue(
        scalarNumber(value),
        flatNumbers([input]),
        order === undefined ? false : scalarNumber(order) !== 0,
        true,
      ),
  ],
  [
    "RANKCOMPAT",
    ([value, input, order]) =>
      rankValue(
        scalarNumber(value),
        flatNumbers([input]),
        order === undefined ? false : scalarNumber(order) !== 0,
        false,
      ),
  ],
]);
