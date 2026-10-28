import { FormulaValueError } from "../errors";
import type { FormulaValue } from "../value";
import { isFormulaErrorValue } from "../value";
import { matchesCriterion, scalarCriterion } from "./criterion";
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
  sumSeries,
  studentTailProbability,
} from "./statistical-numeric";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { numbers, pairedNumbers, scalarBoolean, scalarNumber, values } from "./utils";

function flatNumbers(args: Parameters<FormulaFunctionImplementation>[0]): number[] {
  return args.flatMap(numbers);
}

function pairedData(
  left: FormulaValue,
  right: FormulaValue,
  calculate: (left: number[], right: number[]) => number,
): number {
  const pairs = pairedNumbers(left, right);
  const leftValues = pairs.map(([value]) => value);
  const rightValues = pairs.map(([, value]) => value);
  return calculate(leftValues, rightValues);
}

function pairedRegression(
  knownX: FormulaValue,
  knownY: FormulaValue,
): { slope: number; intercept: number } {
  const pairs = pairedNumbers(knownX, knownY);
  return regression(
    pairs.map(([value]) => value),
    pairs.map(([, value]) => value),
  );
}

function allValues(
  args: Parameters<FormulaFunctionImplementation>[0],
): Parameters<FormulaFunctionImplementation>[0][number][] {
  return args.flatMap(values);
}

function allANumbers(args: Parameters<FormulaFunctionImplementation>[0]): number[] {
  return args.flatMap((argument) => {
    if (Array.isArray(argument)) {
      return argument.flat().flatMap((value) => (value === null ? [] : [coerceANumber(value)]));
    }
    if (argument === null) return [];
    return [coerceANumber(argument)];
  });
}

function coerceANumber(value: Parameters<FormulaFunctionImplementation>[0][number]): number {
  if (isFormulaErrorValue(value)) throw new FormulaValueError(value.value);
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return Number(value);
  return 0;
}

function mean(input: number[]): number {
  return input.reduce((sum, value) => sum + value, 0) / input.length;
}

function variance(input: number[], sample: boolean): number {
  if (input.length < (sample ? 2 : 1)) throw new FormulaValueError("#NUM!");
  const center = mean(input);
  return (
    input.reduce((sum, value) => sum + (value - center) ** 2, 0) / (input.length - (sample ? 1 : 0))
  );
}

function skewness(input: number[], population: boolean): number {
  const center = mean(input);
  const divisor = input.length - (population ? 0 : 1);
  const second = input.reduce((sum, value) => sum + (value - center) ** 2, 0) / divisor;
  const third = input.reduce((sum, value) => sum + (value - center) ** 3, 0) / divisor;
  const raw = third / second ** 1.5;
  return population ? raw : (input.length / ((input.length - 1) * (input.length - 2))) * raw;
}

function kurtosis(input: number[]): number {
  const center = mean(input);
  const divisor = input.length - 1;
  const fourth = input.reduce((sum, value) => sum + (value - center) ** 4, 0) / divisor;
  const second = input.reduce((sum, value) => sum + (value - center) ** 2, 0) / divisor;
  return (
    (input.length * (input.length + 1) * fourth) /
      ((input.length - 1) * (input.length - 2) * second ** 2) -
    (3 * (input.length - 1) ** 2) / ((input.length - 2) * (input.length - 3))
  );
}

function extreme(input: number[], select: (candidate: number, best: number) => number): number {
  if (!input.length) return 0;
  return input.reduce((best, candidate) => (select(candidate, best) < 0 ? candidate : best));
}

function covariance(left: number[], right: number[], population: boolean): number {
  const centerX = mean(left);
  const centerY = mean(right);
  return (
    left.reduce((sum, value, index) => sum + (value - centerX) * (right[index]! - centerY), 0) /
    (left.length - (population ? 0 : 1))
  );
}

function matchingIndexes(
  range: Parameters<FormulaFunctionImplementation>[0][number],
  criteria: Parameters<FormulaFunctionImplementation>[0][number],
): number[] {
  return values(range)
    .map((value, index) =>
      matchesCriterion(value ?? null, scalarCriterion(criteria)) ? index : -1,
    )
    .filter((index) => index >= 0);
}

function criterionValue(
  value: Parameters<FormulaFunctionImplementation>[0][number],
  criteria: Parameters<FormulaFunctionImplementation>[0][number],
): boolean {
  return matchesCriterion(value ?? null, scalarCriterion(criteria));
}

function rankValue(value: number, input: number[], descending: boolean): number | "#N/A" {
  if (!input.includes(value)) return "#N/A";
  const sorted = [...input].sort((left, right) => (descending ? right - left : left - right));
  return sorted.findIndex((entry) => entry === value) + 1;
}

export const statisticalFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  [
    "AVEDEV",
    (args) => {
      const input = flatNumbers(args);
      const center = mean(input);
      return input.reduce((sum, value) => sum + Math.abs(value - center), 0) / input.length;
    },
  ],
  [
    "AVERAGE",
    (args) => {
      const input = flatNumbers(args);
      return input.length === 0 ? "#DIV/0!" : mean(input);
    },
  ],
  [
    "AVERAGEA",
    (args) => {
      const input = allANumbers(args);
      return input.length === 0 ? "#DIV/0!" : mean(input);
    },
  ],
  [
    "AVERAGEIF",
    ([range, criteria, averageRange]) => {
      const indexes = matchingIndexes(range, criteria);
      const source = averageRange === undefined ? range : averageRange;
      const selected = indexes
        .map((index) => values(source)[index])
        .filter((value): value is number => typeof value === "number");
      return selected.length === 0 ? "#DIV/0!" : mean(selected);
    },
  ],
  [
    "AVERAGEIFS",
    (args) => {
      const averageRange = values(args[0]!);
      const indexes = averageRange
        .map((_, index) => index)
        .filter((index) => {
          for (let position = 1; position < args.length; position += 2)
            if (!criterionValue(values(args[position]!)[index] ?? null, args[position + 1]!))
              return false;
          return true;
        })
        .map((index) => averageRange[index])
        .filter((value): value is number => typeof value === "number");
      return indexes.length === 0 ? "#DIV/0!" : mean(indexes);
    },
  ],
  ["COUNT", (args) => allValues(args).filter((value) => typeof value === "number").length],
  ["COUNTA", (args) => allValues(args).filter((value) => value !== null).length],
  [
    "COUNTBLANK",
    ([range]) => values(range).filter((value) => value === null || value === "").length,
  ],
  ["COUNTIF", ([range, criteria]) => matchingIndexes(range, criteria).length],
  [
    "COUNTIFS",
    (args) => {
      const first = values(args[0]!);
      return first
        .map((_, index) => index)
        .filter((index) => {
          for (let position = 0; position < args.length; position += 2)
            if (!criterionValue(values(args[position]!)[index] ?? null, args[position + 1]!))
              return false;
          return true;
        }).length;
    },
  ],
  ["MAX", (args) => extreme(flatNumbers(args), (candidate, best) => best - candidate)],
  ["MAXA", (args) => extreme(allANumbers(args), (candidate, best) => best - candidate)],
  [
    "MAXIFS",
    (args) => {
      const selected = conditionalValues(args);
      return extreme(selected, (candidate, best) => best - candidate);
    },
  ],
  [
    "MEDIAN",
    (args) => {
      const input = flatNumbers(args).sort((left, right) => left - right);
      if (!input.length) return "#NUM!";
      const middle = Math.floor(input.length / 2);
      return input.length % 2 ? input[middle]! : (input[middle - 1]! + input[middle]!) / 2;
    },
  ],
  ["MIN", (args) => extreme(flatNumbers(args), (candidate, best) => candidate - best)],
  ["MINA", (args) => extreme(allANumbers(args), (candidate, best) => candidate - best)],
  [
    "MINIFS",
    (args) => {
      const selected = conditionalValues(args);
      return extreme(selected, (candidate, best) => candidate - best);
    },
  ],
  ["MODE", (args) => mode(flatNumbers(args), "frequent")],
  ["MODE.SNGL", (args) => mode(flatNumbers(args), "frequent")],
  ["MODE.MULT", (args) => mode(flatNumbers(args), "all")],
  ["LARGE", ([range, index]) => ordered(range)[scalarNumber(index) - 1] ?? "#NUM!"],
  ["SMALL", ([range, index]) => ordered(range, true)[scalarNumber(index) - 1] ?? "#NUM!"],
  [
    "RANK",
    ([value, range, order]) =>
      rankValue(
        scalarNumber(value),
        numbers(range),
        order !== undefined && scalarNumber(order) !== 0,
      ),
  ],
  [
    "RANK.EQ",
    ([value, range, order]) =>
      rankValue(
        scalarNumber(value),
        numbers(range),
        order !== undefined && scalarNumber(order) !== 0,
      ),
  ],
  [
    "RANK.AVG",
    ([value, range, order]) => {
      const input = numbers(range);
      const first = rankValue(
        scalarNumber(value),
        input,
        order !== undefined && scalarNumber(order) !== 0,
      );
      if (first === "#N/A") return first;
      const count = input.filter((entry) => entry === scalarNumber(value)).length;
      return first + (count - 1) / 2;
    },
  ],
  [
    "PERCENTILE",
    ([range, probability]) => percentile(numbers(range), scalarNumber(probability), true),
  ],
  [
    "PERCENTILE.INC",
    ([range, probability]) => percentile(numbers(range), scalarNumber(probability), true),
  ],
  [
    "PERCENTILE.EXC",
    ([range, probability]) => percentile(numbers(range), scalarNumber(probability), false),
  ],
  ["QUARTILE", ([range, quartile]) => percentile(numbers(range), scalarNumber(quartile) / 4, true)],
  [
    "QUARTILE.INC",
    ([range, quartile]) => percentile(numbers(range), scalarNumber(quartile) / 4, true),
  ],
  [
    "QUARTILE.EXC",
    ([range, quartile]) => percentile(numbers(range), scalarNumber(quartile) / 4, false),
  ],
  ["VAR", (args) => variance(flatNumbers(args), true)],
  ["VAR.S", (args) => variance(flatNumbers(args), true)],
  ["VARA", (args) => variance(allANumbers(args), true)],
  ["VARP", (args) => variance(flatNumbers(args), false)],
  ["VAR.P", (args) => variance(flatNumbers(args), false)],
  ["VARPA", (args) => variance(allANumbers(args), false)],
  ["STDEV", (args) => Math.sqrt(variance(flatNumbers(args), true))],
  ["STDEV.S", (args) => Math.sqrt(variance(flatNumbers(args), true))],
  ["STDEVA", (args) => Math.sqrt(variance(allANumbers(args), true))],
  ["STDEVP", (args) => Math.sqrt(variance(flatNumbers(args), false))],
  ["STDEV.P", (args) => Math.sqrt(variance(flatNumbers(args), false))],
  ["STDEVPA", (args) => Math.sqrt(variance(allANumbers(args), false))],
  [
    "DEVSQ",
    (args) => {
      const input = flatNumbers(args);
      const center = mean(input);
      return input.reduce((sum, value) => sum + (value - center) ** 2, 0);
    },
  ],
  [
    "GEOMEAN",
    (args) => {
      const input = flatNumbers(args);
      return Math.exp(mean(input.map(Math.log)));
    },
  ],
  [
    "HARMEAN",
    (args) => {
      const input = flatNumbers(args);
      return input.length / input.reduce((sum, value) => sum + 1 / value, 0);
    },
  ],
  ["SKEW", (args) => skewness(flatNumbers(args), false)],
  ["SKEW.P", (args) => skewness(flatNumbers(args), true)],
  ["KURT", (args) => kurtosis(flatNumbers(args))],
  ["COVAR", ([left, right]) => pairedData(left, right, (x, y) => covariance(x, y, true))],
  ["COVARIANCE.P", ([left, right]) => pairedData(left, right, (x, y) => covariance(x, y, true))],
  ["COVARIANCE.S", ([left, right]) => pairedData(left, right, (x, y) => covariance(x, y, false))],
  ["CORREL", ([left, right]) => pairedData(left, right, correlation)],
  ["PEARSON", ([left, right]) => pairedData(left, right, correlation)],
  ["RSQ", ([left, right]) => pairedData(left, right, (x, y) => correlation(x, y) ** 2)],
  ["SLOPE", ([knownY, knownX]) => pairedData(knownX, knownY, (x, y) => regression(x, y).slope)],
  [
    "INTERCEPT",
    ([knownY, knownX]) => pairedData(knownX, knownY, (x, y) => regression(x, y).intercept),
  ],
  [
    "FORECAST",
    ([value, knownY, knownX]) => {
      const regressionLine = pairedRegression(knownX, knownY);
      return regressionLine.intercept + regressionLine.slope * scalarNumber(value);
    },
  ],
  [
    "FORECAST.LINEAR",
    ([value, knownY, knownX]) => {
      const regressionLine = pairedRegression(knownX, knownY);
      return regressionLine.intercept + regressionLine.slope * scalarNumber(value);
    },
  ],
  [
    "STEYX",
    ([knownY, knownX]) => {
      const pairs = pairedNumbers(knownX, knownY);
      const xValues = pairs.map(([value]) => value);
      const yValues = pairs.map(([, value]) => value);
      const line = regression(xValues, yValues);
      const residual = yValues.reduce(
        (sum, value, index) => sum + (value - (line.intercept + line.slope * xValues[index]!)) ** 2,
        0,
      );
      return Math.sqrt(residual / (xValues.length - 2));
    },
  ],
  [
    "NORM.DIST",
    ([value, center, deviation, cumulative]) =>
      distribution(
        scalarNumber(value),
        () => normalCdf((scalarNumber(value) - scalarNumber(center)) / scalarNumber(deviation)),
        Math.exp(
          -((scalarNumber(value) - scalarNumber(center)) ** 2) / (2 * scalarNumber(deviation) ** 2),
        ) /
          (scalarNumber(deviation) * Math.sqrt(2 * Math.PI)),
        scalarBoolean(cumulative ?? true),
      ),
  ],
  [
    "NORMDIST",
    ([value, center, deviation]) =>
      normalCdf((scalarNumber(value) - scalarNumber(center)) / scalarNumber(deviation)),
  ],
  [
    "NORM.INV",
    ([probability, center, deviation]) =>
      normalInverse(scalarNumber(probability)) * scalarNumber(deviation) + scalarNumber(center),
  ],
  [
    "NORMINV",
    ([probability, center, deviation]) =>
      normalInverse(scalarNumber(probability)) * scalarNumber(deviation) + scalarNumber(center),
  ],
  [
    "NORM.S.DIST",
    ([value, cumulative]) =>
      distribution(
        scalarNumber(value),
        () => normalCdf(scalarNumber(value)),
        Math.exp(-(scalarNumber(value) ** 2) / 2) / Math.sqrt(2 * Math.PI),
        cumulative === undefined || scalarBoolean(cumulative),
      ),
  ],
  ["NORMSDIST", ([value]) => normalCdf(scalarNumber(value))],
  ["NORM.S.INV", ([probability]) => normalInverse(scalarNumber(probability))],
  ["NORMSINV", ([probability]) => normalInverse(scalarNumber(probability))],
  [
    "STANDARDIZE",
    ([value, center, deviation]) =>
      (scalarNumber(value) - scalarNumber(center)) / scalarNumber(deviation),
  ],
  [
    "POISSON.DIST",
    ([value, expected, cumulative]) =>
      poisson(
        scalarNumber(value),
        scalarNumber(expected),
        cumulative === undefined || scalarBoolean(cumulative),
      ),
  ],
  [
    "POISSON",
    ([value, expected, cumulative]) =>
      poisson(
        scalarNumber(value),
        scalarNumber(expected),
        cumulative === undefined || scalarBoolean(cumulative),
      ),
  ],
  [
    "EXPON.DIST",
    ([value, lambda, cumulative]) => {
      const rate = scalarNumber(lambda);
      const x = scalarNumber(value);
      return distribution(
        x,
        () => 1 - Math.exp(-rate * x),
        rate * Math.exp(-rate * x),
        scalarBoolean(cumulative),
      );
    },
  ],
  [
    "EXPONDIST",
    ([value, lambda, cumulative]) => {
      const rate = scalarNumber(lambda);
      const x = scalarNumber(value);
      return distribution(
        x,
        () => 1 - Math.exp(-rate * x),
        rate * Math.exp(-rate * x),
        scalarBoolean(cumulative ?? true),
      );
    },
  ],
  [
    "WEIBULL.DIST",
    ([value, shape, scale, cumulative]) => {
      const x = scalarNumber(value);
      const alpha = scalarNumber(shape);
      const beta = scalarNumber(scale);
      return distribution(
        x,
        () => 1 - Math.exp(-((x / beta) ** alpha)),
        (alpha / beta) * (x / beta) ** (alpha - 1) * Math.exp(-((x / beta) ** alpha)),
        scalarBoolean(cumulative),
      );
    },
  ],
  [
    "WEIBULL",
    ([value, shape, scale]) =>
      1 - Math.exp(-((scalarNumber(value) / scalarNumber(scale)) ** scalarNumber(shape))),
  ],
  ["GAMMALN", ([value]) => gammaLn(scalarNumber(value))],
  [
    "GAMMA.DIST",
    ([value, shape, scale, cumulative]) =>
      gammaDistribution(
        scalarNumber(value),
        scalarNumber(shape),
        scalarNumber(scale),
        scalarBoolean(cumulative ?? true),
      ),
  ],
  [
    "GAMMADIST",
    ([value, shape, scale]) =>
      gammaDistribution(scalarNumber(value), scalarNumber(shape), scalarNumber(scale), true),
  ],
  [
    "GAMMA.INV",
    ([probability, shape, scale]) =>
      gammaInverse(scalarNumber(probability), scalarNumber(shape), scalarNumber(scale)),
  ],
  [
    "GAMMAINV",
    ([probability, shape, scale]) =>
      gammaInverse(scalarNumber(probability), scalarNumber(shape), scalarNumber(scale)),
  ],
  [
    "CHISQ.DIST",
    ([value, degrees, cumulative]) =>
      chiSquare(scalarNumber(value), scalarNumber(degrees), scalarBoolean(cumulative ?? true)),
  ],
  [
    "CHIDIST",
    ([value, degrees]) => 1 - chiSquare(scalarNumber(value), scalarNumber(degrees), true),
  ],
  [
    "CHISQ.INV",
    ([probability, degrees]) => chiSquareInverse(scalarNumber(probability), scalarNumber(degrees)),
  ],
  [
    "CHIINV",
    ([probability, degrees]) =>
      chiSquareInverse(1 - scalarNumber(probability), scalarNumber(degrees)),
  ],
  [
    "F.DIST",
    ([value, numerator, denominator, cumulative]) =>
      fDistribution(
        scalarNumber(value),
        scalarNumber(numerator),
        scalarNumber(denominator),
        scalarBoolean(cumulative ?? true),
      ),
  ],
  [
    "FDIST",
    ([value, numerator, denominator]) =>
      1 -
      fDistribution(scalarNumber(value), scalarNumber(numerator), scalarNumber(denominator), true),
  ],
  [
    "F.INV",
    ([probability, numerator, denominator]) =>
      fInverse(scalarNumber(probability), scalarNumber(numerator), scalarNumber(denominator)),
  ],
  [
    "FINV",
    ([probability, numerator, denominator]) =>
      fInverse(1 - scalarNumber(probability), scalarNumber(numerator), scalarNumber(denominator)),
  ],
  [
    "T.DIST",
    ([value, degrees, cumulative]) =>
      tDistribution(
        scalarNumber(value),
        scalarNumber(degrees),
        scalarBoolean(cumulative ?? true),
        false,
      ),
  ],
  [
    "TDIST",
    ([value, degrees, tails]) =>
      tDistribution(scalarNumber(value), scalarNumber(degrees), true, scalarNumber(tails) === 2),
  ],
  [
    "T.INV",
    ([probability, degrees]) => tInverse(scalarNumber(probability), scalarNumber(degrees), false),
  ],
  [
    "TINV",
    ([probability, degrees]) => tInverse(scalarNumber(probability), scalarNumber(degrees), true),
  ],
  [
    "BETA.DIST",
    ([value, alpha, beta, cumulative]) =>
      betaDistribution(
        scalarNumber(value),
        scalarNumber(alpha),
        scalarNumber(beta),
        scalarBoolean(cumulative ?? true),
      ),
  ],
  [
    "BETADIST",
    ([value, alpha, beta]) =>
      incompleteBeta(scalarNumber(value), scalarNumber(alpha), scalarNumber(beta)),
  ],
  [
    "BETA.INV",
    ([probability, alpha, beta]) =>
      inverseIncompleteBeta(scalarNumber(probability), scalarNumber(alpha), scalarNumber(beta)),
  ],
  [
    "BINOM.DIST",
    ([successes, trials, probability, cumulative]) =>
      binomial(
        scalarNumber(successes),
        scalarNumber(trials),
        scalarNumber(probability),
        scalarBoolean(cumulative ?? true),
      ),
  ],
  [
    "BINOMDIST",
    ([successes, trials, probability, cumulative]) =>
      binomial(
        scalarNumber(successes),
        scalarNumber(trials),
        scalarNumber(probability),
        scalarBoolean(cumulative ?? true),
      ),
  ],
  [
    "HYPGEOM.DIST",
    ([successes, sample, populationSuccesses, population, cumulative]) => {
      const x = scalarNumber(successes);
      const n = scalarNumber(sample);
      const m = scalarNumber(populationSuccesses);
      const size = scalarNumber(population);
      const pmf = (combination(m, x) * combination(size - m, n - x)) / combination(size, n);
      return cumulative
        ? sumSeries(
            0,
            Math.trunc(x),
            (index) =>
              (combination(m, index) * combination(size - m, n - index)) / combination(size, n),
          )
        : pmf;
    },
  ],
  [
    "NEGBINOM.DIST",
    ([failures, successes, probability, cumulative]) => {
      const r = scalarNumber(successes);
      const p = scalarNumber(probability);
      const x = scalarNumber(failures);
      const pmf = combination(x + r - 1, x) * p ** r * (1 - p) ** x;
      return cumulative
        ? sumSeries(
            0,
            Math.trunc(x),
            (index) => combination(index + r - 1, index) * p ** r * (1 - p) ** index,
          )
        : pmf;
    },
  ],
  ["FISHER", ([value]) => 0.5 * Math.log((1 + scalarNumber(value)) / (1 - scalarNumber(value)))],
  [
    "FISHERINV",
    ([value]) => (Math.exp(2 * scalarNumber(value)) - 1) / (Math.exp(2 * scalarNumber(value)) + 1),
  ],
  ["PHI", ([value]) => normalCdf(scalarNumber(value)) - 0.5],
  ["GAUSS", ([value]) => normalCdf(scalarNumber(value)) - 0.5],
  [
    "PERMUT",
    ([count, chosen]) =>
      factorial(scalarNumber(count)) / factorial(scalarNumber(count) - scalarNumber(chosen)),
  ],
  ["PERMUTATIONA", ([count, chosen]) => scalarNumber(count) ** scalarNumber(chosen)],
]);

function conditionalValues(args: Parameters<FormulaFunctionImplementation>[0]): number[] {
  const source = values(args[0]!);
  return source
    .map((_, index) => index)
    .filter((index) => {
      for (let position = 1; position < args.length; position += 2)
        if (!criterionValue(values(args[position]!)[index] ?? null, args[position + 1]!))
          return false;
      return true;
    })
    .map((index) => scalarNumber(source[index] ?? 0));
}

function ordered(
  range: Parameters<FormulaFunctionImplementation>[0][number],
  ascending = false,
): number[] {
  return numbers(range).sort((left, right) => (ascending ? left - right : right - left));
}

function percentile(input: number[], probability: number, inclusive: boolean): number {
  const sorted = [...input].sort((left, right) => left - right);
  const position = inclusive
    ? probability * (sorted.length - 1)
    : probability * (sorted.length + 1) - 1;
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  return sorted[lower]! + (position - lower) * (sorted[upper]! - sorted[lower]!);
}

function mode(input: number[], output: "frequent" | "all"): number | string | number[][] {
  const counts = new Map<number, number>();
  for (const value of input) counts.set(value, (counts.get(value) ?? 0) + 1);
  let maximum = 0;
  for (const count of counts.values()) maximum = count > maximum ? count : maximum;
  if (maximum <= 1) return "#N/A";
  const modes = [...counts].filter(([, count]) => count === maximum).map(([value]) => value);
  if (output === "frequent") return modes[0] ?? "#N/A";
  return modes.length ? modes.map((value) => [value]) : "#N/A";
}

function regression(xValues: number[], yValues: number[]): { slope: number; intercept: number } {
  const centerX = mean(xValues);
  const centerY = mean(yValues);
  const slope =
    xValues.reduce(
      (sum, value, index) => sum + (value - centerX) * (yValues[index]! - centerY),
      0,
    ) / xValues.reduce((sum, value) => sum + (value - centerX) ** 2, 0);
  return { slope, intercept: centerY - slope * centerX };
}

function correlation(left: number[], right: number[]): number {
  const deviation = Math.sqrt(variance(left, true) * variance(right, true));
  return covariance(left, right, false) / deviation;
}

function distribution(
  value: number,
  cumulative: () => number,
  mass: number,
  wantCumulative: boolean,
): number {
  if (value < 0) throw new FormulaValueError("#NUM!");
  return wantCumulative ? cumulative() : mass;
}

function poisson(value: number, expected: number, cumulative: boolean): number {
  if (value < 0 || expected < 0) throw new FormulaValueError("#NUM!");
  if (!cumulative)
    return (Math.exp(-expected) * expected ** Math.trunc(value)) / factorial(Math.trunc(value));
  return sumSeries(
    0,
    Math.trunc(value),
    (index) => (Math.exp(-expected) * expected ** index) / factorial(index),
  );
}

function gammaDistribution(
  value: number,
  shape: number,
  scale: number,
  cumulative: boolean,
): number {
  if (value < 0 || shape <= 0 || scale <= 0) throw new FormulaValueError("#NUM!");
  const standardized = value / scale;
  if (!cumulative)
    return (Math.exp(-standardized) * standardized ** (shape - 1)) / (scale * gamma(shape));
  return lowerGamma(shape, standardized) / gamma(shape);
}

function gammaInverse(probability: number, shape: number, scale: number): number {
  let low = 0;
  let high = shape * scale + 10 * Math.sqrt(shape) * scale + 10;
  for (let index = 0; index < 100; index += 1) {
    const middle = (low + high) / 2;
    if (gammaDistribution(middle, shape, scale, true) < probability) low = middle;
    else high = middle;
  }
  return (low + high) / 2;
}

function chiSquare(value: number, degrees: number, cumulative: boolean): number {
  return cumulative
    ? lowerGamma(degrees / 2, value / 2) / gamma(degrees / 2)
    : (Math.exp(-value / 2) * value ** (degrees / 2 - 1)) /
        (2 ** (degrees / 2) * gamma(degrees / 2));
}

function chiSquareInverse(probability: number, degrees: number): number {
  let low = 0;
  let high = degrees * 10;
  for (let index = 0; index < 100; index += 1) {
    const middle = (low + high) / 2;
    if (chiSquare(middle, degrees, true) < probability) low = middle;
    else high = middle;
  }
  return (low + high) / 2;
}

function fDistribution(
  value: number,
  numerator: number,
  denominator: number,
  cumulative: boolean,
): number {
  const x = (numerator * value) / (numerator * value + denominator);
  const cdf = incompleteBeta(x, numerator / 2, denominator / 2);
  if (cumulative) return cdf;
  const scale = numerator / denominator;
  return (
    (scale ** (numerator / 2) *
      value ** (numerator / 2 - 1) *
      (1 + scale * value) ** (-(numerator + denominator) / 2)) /
    betaFunction(numerator / 2, denominator / 2)
  );
}

function fInverse(probability: number, numerator: number, denominator: number): number {
  const x = inverseIncompleteBeta(probability, numerator / 2, denominator / 2);
  return (denominator * x) / (numerator * (1 - x));
}

function tDistribution(
  value: number,
  degrees: number,
  cumulative: boolean,
  twoTailed: boolean,
): number {
  const tail = studentTailProbability(value, degrees, twoTailed ? 2 : 1);
  if (!cumulative)
    return (
      (gamma((degrees + 1) / 2) / (Math.sqrt(degrees * Math.PI) * gamma(degrees / 2))) *
      (1 + (value * value) / degrees) ** (-(degrees + 1) / 2)
    );
  return value >= 0 ? 1 - tail : tail;
}

function tInverse(probability: number, degrees: number, twoTailed: boolean): number {
  const target = twoTailed ? 1 - probability / 2 : probability;
  let low = -10;
  let high = 10;
  for (let index = 0; index < 100; index += 1) {
    const middle = (low + high) / 2;
    if (tDistribution(middle, degrees, true, false) < target) low = middle;
    else high = middle;
  }
  return (low + high) / 2;
}

function betaDistribution(value: number, alpha: number, beta: number, cumulative: boolean): number {
  if (!cumulative)
    return (value ** (alpha - 1) * (1 - value) ** (beta - 1)) / betaFunction(alpha, beta);
  return incompleteBeta(value, alpha, beta);
}

function binomial(
  successes: number,
  trials: number,
  probability: number,
  cumulative: boolean,
): number {
  const mass = (index: number) =>
    combination(trials, index) * probability ** index * (1 - probability) ** (trials - index);
  return cumulative ? sumSeries(0, Math.trunc(successes), mass) : mass(Math.trunc(successes));
}
