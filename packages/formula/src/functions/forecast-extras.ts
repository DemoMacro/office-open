import { FormulaValueError } from "../errors";
import type { FormulaMatrix, FormulaValue } from "../value";
import { isMatrix } from "../value";
import {
  fitLinearRegression,
  predictLinearValues,
  predictionRows,
  regressionObservations,
} from "./regression";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { numbers, pairedNumbers, scalarBoolean, scalarNumber, scalarText } from "./utils";

type Arguments = Parameters<FormulaFunctionImplementation>[0];

interface RegressionLine {
  slope: number;
  intercept: number;
}

function linearRegression(pairs: Array<[number, number]>): RegressionLine {
  if (pairs.length < 2) throw new FormulaValueError("#DIV/0!");
  const meanX = pairs.reduce((sum, [x]) => sum + x, 0) / pairs.length;
  const meanY = pairs.reduce((sum, [, y]) => sum + y, 0) / pairs.length;
  let covariance = 0;
  let variance = 0;
  for (const [x, y] of pairs) {
    covariance += (x - meanX) * (y - meanY);
    variance += (x - meanX) ** 2;
  }
  if (variance === 0) throw new FormulaValueError("#DIV/0!");
  const slope = covariance / variance;
  return { slope, intercept: meanY - slope * meanX };
}

function frequency(data: FormulaValue, bins: FormulaValue): FormulaMatrix {
  const values = numbers(data);
  const boundaries = numbers(bins).sort((left, right) => left - right);
  const counts = Array.from({ length: boundaries.length + 1 }, () => 0);
  for (const value of values) {
    const index = boundaries.findIndex((boundary) => value <= boundary);
    counts[index === -1 ? boundaries.length : index] += 1;
  }
  return counts.map((count) => [count]);
}

function schedule(presentValue: FormulaValue, schedules: FormulaValue): number {
  if (!isMatrix(schedules)) return scalarNumber(presentValue);
  return schedules
    .flat()
    .reduce(
      (compound: number, rate) => compound * (1 + scalarNumber(rate)),
      scalarNumber(presentValue),
    );
}

function seriesSum(
  input: FormulaValue,
  exponent: FormulaValue,
  step: FormulaValue,
  coefficients: FormulaValue,
): FormulaValue {
  const x = scalarNumber(input);
  const firstExponent = scalarNumber(exponent);
  const exponentStep = scalarNumber(step);
  return numbers(coefficients).reduce(
    (sum, coefficient, index) => sum + coefficient * x ** (firstExponent + index * exponentStep),
    0,
  );
}

function sortedTimeline(values: FormulaValue, timeline: FormulaValue): Array<[number, number]> {
  return pairedNumbers(timeline, values).sort(([left], [right]) => left - right);
}

function seasonalPeriod(
  pairs: Array<[number, number]>,
  requested: FormulaValue | undefined,
): number {
  const mode = requested === undefined ? 1 : scalarNumber(requested);
  if (mode === 0) return 1;
  const maximum = Math.floor(pairs.length / 2);
  const candidate = mode === 1 ? (pairs.length >= 8 ? 4 : 1) : Math.trunc(mode);
  return Math.max(1, Math.min(candidate, Math.max(1, maximum)));
}

function seasonalIndices(pairs: Array<[number, number]>, period: number): number[] {
  if (period === 1) return [1];
  const groups = Array.from({ length: period }, (): number[] => []);
  pairs.forEach(([, value], index) => groups[index % period]?.push(value));
  const overall = pairs.reduce((sum, [, value]) => sum + value, 0) / pairs.length;
  return groups.map((group) =>
    group.length ? group.reduce((sum, value) => sum + value, 0) / group.length / overall : 1,
  );
}

interface EtsModel {
  alpha: number;
  beta: number;
  gamma: number;
  level: number;
  trend: number;
  seasonal: number[];
  residuals: number[];
  period: number;
  multiplicative: boolean;
}

function etsState(
  pairs: Array<[number, number]>,
  period: number,
  alpha: number,
  beta: number,
  gamma: number,
  multiplicative: boolean,
): EtsModel {
  const values = pairs.map(([, value]) => value);
  const initialSeasonal = seasonalIndices(pairs, period);
  let level = values.reduce((sum, value) => sum + value, 0) / values.length;
  let trend = (values.at(-1)! - values[0]!) / Math.max(1, values.length - 1);
  const seasonal = [...initialSeasonal];
  const residuals: number[] = [];
  pairs.forEach(([, value], index) => {
    const phase = index % period;
    const seasonalValue = multiplicative
      ? Math.max(Number.EPSILON, seasonal[phase]!)
      : seasonal[phase]!;
    const forecast =
      (level + trend) * (multiplicative ? seasonalValue : 1) + (multiplicative ? 0 : seasonalValue);
    residuals.push(value - forecast);
    const deseasonalized = multiplicative
      ? value / Math.max(Number.EPSILON, seasonalValue)
      : value - seasonalValue;
    const previousLevel = level;
    level = alpha * deseasonalized + (1 - alpha) * (level + trend);
    trend = beta * (level - previousLevel) + (1 - beta) * trend;
    const seasonalUpdate =
      value - level === 0
        ? 0
        : multiplicative
          ? value / Math.max(Number.EPSILON, level) - 1
          : value - level;
    seasonal[phase] = gamma * seasonalUpdate + (1 - gamma) * seasonal[phase]!;
  });
  return { alpha, beta, gamma, level, trend, seasonal, residuals, period, multiplicative };
}

function fitEts(pairs: Array<[number, number]>, period: number, multiplicative: boolean): EtsModel {
  let best: EtsModel | undefined;
  let bestError = Number.POSITIVE_INFINITY;
  for (let alpha = 0.05; alpha <= 0.96; alpha += 0.05) {
    for (let beta = 0.05; beta <= 0.96; beta += 0.05) {
      for (let gamma = 0.05; gamma <= 0.96; gamma += 0.05) {
        const model = etsState(pairs, period, alpha, beta, gamma, multiplicative);
        const error = model.residuals.reduce((total, residual) => total + residual ** 2, 0);
        if (Number.isFinite(error) && error < bestError) {
          bestError = error;
          best = model;
        }
      }
    }
  }
  return best ?? etsState(pairs, period, 0.2, 0.1, 0.1, multiplicative);
}

function isEtsModel(value: EtsModel | FormulaValue): value is EtsModel {
  return typeof value === "object" && value !== null && "residuals" in value;
}

function etsModel(args: Arguments, multiplicative: boolean): EtsModel | FormulaValue {
  if (args.length < 3 || args.length > 6) return "#VALUE!";
  const pairs = sortedTimeline(args[1]!, args[2]!);
  if (pairs.length < 2) return "#VALUE!";
  return fitEts(pairs, seasonalPeriod(pairs, args[3]), multiplicative);
}

function etsForecastValue(model: EtsModel, target: number): number {
  const phase = Math.max(0, Math.round(target)) % model.period;
  const seasonalValue = model.seasonal[phase] ?? 1;
  const base = model.level + model.trend;
  return model.multiplicative ? base * seasonalValue : base + seasonalValue;
}

function exponentialSmoothingForecast(args: Arguments, multiplicative: boolean): FormulaValue {
  const model = etsModel(args, multiplicative);
  if (!isEtsModel(model)) return model;
  const timeline = sortedTimeline(args[1]!, args[2]!).map(([time]) => time);
  const target = scalarNumber(args[0]!);
  const maximum = Math.max(...timeline);
  return etsForecastValue(model, Math.max(target, maximum));
}

function predictionInterval(args: Arguments, multiplicative: boolean): FormulaValue {
  if (args.length < 4 || args.length > 7) return "#VALUE!";
  const confidence = scalarNumber(args[3]!);
  if (confidence <= 0 || confidence >= 1) return "#NUM!";
  const model = etsModel([args[0]!, args[1]!, args[2]!, args[4], args[5]], multiplicative);
  if (!isEtsModel(model)) return model;
  const timeline = sortedTimeline(args[1]!, args[2]!).map(([time]) => time);
  const target = scalarNumber(args[0]!);
  const forecast = etsForecastValue(model, Math.max(target, Math.max(...timeline)));
  const meanSquare =
    model.residuals.reduce((total, value) => total + value ** 2, 0) /
    Math.max(1, model.residuals.length);
  const z = inverseNormal(1 - (1 - confidence) / 2);
  return z * Math.sqrt(meanSquare) * (multiplicative ? Math.abs(forecast) : 1);
}

function smoothingStatistic(args: Arguments): FormulaValue {
  if (args.length < 4 || args.length > 6) return "#VALUE!";
  const statistic = scalarText(args[3]!).toUpperCase();
  const model = etsModel([args[0]!, args[1]!, args[2]!], false);
  if (!isEtsModel(model)) return model;
  const actual = sortedTimeline(args[1]!, args[2]!).map(([, value]) => value);
  const meanSquare =
    model.residuals.reduce((total, value) => total + value ** 2, 0) /
    Math.max(1, model.residuals.length);
  switch (statistic) {
    case "ALPHA":
      return model.alpha;
    case "BETA":
      return model.beta;
    case "GAMMA":
      return model.gamma;
    case "MASE": {
      const scale =
        actual
          .slice(1)
          .reduce((total, value, index) => total + Math.abs(value - actual[index]!), 0) /
        Math.max(1, actual.length - 1);
      return scale === 0 ? 0 : Math.sqrt(meanSquare) / scale;
    }
    case "RMSE":
      return Math.sqrt(meanSquare);
    case "STEPDETECTED": {
      const threshold = 2 * Math.sqrt(meanSquare);
      return model.residuals.some((residual) => Math.abs(residual) > threshold) ? 1 : 0;
    }
    default:
      return "#NUM!";
  }
}

function inverseNormal(probability: number): number {
  if (probability <= 0 || probability >= 1) throw new FormulaValueError("#NUM!");
  const rational = probability - 0.5;
  if (Math.abs(rational) < 0.42) {
    const input = rational * rational;
    return (
      (rational *
        (((25.44106049637 * input + 4.284666683759) * input + 0.248571446382) * input + 1)) /
      ((((21.80581851671 * input + 34.36862574235) * input + 9.807903581192) * input +
        3.207378991734) *
        input +
        1)
    );
  }
  const tail = rational < 0 ? probability : 1 - probability;
  const logarithm = Math.log(-Math.log(tail));
  const output =
    2.822928158603 *
    (1 + 0.0455139640395 * logarithm) *
    Math.exp((1.190501027024 * logarithm - 1.515063048907) * logarithm - 0.7069293429659);
  return rational < 0 ? -output : output;
}

function lineStatistics(args: Arguments, logarithmic: boolean): FormulaValue {
  if (args.length < 1) return "#VALUE!";
  const knownY = args[0]!;
  const knownX: FormulaValue = args[1] ?? [
    Array.from({ length: isMatrix(knownY) ? knownY.flat().length : 1 }, (_, index) => index + 1),
  ];
  const observations = regressionObservations(knownY, knownX);
  if (observations === undefined || observations.response.length < 2) return "#REF!";
  const includeConstant = args[2] === undefined ? true : scalarBoolean(args[2]);
  const responseValues = logarithmic
    ? observations.response.map((value) => {
        if (value <= 0) throw new FormulaValueError("#NUM!");
        return Math.log(value);
      })
    : observations.response;
  try {
    const returnStatistics = args[3] !== undefined && scalarBoolean(args[3]);
    const statistics = fitLinearRegression(
      responseValues,
      observations.predictors,
      includeConstant,
      returnStatistics,
    );
    const transformCoefficient = (value: number) => (logarithmic ? Math.exp(value) : value);
    const coefficients = [...statistics.coefficients].reverse().map(transformCoefficient);
    if (args[3] === undefined || !scalarBoolean(args[3])) return [coefficients];
    return [
      coefficients,
      [...statistics.standardErrors].reverse(),
      [statistics.rSquared, statistics.standardErrorY],
      [statistics.fStatistic, statistics.degreesOfFreedom],
      [statistics.regressionSumSquares, statistics.residualSumSquares],
    ];
  } catch (error) {
    if (error instanceof FormulaValueError) return error.value;
    return "#VALUE!";
  }
}

function trendValues(args: Arguments, logarithmic: boolean): FormulaValue {
  if (args.length < 1) return "#VALUE!";
  const knownY = args[0]!;
  const knownX: FormulaValue = args[1] ?? [
    Array.from({ length: isMatrix(knownY) ? knownY.flat().length : 1 }, (_, index) => index + 1),
  ];
  const regression = regressionObservations(knownY, knownX);
  if (regression === undefined || regression.response.length < 2) return "#REF!";
  const includeConstant = args[3] === undefined ? true : scalarBoolean(args[3]);
  const transformedResponse = logarithmic
    ? regression.response.map((value) => {
        if (value <= 0) throw new FormulaValueError("#NUM!");
        return Math.log(value);
      })
    : regression.response;
  const statistics = fitLinearRegression(
    transformedResponse,
    regression.predictors,
    includeConstant,
    false,
  );
  const newX: FormulaValue = args[2] ?? knownX;
  const predictions = predictionRows(newX, regression.predictors[0]?.length ?? 0);
  if (predictions === undefined) return "#REF!";
  return predictLinearValues(statistics, predictions, logarithmic);
}

function requireArgumentCount(args: Arguments, minimum: number, maximum: number): FormulaValue {
  return args.length < minimum || args.length > maximum ? "#VALUE!" : 0;
}

export const forecastFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  ["FVSCHEDULE", ([presentValue, schedules]) => schedule(presentValue, schedules)],
  ["FREQUENCY", ([data, bins]) => frequency(data, bins)],
  [
    "SERIESSUM",
    ([input, exponent, step, coefficients]) => seriesSum(input, exponent!, step!, coefficients!),
  ],
  ["FORECAST.LINEAR", (args) => forecastLinear(args)],
  ["FORECASTLINEAR", (args) => forecastLinear(args)],
  ["FORECAST.ETS", (args) => exponentialSmoothingForecast(args, false)],
  ["FORECAST.ETS.ADD", (args) => exponentialSmoothingForecast(args, false)],
  ["FORECAST.ETS.MULT", (args) => exponentialSmoothingForecast(args, true)],
  ["FORECASTETS", (args) => exponentialSmoothingForecast(args, false)],
  ["FORECAST.ETS.CONFINT", (args) => predictionInterval(args, false)],
  ["FORECAST.ETS.PI.ADD", (args) => predictionInterval(args, false)],
  ["FORECAST.ETS.PI.MULT", (args) => predictionInterval(args, true)],
  ["FORECASTETSCONFINT", (args) => predictionInterval(args, false)],
  [
    "FORECAST.ETS.SEASONALITY",
    ([values, timeline, mode]) => {
      const pairs = sortedTimeline(values!, timeline!);
      return seasonalPeriod(pairs, mode);
    },
  ],
  [
    "FORECASTETSSEASONALITY",
    ([values, timeline, mode]) => {
      if (requireArgumentCount([values!, timeline!, mode ?? true], 2, 3)) return "#VALUE!";
      return seasonalPeriod(sortedTimeline(values!, timeline!), mode);
    },
  ],
  ["FORECAST.ETS.STAT", (args) => smoothingStatistic(args)],
  ["FORECAST.ETS.STAT.ADD", (args) => smoothingStatistic(args)],
  ["FORECAST.ETS.STAT.MULT", (args) => smoothingStatistic(args)],
  ["FORECASTETSSTAT", (args) => smoothingStatistic(args)],
  ["LINEST", (args) => lineStatistics(args, false)],
  ["LOGEST", (args) => lineStatistics(args, true)],
  ["TREND", (args) => trendValues(args, false)],
]);

function forecastLinear(args: Arguments): FormulaValue {
  if (args.length !== 3) return "#VALUE!";
  const pairs = pairedNumbers(args[2]!, args[1]!);
  const line = linearRegression(pairs);
  return line.intercept + line.slope * scalarNumber(args[0]!);
}
