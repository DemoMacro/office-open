import { FormulaValueError } from "../errors";

export function requirePositive(value: number): number {
  if (value <= 0 || !Number.isFinite(value)) throw new FormulaValueError("#NUM!");
  return value;
}

export function requireNonNegative(value: number): number {
  if (value < 0 || !Number.isFinite(value)) throw new FormulaValueError("#NUM!");
  return value;
}

export function requireProbability(value: number): number {
  if (value < 0 || value > 1 || !Number.isFinite(value)) throw new FormulaValueError("#NUM!");
  return value;
}

export function erf(value: number): number {
  const sign = Math.sign(value);
  const absolute = Math.abs(value);
  const term = 1 / (1 + absolute / 2);
  const approximation =
    term *
    Math.exp(
      -absolute * absolute -
        1.26551223 +
        term *
          (1.00002368 +
            term *
              (0.37409196 +
                term *
                  (0.09678418 +
                    term *
                      (-0.18628806 +
                        term *
                          (0.27886807 +
                            term *
                              (-1.13520398 +
                                term *
                                  (1.48851587 + term * (-0.82215223 + term * 0.17087277)))))))),
    );
  return sign * (1 - approximation);
}

export function normalCdf(value: number): number {
  return 0.5 * (1 + erf(value / Math.SQRT2));
}

export function normalInverse(probability: number): number {
  requireProbability(probability);
  if (probability === 0.5) return 0;
  if (probability < 0.5) return -inverseNormalTail(1 - probability);
  return inverseNormalTail(probability);
}

function inverseNormalTail(probability: number): number {
  let low = 0;
  let high = 10;
  for (let index = 0; index < 100; index += 1) {
    const middle = (low + high) / 2;
    if (normalCdf(middle) < probability) low = middle;
    else high = middle;
  }
  return (low + high) / 2;
}

export function gammaLn(value: number): number {
  requirePositive(value);
  if (value < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * value)) - gammaLn(1 - value);
  const coefficients = [
    76.18009172947146, -86.5053203294168, 24.01409824083091, -1.231739572450155,
    0.1208650973866179e-2, -0.5395239384953e-5,
  ];
  let series = 1.000000000190015;
  let denominator = value;
  const offset = value + 5.5 - (value + 0.5) * Math.log(value + 5.5);
  for (const coefficient of coefficients) series += coefficient / ++denominator;
  return -offset + Math.log((Math.sqrt(2 * Math.PI) * series) / value);
}

export function gamma(value: number): number {
  if (value < 0.5) return Math.PI / (Math.sin(Math.PI * value) * gamma(1 - value));
  return Math.exp(gammaLn(value));
}

export function lowerGamma(shape: number, value: number): number {
  if (value <= 0) return 0;
  if (shape < 1)
    return lowerGamma(shape + 1, value) / shape + (Math.exp(-value) * value ** shape) / shape;
  let sum = 1 / shape;
  let term = sum;
  for (let index = 1; index < 200; index += 1) {
    term *= value / (shape + index);
    sum += term;
    if (term < sum * 1e-14) break;
  }
  return sum * Math.exp(-value + shape * Math.log(value) - gammaLn(shape));
}

export function betaFunction(left: number, right: number): number {
  return Math.exp(gammaLn(left) + gammaLn(right) - gammaLn(left + right));
}

function continuedFraction(value: number, alpha: number, beta: number): number {
  const maximum = 500;
  const epsilon = 1e-15;
  const minimum = Number.MIN_VALUE;
  const alphaBeta = alpha + beta;
  const alphaPlus = alpha + 1;
  const alphaMinus = alpha - 1;
  let upper = 1;
  let lower = 1 - (alphaBeta * value) / alphaPlus;
  if (Math.abs(lower) < minimum) lower = minimum;
  lower = 1 / lower;
  let result = lower;
  for (let index = 1; index <= maximum; index += 1) {
    const doubled = 2 * index;
    let coefficient =
      (index * (beta - index) * value) / ((alphaMinus + doubled) * (alpha + doubled));
    lower = 1 + coefficient * lower;
    if (Math.abs(lower) < minimum) lower = minimum;
    upper = 1 + coefficient / upper;
    if (Math.abs(upper) < minimum) upper = minimum;
    lower = 1 / lower;
    result *= lower * upper;
    coefficient =
      (-(alpha + index) * (alphaBeta + index) * value) /
      ((alpha + doubled) * (alphaPlus + doubled));
    lower = 1 + coefficient * lower;
    if (Math.abs(lower) < minimum) lower = minimum;
    upper = 1 + coefficient / upper;
    if (Math.abs(upper) < minimum) upper = minimum;
    lower = 1 / lower;
    const delta = lower * upper;
    result *= delta;
    if (Math.abs(delta - 1) < epsilon) break;
  }
  return result;
}

export function incompleteBeta(value: number, alpha: number, beta: number): number {
  if (value <= 0) return 0;
  if (value >= 1) return 1;
  const scale = Math.exp(
    gammaLn(alpha + beta) -
      gammaLn(alpha) -
      gammaLn(beta) +
      alpha * Math.log(value) +
      beta * Math.log(1 - value),
  );
  if (value < (alpha + 1) / (alpha + beta + 2))
    return (scale / alpha) * continuedFraction(value, alpha, beta);
  return 1 - (scale / beta) * continuedFraction(1 - value, beta, alpha);
}

export function inverseIncompleteBeta(probability: number, alpha: number, beta: number): number {
  requireProbability(probability);
  let low = 0;
  let high = 1;
  for (let index = 0; index < 100; index += 1) {
    const middle = (low + high) / 2;
    if (incompleteBeta(middle, alpha, beta) < probability) low = middle;
    else high = middle;
  }
  return (low + high) / 2;
}

export function factorial(value: number): number {
  if (value < 0 || !Number.isInteger(value)) throw new FormulaValueError("#NUM!");
  let output = 1;
  for (let index = 2; index <= value; index += 1) output *= index;
  return output;
}

export function combination(total: number, chosen: number): number {
  if (chosen < 0 || chosen > total) return 0;
  let output = 1;
  for (let index = 1; index <= chosen; index += 1)
    output = (output * (total - chosen + index)) / index;
  return output;
}

export function sumSeries(from: number, to: number, callback: (index: number) => number): number {
  let output = 0;
  for (let index = from; index <= to; index += 1) output += callback(index);
  return output;
}

export function studentTailProbability(value: number, degrees: number, tails: number): number {
  if (degrees <= 0 || !Number.isFinite(value)) throw new FormulaValueError("#NUM!");
  if (tails !== 1 && tails !== 2) throw new FormulaValueError("#NUM!");
  const transformed = degrees / (degrees + value ** 2);
  const probability = incompleteBeta(transformed, degrees / 2, 0.5);
  return tails === 1 ? probability / 2 : probability;
}
