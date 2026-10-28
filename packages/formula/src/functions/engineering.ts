import { FormulaValueError } from "../errors";
import type { FormulaValue } from "../value";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { scalarNumber, scalarText } from "./utils";

export interface ComplexNumber {
  real: number;
  imaginary: number;
}

interface BaseFormat {
  radix: number;
  digits: number;
}

interface LinearUnit {
  kind: "linear";
  factor: number;
  metric: boolean;
}

interface TemperatureUnit {
  kind: "temperature";
  factor: number;
  fromKelvin: (value: number) => number;
  toKelvin: (value: number) => number;
}

type ConvertUnit = LinearUnit | TemperatureUnit;

function isLinearUnit(unit: ConvertUnit): unit is LinearUnit {
  return unit.kind === "linear";
}

const baseFormats = new Map<string, BaseFormat>([
  ["binary", { radix: 2, digits: 10 }],
  ["octal", { radix: 8, digits: 10 }],
  ["hexadecimal", { radix: 16, digits: 10 }],
]);

function requireBaseFormat(name: string): BaseFormat {
  const format = baseFormats.get(name);
  if (!format) throw new FormulaValueError("#NUM!");
  return format;
}

function requireInteger(value: number): number {
  if (!Number.isInteger(value)) throw new FormulaValueError("#NUM!");
  return value;
}

function optionalPlaces(value: FormulaValue | undefined): number | undefined {
  if (value === undefined) return undefined;
  return requireInteger(Math.trunc(scalarNumber(value)));
}

function parseBaseNumber(text: string, format: BaseFormat): number {
  const digits = text.trim();
  if (digits.length === 0 || digits.length > format.digits) throw new FormulaValueError("#NUM!");
  let output = 0;
  for (const character of digits.toLowerCase()) {
    const codePoint = character.codePointAt(0) ?? 0;
    const digit = codePoint >= 97 ? codePoint - 87 : codePoint - 48;
    if (digit < 0 || digit >= format.radix) throw new FormulaValueError("#NUM!");
    output = output * format.radix + digit;
  }
  if (digits.length === format.digits && output >= format.radix ** (format.digits - 1))
    return output - format.radix ** format.digits;
  return output;
}

function decimalToBase(value: number, format: BaseFormat, places: number | undefined): string {
  const limit = format.radix ** (format.digits - 1);
  if (value < -limit || value >= limit) throw new FormulaValueError("#NUM!");
  let digits =
    value < 0
      ? (format.radix ** format.digits + value).toString(format.radix)
      : value.toString(format.radix);
  const width = places ?? digits.length;
  if (width < 1 || width > format.digits || width < digits.length)
    throw new FormulaValueError("#NUM!");
  while (digits.length < width) digits = `0${digits}`;
  return digits.toUpperCase();
}

function baseToDecimal(input: string, format: BaseFormat): number {
  return parseBaseNumber(scalarText(input), format);
}

function decimalToBaseFunction(name: string): FormulaFunctionImplementation {
  const format = requireBaseFormat(name);
  return ([value, places]) =>
    decimalToBase(requireInteger(scalarNumber(value)), format, optionalPlaces(places));
}

function crossBaseFunction(sourceName: string, targetName: string): FormulaFunctionImplementation {
  const source = requireBaseFormat(sourceName);
  const target = requireBaseFormat(targetName);
  return ([value, places]) =>
    decimalToBase(
      requireInteger(baseToDecimal(scalarText(value), source)),
      target,
      optionalPlaces(places),
    );
}

function harmonic(value: number): number {
  let output = 0;
  for (let index = 1; index <= value; index += 1) output += 1 / index;
  return output;
}

function logGamma(value: number): number {
  if (value < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * value)) - logGamma(1 - value);
  const coefficients = [
    0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313,
    -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6,
    1.5056327351493116e-7,
  ];
  value -= 1;
  let sum = coefficients[0]!;
  for (let index = 1; index < coefficients.length; index += 1)
    sum += coefficients[index]! / (value + index);
  const shifted = value + 7.5;
  return 0.5 * Math.log(2 * Math.PI) + (value + 0.5) * Math.log(shifted) - shifted + Math.log(sum);
}

function logFactorial(value: number): number {
  return logGamma(value + 1);
}

function factorial(value: number): number {
  return Math.exp(logFactorial(value));
}

function besselINonnegative(order: number, value: number): number {
  if (order > 10_000) throw new FormulaValueError("#NUM!");
  let term = Math.exp(order * Math.log(value / 2) - logFactorial(order));
  let output = term;
  for (let index = 1; index < 1_000; index += 1) {
    term *= (value * value) / (4 * index * (order + index));
    output += term;
    if (term <= output * Number.EPSILON) break;
  }
  return output;
}

function besselJNonnegative(order: number, value: number): number {
  if (order > 10_000) throw new FormulaValueError("#NUM!");
  if (value === 0) return order === 0 ? 1 : 0;
  if (value <= 8) {
    let term = Math.exp(order * Math.log(value / 2) - logFactorial(order));
    let output = term;
    for (let index = 1; index < 1_000; index += 1) {
      term *= (-value * value) / (4 * index * (order + index));
      output += term;
      if (Math.abs(term) <= Math.abs(output) * Number.EPSILON) break;
    }
    return output;
  }
  const start = order + Math.ceil(Math.sqrt(40 * order)) + 20;
  let next = 0;
  let current = 1e-100;
  let evenSum = 0;
  let desired = 0;
  for (let index = start; index >= 1; index -= 1) {
    const previous = (2 * index * current) / value - next;
    if (index === order) desired = previous;
    if (index % 2 === 0) evenSum += 2 * previous;
    next = current;
    current = previous;
  }
  evenSum += current;
  return desired / evenSum;
}

function besselKZeroOne(value: number): [number, number] {
  const quarter = (value * value) / 4;
  let coefficient = 1;
  let modifiedZero = 1;
  let derivativeSum = 0;
  let harmonicSum = 0;
  for (let index = 1; index < 1_000; index += 1) {
    coefficient *= quarter / (index * index);
    modifiedZero += coefficient;
    derivativeSum += index * coefficient;
    harmonicSum += harmonic(index) * coefficient;
    if (coefficient <= modifiedZero * Number.EPSILON) break;
  }
  const logarithm = Math.log(value / 2) + 0.5772156649015329;
  const secondKindZero = -logarithm * modifiedZero + harmonicSum;
  const secondKindOne =
    modifiedZero / value +
    (2 / value) * (logarithm * derivativeSum - derivativeHarmonicSum(derivativeSum, harmonicSum));
  return [secondKindZero, secondKindOne];
}

function derivativeHarmonicSum(derivativeSum: number, harmonicSum: number): number {
  return derivativeSum - harmonicSum;
}

function besselKNonnegative(order: number, value: number): number {
  if (value <= 0) throw new FormulaValueError("#NUM!");
  if (order > 10_000) throw new FormulaValueError("#NUM!");
  const [zero, one] = besselKZeroOne(value);
  if (order === 0) return zero;
  let previous = zero;
  let current = one;
  for (let index = 1; index < order; index += 1) {
    const next = previous + ((2 * index) / value) * current;
    previous = current;
    current = next;
  }
  return current;
}

function besselYNonnegative(order: number, value: number): number {
  if (value <= 0) throw new FormulaValueError("#NUM!");
  if (order > 10_000) throw new FormulaValueError("#NUM!");
  if (value > 8) {
    const mu = 4 * order * order;
    let pTerm = 1;
    let pSeries = 1;
    let qTerm = 1;
    let qSeries = (mu - 1) / value;
    for (let index = 1; index < 8; index += 1) {
      let product = 1;
      for (let inner = 0; inner < index; inner += 1) product *= mu - (2 * inner + 1) ** 2;
      pTerm *= -product / ((2 * index) ** 2 * value ** 2);
      pSeries += pTerm;
      qTerm *= -product / (factorial(2 * index + 1) * value ** 2);
      qSeries += qTerm * (mu - 1) ** (index === 1 ? 0 : 1);
    }
    const phase = value - (Math.PI / 2) * order - Math.PI / 4;
    const factor = Math.sqrt(2 / (Math.PI * value));
    return factor * (Math.sin(phase) * pSeries + Math.cos(phase) * qSeries);
  }
  const quarter = (value * value) / 4;
  let coefficient = 1;
  let besselZero = 1;
  let zeroHarmonic = 0;
  let derivativeHarmonic = 0;
  for (let index = 1; index < 1_000; index += 1) {
    coefficient *= -quarter / (index * index);
    besselZero += coefficient;
    const magnitude = Math.abs(coefficient);
    zeroHarmonic += harmonic(index) * magnitude;
    derivativeHarmonic += index * harmonic(index) * magnitude;
    if (magnitude <= besselZero * Number.EPSILON) break;
  }
  let besselOne = Math.exp(Math.log(value / 2) - logFactorial(1));
  let oneDerivative = 1;
  let oneHarmonic = 0;
  let current = 1;
  for (let index = 1; index < 1_000; index += 1) {
    current *= -quarter / (index * (index + 1));
    besselOne += current;
    const magnitude = Math.abs(current);
    oneDerivative += (index + 1) * magnitude;
    oneHarmonic += index * harmonic(index) * magnitude;
    if (magnitude <= Math.abs(besselOne) * Number.EPSILON) break;
  }
  const logarithm = Math.log(value / 2) + 0.5772156649015329;
  const secondKindZero = (2 / Math.PI) * (logarithm * besselZero + zeroHarmonic);
  const secondKindOne =
    (2 / Math.PI) * (logarithm * besselOne - besselZero / value - (2 / value) * derivativeHarmonic);
  if (order === 0) return secondKindZero;
  if (order === 1) return secondKindOne;
  let previous = secondKindZero;
  let currentOrder = secondKindOne;
  for (let index = 1; index < order; index += 1) {
    const next = ((2 * index) / value) * currentOrder - previous;
    previous = currentOrder;
    currentOrder = next;
  }
  return currentOrder;
}

function besselIFunction(args: FormulaValue[]): number {
  const order = Math.trunc(scalarNumber(args[1] ?? 0));
  if (order < 0) throw new FormulaValueError("#NUM!");
  return besselINonnegative(order, scalarNumber(args[0] ?? 0));
}

function besselJFunction(args: FormulaValue[]): number {
  const order = Math.trunc(scalarNumber(args[1] ?? 0));
  const value = scalarNumber(args[0] ?? 0);
  const result = besselJNonnegative(Math.abs(order), Math.abs(value));
  return (order < 0 ? -1 : 1) * (value < 0 && order % 2 !== 0 ? -1 : 1) * result;
}

function besselKFunction(args: FormulaValue[]): number {
  const order = Math.trunc(scalarNumber(args[1] ?? 0));
  if (order < 0) throw new FormulaValueError("#NUM!");
  return besselKNonnegative(order, scalarNumber(args[0] ?? 0));
}

function besselYFunction(args: FormulaValue[]): number {
  const order = Math.trunc(scalarNumber(args[1] ?? 0));
  if (order < 0) throw new FormulaValueError("#NUM!");
  return besselYNonnegative(order, scalarNumber(args[0] ?? 0));
}

function erfcPositive(value: number): number {
  if (value < 3) {
    let term = value;
    let output = value;
    for (let index = 1; index < 1_000; index += 1) {
      term *= (-value * value) / index;
      const contribution = term / (2 * index + 1);
      output += contribution;
      if (Math.abs(contribution) <= Math.abs(output) * Number.EPSILON) break;
    }
    return 1 - (2 / Math.sqrt(Math.PI)) * output;
  }
  let result = value;
  let numerator = value;
  let denominator = 0;
  for (let index = 1; index < 1_000; index += 1) {
    const coefficient = index / 2;
    denominator = value + coefficient * denominator;
    if (Math.abs(denominator) < 1e-300) denominator = 1e-300;
    denominator = 1 / denominator;
    numerator = value + coefficient / numerator;
    const delta = numerator * denominator;
    result *= delta;
    if (Math.abs(delta - 1) <= Number.EPSILON) break;
  }
  return Math.exp(-value * value) / (Math.sqrt(Math.PI) * result);
}

function erf(value: number): number {
  if (value < 0) return -erf(-value);
  return 1 - erfcPositive(value);
}

export function parseComplex(value: FormulaValue): ComplexNumber {
  const text = scalarText(value).replace(/\s+/gu, "");
  if (!text) throw new FormulaValueError("#NUM!");
  const boundaries = [0];
  for (let index = 1; index < text.length; index += 1) {
    const character = text[index]!;
    if ((character === "+" || character === "-") && text[index - 1]?.toLowerCase() !== "e")
      boundaries.push(index);
  }
  boundaries.push(text.length);
  let real = 0;
  let imaginary = 0;
  let found = false;
  for (let index = 0; index < boundaries.length - 1; index += 1) {
    const start = boundaries[index]!;
    const end = boundaries[index + 1]!;
    if (start === end) continue;
    const segment = text.slice(start, end);
    found = true;
    const unit = segment[segment.length - 1]?.toLowerCase();
    if (unit === "i" || unit === "j") {
      const coefficient = segment.slice(0, -1);
      const number =
        coefficient === "" || coefficient === "+"
          ? 1
          : coefficient === "-"
            ? -1
            : Number(coefficient);
      if (!Number.isFinite(number)) throw new FormulaValueError("#NUM!");
      imaginary += number;
    } else {
      const number = Number(segment);
      if (!Number.isFinite(number)) throw new FormulaValueError("#NUM!");
      real += number;
    }
  }
  if (!found) throw new FormulaValueError("#NUM!");
  return {
    real: Object.is(real, -0) ? 0 : real,
    imaginary: Object.is(imaginary, -0) ? 0 : imaginary,
  };
}

export function formatComplex(value: ComplexNumber, suffix = "i"): string {
  if (value.imaginary === 0) return `${value.real}`;
  const coefficient = Math.abs(value.imaginary) === 1 ? "" : `${Math.abs(value.imaginary)}`;
  if (value.real === 0) return `${value.imaginary < 0 ? "-" : ""}${coefficient}${suffix}`;
  return `${value.real}${value.imaginary < 0 ? "-" : "+"}${coefficient}${suffix}`;
}

function addComplex(left: ComplexNumber, right: ComplexNumber): ComplexNumber {
  return { real: left.real + right.real, imaginary: left.imaginary + right.imaginary };
}

function subtractComplex(left: ComplexNumber, right: ComplexNumber): ComplexNumber {
  return { real: left.real - right.real, imaginary: left.imaginary - right.imaginary };
}

function multiplyComplex(left: ComplexNumber, right: ComplexNumber): ComplexNumber {
  return {
    real: left.real * right.real - left.imaginary * right.imaginary,
    imaginary: left.real * right.imaginary + left.imaginary * right.real,
  };
}

function divideComplex(left: ComplexNumber, right: ComplexNumber): ComplexNumber {
  const denominator = right.real * right.real + right.imaginary * right.imaginary;
  if (denominator === 0) throw new FormulaValueError("#NUM!");
  return {
    real: (left.real * right.real + left.imaginary * right.imaginary) / denominator,
    imaginary: (left.imaginary * right.real - left.real * right.imaginary) / denominator,
  };
}

function complexLogarithm(value: ComplexNumber): ComplexNumber {
  if (value.real === 0 && value.imaginary === 0) throw new FormulaValueError("#NUM!");
  return {
    real: Math.log(Math.hypot(value.real, value.imaginary)),
    imaginary: Math.atan2(value.imaginary, value.real),
  };
}

function exponentialComplex(value: ComplexNumber): ComplexNumber {
  const magnitude = Math.exp(value.real);
  return {
    real: magnitude * Math.cos(value.imaginary),
    imaginary: magnitude * Math.sin(value.imaginary),
  };
}

function formatComplexResult(
  implementation: (value: ComplexNumber) => ComplexNumber,
): FormulaFunctionImplementation {
  return ([value]) => formatComplex(implementation(parseComplex(value)));
}

function temperatureUnit(
  fromKelvin: (value: number) => number,
  toKelvin: (value: number) => number,
): TemperatureUnit {
  return { kind: "temperature", factor: 1, fromKelvin, toKelvin };
}

function linearUnit(factor: number, metric = false): LinearUnit {
  return { kind: "linear", factor, metric };
}

const convertUnits = new Map<string, ConvertUnit>([
  ["g", linearUnit(0.001, true)],
  ["sg", linearUnit(14.5939029372)],
  ["lbm", linearUnit(0.45359237)],
  ["u", linearUnit(1.6605390666e-27)],
  ["ozm", linearUnit(0.028349523125)],
  ["stone", linearUnit(6.35029318)],
  ["grain", linearUnit(6.479891e-5)],
  ["cwt", linearUnit(45.359237)],
  ["shweight", linearUnit(45.359237)],
  ["uk_cwt", linearUnit(50.80234544)],
  ["ton", linearUnit(907.18474)],
  ["uk_ton", linearUnit(1016.0469088)],
  ["LTON", linearUnit(1016.0469088)],
  ["STON", linearUnit(907.18474)],
  ["m", linearUnit(1, true)],
  ["mi", linearUnit(1609.344)],
  ["survey_mi", linearUnit(1609.3472186944373)],
  ["Nmi", linearUnit(1852)],
  ["in", linearUnit(0.0254)],
  ["survey_in", linearUnit(0.0254000508001016)],
  ["ft", linearUnit(0.3048)],
  ["survey_ft", linearUnit(0.3048006096012192)],
  ["yd", linearUnit(0.9144)],
  ["survey_yd", linearUnit(0.9144018288036576)],
  ["ang", linearUnit(1e-10)],
  ["Pica", linearUnit(0.0254 / 72)],
  ["pica", linearUnit(0.0254 / 72)],
  ["Picapt", linearUnit(0.0254 / 6)],
  ["picapt", linearUnit(0.0254 / 6)],
  ["ell", linearUnit(1.143)],
  ["ly", linearUnit(9.4607304725808e15)],
  ["parsec", linearUnit(Number("3.0856775814913673e+16"))],
  ["yr", linearUnit(31557600)],
  ["day", linearUnit(86400)],
  ["hr", linearUnit(3600)],
  ["mn", linearUnit(60)],
  ["sec", linearUnit(1, true)],
  ["Pa", linearUnit(1, true)],
  ["p", linearUnit(1, true)],
  ["atm", linearUnit(101325)],
  ["at", linearUnit(98066.5)],
  ["mmHg", linearUnit(133.322387415)],
  ["psi", linearUnit(6894.757293168)],
  ["Torr", linearUnit(101325 / 760)],
  ["N", linearUnit(1, true)],
  ["dyn", linearUnit(1e-5, true)],
  ["lbf", linearUnit(4.4482216152605)],
  ["pond", linearUnit(0.00980665, true)],
  ["J", linearUnit(1, true)],
  ["e", linearUnit(1e-7, true)],
  ["c", linearUnit(4.184, true)],
  ["cal", linearUnit(4.184, true)],
  ["eV", linearUnit(1.602176634e-19, true)],
  ["HPh", linearUnit(2684519.537696)],
  ["Wh", linearUnit(3600, true)],
  ["flb", linearUnit(1.3558179483314)],
  ["BTU", linearUnit(1055.05585262)],
  ["HP", linearUnit(745.69987158227)],
  ["PS", linearUnit(735.49875, true)],
  ["W", linearUnit(1, true)],
  ["T", linearUnit(1, true)],
  ["ga", linearUnit(1e-4, true)],
  ["tsp", linearUnit(4.92892159375e-6)],
  ["tbs", linearUnit(1.478676478125e-5)],
  ["oz", linearUnit(2.95735295625e-5)],
  ["cup", linearUnit(0.0002365882365)],
  ["pt", linearUnit(0.000473176473)],
  ["us_pt", linearUnit(0.000473176473)],
  ["uk_pt", linearUnit(0.00056826125)],
  ["qt", linearUnit(0.000946352946)],
  ["gal", linearUnit(0.003785411784)],
  ["l", linearUnit(0.001, true)],
  ["lt", linearUnit(0.001, true)],
  ["ha", linearUnit(10000)],
  [
    "C",
    temperatureUnit(
      (value) => value - 273.15,
      (value) => value + 273.15,
    ),
  ],
  [
    "cel",
    temperatureUnit(
      (value) => value - 273.15,
      (value) => value + 273.15,
    ),
  ],
  [
    "F",
    temperatureUnit(
      (value) => ((value - 273.15) * 9) / 5 + 32,
      (value) => ((value - 32) * 5) / 9 + 273.15,
    ),
  ],
  [
    "fah",
    temperatureUnit(
      (value) => ((value - 273.15) * 9) / 5 + 32,
      (value) => ((value - 32) * 5) / 9 + 273.15,
    ),
  ],
  [
    "K",
    temperatureUnit(
      (value) => value,
      (value) => value,
    ),
  ],
  [
    "kel",
    temperatureUnit(
      (value) => value,
      (value) => value,
    ),
  ],
  [
    "Rank",
    temperatureUnit(
      (value) => (value * 5) / 9,
      (value) => (value * 9) / 5,
    ),
  ],
  [
    "Reaum",
    temperatureUnit(
      (value) => (value * 5) / 4,
      (value) => (value * 4) / 5,
    ),
  ],
]);

const metricPrefixes = new Map<string, number>([
  ["Y", 1e24],
  ["Z", 1e21],
  ["E", 1e18],
  ["P", 1e15],
  ["T", 1e12],
  ["G", 1e9],
  ["M", 1e6],
  ["k", 1e3],
  ["h", 1e2],
  ["da", 1e1],
  ["d", 1e-1],
  ["c", 1e-2],
  ["m", 1e-3],
  ["u", 1e-6],
  ["µ", 1e-6],
  ["n", 1e-9],
  ["p", 1e-12],
  ["f", 1e-15],
  ["a", 1e-18],
  ["z", 1e-21],
  ["y", 1e-24],
]);

function lookupConvertUnit(unit: string): ConvertUnit | undefined {
  const exact = convertUnits.get(unit);
  if (exact) return exact;
  for (const [baseUnit, definition] of convertUnits) {
    if (!isLinearUnit(definition) || !definition.metric) continue;
    if (!unit.endsWith(baseUnit)) continue;
    const prefix = unit.slice(0, unit.length - baseUnit.length);
    const factor = metricPrefixes.get(prefix);
    if (factor === undefined) continue;
    return {
      kind: "linear",
      factor: definition.factor * factor,
      metric: definition.metric,
    };
  }
  return undefined;
}

function convertTemperature(
  value: number,
  source: TemperatureUnit,
  target: TemperatureUnit,
): number {
  return target.fromKelvin(source.toKelvin(value));
}

function convertNumber(value: number, sourceUnit: string, targetUnit: string): number {
  const source = lookupConvertUnit(sourceUnit);
  const target = lookupConvertUnit(targetUnit);
  if (!source || !target || source.kind !== target.kind) throw new FormulaValueError("#NUM!");
  if (source.kind === "temperature" && target.kind === "temperature")
    return convertTemperature(value, source, target);
  if (source.kind === "linear" && target.kind === "linear")
    return (value * source.factor) / target.factor;
  throw new FormulaValueError("#NUM!");
}

function deltaFunction([value, step]: FormulaValue[]): number {
  return scalarNumber(value ?? 0) === scalarNumber(step ?? 0) ? 1 : 0;
}

function greaterOrEqualStep([value, step]: FormulaValue[]): number {
  return scalarNumber(value ?? 0) >= scalarNumber(step ?? 0) ? 1 : 0;
}

export const engineeringFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  ["BIN2DEC", ([value]) => baseToDecimal(scalarText(value), requireBaseFormat("binary"))],
  ["OCT2DEC", ([value]) => baseToDecimal(scalarText(value), requireBaseFormat("octal"))],
  ["HEX2DEC", ([value]) => baseToDecimal(scalarText(value), requireBaseFormat("hexadecimal"))],
  ["DEC2BIN", decimalToBaseFunction("binary")],
  ["DEC2OCT", decimalToBaseFunction("octal")],
  ["DEC2HEX", decimalToBaseFunction("hexadecimal")],
  ["BIN2OCT", crossBaseFunction("binary", "octal")],
  ["BIN2HEX", crossBaseFunction("binary", "hexadecimal")],
  ["OCT2BIN", crossBaseFunction("octal", "binary")],
  ["OCT2HEX", crossBaseFunction("octal", "hexadecimal")],
  ["HEX2BIN", crossBaseFunction("hexadecimal", "binary")],
  ["HEX2OCT", crossBaseFunction("hexadecimal", "octal")],
  [
    "COMPLEX",
    ([real, imaginary, suffix]) => {
      const unit = suffix === undefined ? "i" : scalarText(suffix);
      if (unit !== "i" && unit !== "j") throw new FormulaValueError("#NUM!");
      return formatComplex({ real: scalarNumber(real), imaginary: scalarNumber(imaginary) }, unit);
    },
  ],
  ["IMABS", ([value]) => Math.hypot(parseComplex(value).real, parseComplex(value).imaginary)],
  ["IMAGINARY", ([value]) => parseComplex(value).imaginary],
  ["IMREAL", ([value]) => parseComplex(value).real],
  [
    "IMCONJUGATE",
    formatComplexResult((value) => ({ real: value.real, imaginary: -value.imaginary })),
  ],
  [
    "IMARGUMENT",
    ([value]) => {
      const complex = parseComplex(value);
      if (complex.real === 0 && complex.imaginary === 0) throw new FormulaValueError("#NUM!");
      return Math.atan2(complex.imaginary, complex.real);
    },
  ],
  [
    "IMSUM",
    (args) =>
      formatComplex(
        args.reduce<ComplexNumber>((left, right) => addComplex(left, parseComplex(right)), {
          real: 0,
          imaginary: 0,
        }),
      ),
  ],
  [
    "IMSUB",
    ([left, right]) => formatComplex(subtractComplex(parseComplex(left), parseComplex(right))),
  ],
  [
    "IMPRODUCT",
    (args) =>
      formatComplex(
        args.reduce<ComplexNumber>((left, right) => multiplyComplex(left, parseComplex(right)), {
          real: 1,
          imaginary: 0,
        }),
      ),
  ],
  [
    "IMDIV",
    ([left, right]) => formatComplex(divideComplex(parseComplex(left), parseComplex(right))),
  ],
  ["IMEXP", formatComplexResult(exponentialComplex)],
  ["IMLN", formatComplexResult(complexLogarithm)],
  [
    "IMLOG10",
    formatComplexResult((value) => ({
      real: complexLogarithm(value).real / Math.LN10,
      imaginary: complexLogarithm(value).imaginary / Math.LN10,
    })),
  ],
  [
    "IMLOG2",
    formatComplexResult((value) => ({
      real: complexLogarithm(value).real / Math.LN2,
      imaginary: complexLogarithm(value).imaginary / Math.LN2,
    })),
  ],
  [
    "IMPOWER",
    ([value, exponent]) => {
      const base = parseComplex(value);
      const power = parseComplex(exponent);
      if (base.real === 0 && base.imaginary === 0) {
        if (power.real === 0 && power.imaginary === 0) throw new FormulaValueError("#NUM!");
        return formatComplex({ real: 0, imaginary: 0 });
      }
      return formatComplex(exponentialComplex(multiplyComplex(power, complexLogarithm(base))));
    },
  ],
  [
    "IMSQRT",
    formatComplexResult((value) => {
      const magnitude = Math.sqrt(Math.hypot(value.real, value.imaginary));
      const angle = Math.atan2(value.imaginary, value.real) / 2;
      return { real: magnitude * Math.cos(angle), imaginary: magnitude * Math.sin(angle) };
    }),
  ],
  [
    "IMSIN",
    formatComplexResult((value) => ({
      real: Math.sin(value.real) * Math.cosh(value.imaginary),
      imaginary: Math.cos(value.real) * Math.sinh(value.imaginary),
    })),
  ],
  [
    "IMCOS",
    formatComplexResult((value) => ({
      real: Math.cos(value.real) * Math.cosh(value.imaginary),
      imaginary: -Math.sin(value.real) * Math.sinh(value.imaginary),
    })),
  ],
  [
    "IMTAN",
    formatComplexResult((value) => {
      const denominator = Math.cos(2 * value.real) + Math.cosh(2 * value.imaginary);
      return {
        real: Math.sin(2 * value.real) / denominator,
        imaginary: Math.sinh(2 * value.imaginary) / denominator,
      };
    }),
  ],
  ["BESSELI", (args) => besselIFunction(args)],
  ["BESSELJ", (args) => besselJFunction(args)],
  ["BESSELK", (args) => besselKFunction(args)],
  ["BESSELY", (args) => besselYFunction(args)],
  [
    "ERF",
    ([lower, upper]) => {
      const lowerValue = erf(scalarNumber(lower));
      return upper === undefined ? lowerValue : erf(scalarNumber(upper)) - lowerValue;
    },
  ],
  ["ERF.PRECISE", ([value]) => erf(scalarNumber(value))],
  [
    "ERFC",
    ([value]) => {
      const number = scalarNumber(value);
      if (number < 0) throw new FormulaValueError("#NUM!");
      return erfcPositive(number);
    },
  ],
  [
    "ERFC.PRECISE",
    ([value]) => erfcPositive(-scalarNumber(value) === 0 ? 0 : -scalarNumber(value)),
  ],
  ["DELTA", deltaFunction],
  ["GESTEP", greaterOrEqualStep],
  [
    "CONVERT",
    ([value, source, target]) =>
      convertNumber(scalarNumber(value), scalarText(source), scalarText(target)),
  ],
]);
