import { FormulaValueError } from "../errors";
import type { FormulaValue } from "../value";
import { matchesCriterion, scalarCriterion } from "./criterion";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { assertFormulaCellCount, numbers, scalarNumber, scalarText, values } from "./utils";

function gcd(left: number, right: number): number {
  left = Math.abs(left);
  right = Math.abs(right);
  while (right) [left, right] = [right, left % right];
  return left;
}

function lcm(left: number, right: number): number {
  if (left === 0 || right === 0) return 0;
  return Math.abs(left * right) / gcd(left, right);
}

function factorial(value: number): number {
  if (value < 0 || !Number.isInteger(value)) throw new FormulaValueError("#NUM!");
  let output = 1;
  for (let index = 2; index <= value; index += 1) output *= index;
  return output;
}

function combination(total: number, chosen: number): number {
  if (
    total < 0 ||
    chosen < 0 ||
    !Number.isInteger(total) ||
    !Number.isInteger(chosen) ||
    chosen > total
  )
    throw new FormulaValueError("#NUM!");
  let output = 1;
  for (let index = 1; index <= chosen; index += 1) output *= (total - chosen + index) / index;
  if (!Number.isFinite(output)) throw new FormulaValueError("#NUM!");
  return output;
}

function matrix(value: FormulaValue): number[][] {
  if (!Array.isArray(value)) return [[scalarNumber(value)]];
  return value.map((row) => row.map(scalarNumber));
}

function integerNumbers(args: Parameters<FormulaFunctionImplementation>[0]): number[] {
  const values = args.flatMap(numbers).map(Math.trunc);
  if (values.some((value) => value < 0)) throw new FormulaValueError("#NUM!");
  return values;
}

export const mathFunctions: FormulaFunctionTable = new Map<string, FormulaFunctionImplementation>([
  ["ABS", ([value]) => Math.abs(scalarNumber(value))],
  ["ACOS", ([value]) => Math.acos(scalarNumber(value))],
  ["ACOSH", ([value]) => Math.acosh(scalarNumber(value))],
  ["ACOT", ([value]) => Math.PI / 2 - Math.atan(scalarNumber(value))],
  ["ACOTH", ([value]) => Math.atanh(1 / scalarNumber(value))],
  ["ARABIC", ([value]) => romanToNumber(scalarText(value))],
  ["ASIN", ([value]) => Math.asin(scalarNumber(value))],
  ["ASINH", ([value]) => Math.asinh(scalarNumber(value))],
  ["ATAN", ([value]) => Math.atan(scalarNumber(value))],
  ["ATAN2", ([x, y]) => Math.atan2(scalarNumber(y), scalarNumber(x))],
  ["ATANH", ([value]) => Math.atanh(scalarNumber(value))],
  [
    "CEILING",
    ([value, significance, mode]) =>
      ceiling(
        scalarNumber(value),
        scalarNumber(significance),
        mode !== undefined && scalarNumber(mode) !== 0 ? "toward-zero" : "away-from-zero",
      ),
  ],
  [
    "CEILING.MATH",
    ([value, significance, mode]) =>
      ceiling(
        scalarNumber(value),
        significance === undefined ? 1 : scalarNumber(significance),
        mode !== undefined && scalarNumber(mode) !== 0 ? "away-from-zero" : "toward-zero",
      ),
  ],
  [
    "CEILING.PRECISE",
    ([value, significance]) =>
      ceilingPrecise(
        scalarNumber(value),
        significance === undefined ? 1 : scalarNumber(significance),
      ),
  ],
  [
    "COMBIN",
    ([count, chosen]) =>
      combination(Math.trunc(scalarNumber(count)), Math.trunc(scalarNumber(chosen))),
  ],
  [
    "COMBINA",
    ([count, chosen]) => {
      const total = Math.trunc(scalarNumber(count));
      const selected = Math.trunc(scalarNumber(chosen));
      if (total < 1) throw new FormulaValueError("#NUM!");
      return combination(total + selected - 1, selected);
    },
  ],
  ["COSH", ([value]) => Math.cosh(scalarNumber(value))],
  ["COS", ([value]) => Math.cos(scalarNumber(value))],
  ["COT", ([value]) => 1 / Math.tan(scalarNumber(value))],
  ["COTH", ([value]) => 1 / Math.tanh(scalarNumber(value))],
  ["CSC", ([value]) => 1 / Math.sin(scalarNumber(value))],
  ["CSCH", ([value]) => 1 / Math.sinh(scalarNumber(value))],
  ["DECIMAL", ([value, radix]) => Number.parseInt(scalarText(value), scalarNumber(radix))],
  ["DEGREES", ([value]) => (scalarNumber(value) * 180) / Math.PI],
  [
    "EVEN",
    ([value]) => {
      const number = scalarNumber(value);
      return number >= 0 ? Math.ceil(number / 2) * 2 : Math.floor(number / 2) * 2;
    },
  ],
  ["EXP", ([value]) => Math.exp(scalarNumber(value))],
  ["FALSE", () => false],
  ["FACT", ([value]) => factorial(Math.trunc(scalarNumber(value)))],
  [
    "FACTDOUBLE",
    ([value]) => {
      let output = 1;
      for (let index = Math.trunc(scalarNumber(value)); index > 1; index -= 2) output *= index;
      return output;
    },
  ],
  [
    "FLOOR",
    ([value, significance, mode]) =>
      floor(
        scalarNumber(value),
        scalarNumber(significance),
        mode !== undefined && scalarNumber(mode) !== 0 ? "toward-zero" : "down",
      ),
  ],
  [
    "FLOOR.MATH",
    ([value, significance, mode]) =>
      floor(
        scalarNumber(value),
        significance === undefined ? 1 : scalarNumber(significance),
        mode === undefined || scalarNumber(mode) >= 0 ? "down" : "toward-zero",
      ),
  ],
  [
    "FLOOR.PRECISE",
    ([value, significance]) =>
      floorPrecise(
        scalarNumber(value),
        significance === undefined ? 1 : scalarNumber(significance),
      ),
  ],
  [
    "GCD",
    (args) => {
      const values = integerNumbers(args);
      if (values.length === 0) throw new FormulaValueError("#NUM!");
      return values.reduce((left, right) => gcd(left, right));
    },
  ],
  ["INT", ([value]) => Math.floor(scalarNumber(value))],
  [
    "LCM",
    (args) => {
      const values = integerNumbers(args);
      if (values.length === 0) throw new FormulaValueError("#NUM!");
      return values.reduce((left, right) => lcm(left, right));
    },
  ],
  ["LN", ([value]) => Math.log(scalarNumber(value))],
  [
    "LOG",
    ([value, base]) =>
      Math.log(scalarNumber(value)) / Math.log(base === undefined ? 10 : scalarNumber(base)),
  ],
  ["LOG10", ([value]) => Math.log10(scalarNumber(value))],
  ["MDETERM", ([value]) => determinant(matrix(value))],
  ["MINVERSE", ([value]) => inverse(matrix(value))],
  ["MMULT", ([left, right]) => multiplyMatrices(matrix(left), matrix(right))],
  [
    "MROUND",
    ([value, multiple]) => {
      const number = scalarNumber(value);
      const divisor = scalarNumber(multiple);
      if (divisor === 0 || (number > 0 && divisor < 0)) throw new FormulaValueError("#NUM!");
      return Math.sign(number) * Math.floor(Math.abs(number / divisor) + 0.5) * Math.abs(divisor);
    },
  ],
  [
    "MOD",
    ([value, divisor]) => {
      const number = scalarNumber(divisor);
      if (number === 0) throw new FormulaValueError("#DIV/0!");
      return scalarNumber(value) - number * Math.floor(scalarNumber(value) / number);
    },
  ],
  [
    "MULTINOMIAL",
    (args) => {
      const total = args.flatMap(numbers).reduce((sum, value) => sum + value, 0);
      let output = factorial(total);
      for (const value of args.flatMap(numbers)) output /= factorial(value);
      return output;
    },
  ],
  [
    "ODD",
    ([value]) => {
      const number = scalarNumber(value);
      if (number === 0) return 1;
      const rounded = number > 0 ? Math.ceil(Math.abs(number)) : -Math.ceil(Math.abs(number));
      return rounded % 2 === 0 ? rounded + Math.sign(number) : rounded;
    },
  ],
  ["PI", () => Math.PI],
  ["POWER", ([value, exponent]) => scalarNumber(value) ** scalarNumber(exponent)],
  ["PRODUCT", (args) => args.flatMap(numbers).reduce((left, right) => left * right, 1)],
  ["QUOTIENT", ([value, divisor]) => Math.trunc(scalarNumber(value) / scalarNumber(divisor))],
  ["RADIANS", ([value]) => (scalarNumber(value) * Math.PI) / 180],
  ["RAND", (_args, context) => context.workbook.random?.() ?? Math.random()],
  [
    "RANDBETWEEN",
    ([minimum, maximum], context) =>
      Math.floor(
        (context.workbook.random?.() ?? Math.random()) *
          (scalarNumber(maximum) - scalarNumber(minimum) + 1),
      ) + scalarNumber(minimum),
  ],
  [
    "ROMAN",
    ([value, form]) => {
      const number = scalarNumber(value);
      if (form === undefined || form === null) return romanForm(number, 0);
      const text = scalarText(form).toUpperCase();
      if (text === "TRUE") return romanForm(number, 0);
      if (text === "FALSE") return romanForm(number, 4);
      const selected = scalarNumber(form);
      if (!Number.isInteger(selected) || selected < 0 || selected > 4)
        throw new FormulaValueError("#VALUE!");
      return romanForm(number, selected);
    },
  ],
  [
    "ROUND",
    ([value, digits]) =>
      round(scalarNumber(value), digits === undefined ? 0 : scalarNumber(digits)),
  ],
  [
    "ROUNDDOWN",
    ([value, digits]) =>
      roundDirection(scalarNumber(value), digits === undefined ? 0 : scalarNumber(digits), "down"),
  ],
  [
    "ROUNDUP",
    ([value, digits]) =>
      roundDirection(scalarNumber(value), digits === undefined ? 0 : scalarNumber(digits), "up"),
  ],
  ["SEC", ([value]) => 1 / Math.cos(scalarNumber(value))],
  ["SECH", ([value]) => 1 / Math.cosh(scalarNumber(value))],
  ["SIGN", ([value]) => Math.sign(scalarNumber(value))],
  ["SIN", ([value]) => Math.sin(scalarNumber(value))],
  ["SINH", ([value]) => Math.sinh(scalarNumber(value))],
  ["SQRT", ([value]) => Math.sqrt(scalarNumber(value))],
  ["SQRTPI", ([value]) => Math.sqrt(Math.PI * scalarNumber(value))],
  ["SUM", (args) => args.flatMap(numbers).reduce((sum, value) => sum + value, 0)],
  ["SUMIF", ([range, criteria, sumRange]) => sumMatching(range, criteria, sumRange ?? range)],
  [
    "SUMIFS",
    (args) => {
      const sumRange = args[0]!;
      let total = 0;
      for (let index = 1; index < args.length; index += 2) {
        if (index + 1 >= args.length) throw new FormulaValueError("#VALUE!");
      }
      for (let rowIndex = 0; rowIndex < values(sumRange).length; rowIndex += 1) {
        let include = true;
        for (let index = 1; index < args.length; index += 2) {
          if (
            !matchesCriterion(
              values(args[index]!)[rowIndex] ?? null,
              scalarCriterion(args[index + 1]!),
            )
          )
            include = false;
        }
        if (include) total += scalarNumber(values(sumRange)[rowIndex] ?? 0);
      }
      return total;
    },
  ],
  ["SUMSQ", (args) => args.flatMap(numbers).reduce((sum, value) => sum + value ** 2, 0)],
  [
    "SUMX2MY2",
    ([left, right]) => zipNumbers(left, right).reduce((sum, [x, y]) => sum + x ** 2 - y ** 2, 0),
  ],
  [
    "SUMX2PY2",
    ([left, right]) => zipNumbers(left, right).reduce((sum, [x, y]) => sum + x ** 2 + y ** 2, 0),
  ],
  [
    "SUMXMY2",
    ([left, right]) => zipNumbers(left, right).reduce((sum, [x, y]) => sum + (x - y) ** 2, 0),
  ],
  [
    "SUMPRODUCT",
    (args) => {
      const matrices = args.map((argument) => (Array.isArray(argument) ? argument : [[argument]]));
      const rows = matrices[0]?.length ?? 0;
      const columns = matrices[0]?.[0]?.length ?? 0;
      if (
        matrices.some(
          (matrix) => matrix.length !== rows || matrix.some((row) => row.length !== columns),
        )
      )
        throw new FormulaValueError("#VALUE!");
      let total = 0;
      for (let row = 0; row < rows; row += 1) {
        for (let column = 0; column < columns; column += 1) {
          const entry = matrices.reduce((product, matrix) => {
            const value = matrix[row]?.[column] ?? null;
            return product * (typeof value === "number" ? value : 0);
          }, 1);
          total += entry;
        }
      }
      return total;
    },
  ],
  ["TAN", ([value]) => Math.tan(scalarNumber(value))],
  ["TANH", ([value]) => Math.tanh(scalarNumber(value))],
  ["TRUNC", ([value, digits]) => truncateNumber([value, digits], "trunc")],
]);

function truncateNumber(args: FormulaValue[], mode: "down" | "up" | "trunc"): number {
  const digits = args[1] === undefined ? 0 : scalarNumber(args[1]);
  const factor = 10 ** digits;
  const value = scalarNumber(args[0]!) * factor;
  const rounded =
    mode === "down" ? Math.floor(value) : mode === "up" ? Math.ceil(value) : Math.trunc(value);
  return rounded / factor;
}

function ceiling(
  value: number,
  significance: number,
  mode: "up" | "toward-zero" | "away-from-zero",
): number {
  if (significance === 0) return 0;
  const quotient = value / significance;
  const rounded =
    mode === "up"
      ? Math.ceil(quotient)
      : mode === "toward-zero"
        ? Math.trunc(quotient)
        : Math.ceil(Math.abs(quotient)) * Math.sign(value || 1);
  return rounded * significance;
}

function floor(value: number, significance: number, mode: "down" | "toward-zero"): number {
  if (significance === 0) return 0;
  const quotient = value / significance;
  return (mode === "down" ? Math.floor(quotient) : Math.trunc(quotient)) * significance;
}

function ceilingPrecise(value: number, significance: number): number {
  const step = Math.abs(significance);
  return step === 0 ? 0 : Math.ceil(value / step) * step;
}

function floorPrecise(value: number, significance: number): number {
  const step = Math.abs(significance);
  return step === 0 ? 0 : Math.floor(value / step) * step;
}

function round(value: number, digits: number): number {
  const factor = 10 ** digits;
  return (Math.sign(value) * Math.round(Math.abs(value) * factor)) / factor;
}

function roundDirection(value: number, digits: number, direction: "down" | "up"): number {
  const factor = 10 ** digits;
  return (
    ((direction === "down"
      ? Math.floor(Math.abs(value) * factor)
      : Math.ceil(Math.abs(value) * factor)) *
      Math.sign(value)) /
    factor
  );
}

function sumMatching(range: FormulaValue, criteria: FormulaValue, sumRange: FormulaValue): number {
  const rangeValues = values(range);
  const sumValues = values(sumRange);
  return rangeValues.reduce(
    (total: number, value, index) =>
      matchesCriterion(value ?? null, scalarCriterion(criteria))
        ? total + scalarNumber(sumValues[index] ?? 0)
        : total,
    0,
  );
}

function zipNumbers(left: FormulaValue, right: FormulaValue): [number, number][] {
  const leftValues = values(left);
  const rightValues = values(right);
  return leftValues.map((value, index) => [
    scalarNumber(value),
    scalarNumber(rightValues[index] ?? 0),
  ]);
}

function determinant(input: number[][]): number {
  const size = input.length;
  if (input.some((row) => row.length !== size)) throw new FormulaValueError("#VALUE!");
  const matrix = input.map((row) => [...row]);
  let output = 1;
  for (let column = 0; column < size; column += 1) {
    let pivotRow = column;
    for (let row = column + 1; row < size; row += 1)
      if (Math.abs(matrix[row]![column]!) > Math.abs(matrix[pivotRow]![column]!)) pivotRow = row;
    const pivot = matrix[pivotRow]![column]!;
    if (pivot === 0) return 0;
    if (pivotRow !== column) {
      [matrix[column], matrix[pivotRow]] = [matrix[pivotRow]!, matrix[column]!];
      output *= -1;
    }
    output *= pivot;
    for (let row = column + 1; row < size; row += 1) {
      const factor = matrix[row]![column]! / pivot;
      for (let value = column; value < size; value += 1)
        matrix[row]![value]! -= factor * matrix[column]![value]!;
    }
  }
  return output;
}

function inverse(input: number[][]): number[][] {
  const size = input.length;
  if (size === 0 || input.some((row) => row.length !== size))
    throw new FormulaValueError("#VALUE!");
  const augmented = input.map((row, rowIndex) => [
    ...row,
    ...Array.from({ length: size }, (_, columnIndex) => (rowIndex === columnIndex ? 1 : 0)),
  ]);
  for (let column = 0; column < size; column += 1) {
    let pivotRow = column;
    for (let row = column + 1; row < size; row += 1)
      if (Math.abs(augmented[row]![column]!) > Math.abs(augmented[pivotRow]![column]!))
        pivotRow = row;
    if (augmented[pivotRow]![column]! === 0) throw new FormulaValueError("#NUM!");
    [augmented[column], augmented[pivotRow]] = [augmented[pivotRow]!, augmented[column]!];
    const pivot = augmented[column]![column]!;
    for (let index = 0; index < size * 2; index += 1) augmented[column]![index]! /= pivot;
    for (let row = 0; row < size; row += 1) {
      if (row === column) continue;
      const factor = augmented[row]![column]!;
      for (let index = 0; index < size * 2; index += 1)
        augmented[row]![index]! -= factor * augmented[column]![index]!;
    }
  }
  return augmented.map((row) => row.slice(size));
}

function multiplyMatrices(left: number[][], right: number[][]): number[][] {
  const columnCount = left[0]?.length ?? 0;
  if (right.length !== columnCount || left.some((row) => row.length !== columnCount))
    throw new FormulaValueError("#VALUE!");
  assertFormulaCellCount(left.length, right[0]!.length);
  return left.map((row) =>
    right[0]!.map((_, column) =>
      row.reduce((sum, value, index) => sum + value * right[index]![column]!, 0),
    ),
  );
}

function romanToNumber(value: string): number {
  const values: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  return value
    .toUpperCase()
    .split("")
    .reduce(
      (output, character, index, characters) =>
        values[character]! < (values[characters[index + 1]!] ?? 0)
          ? output - values[character]!
          : output + values[character]!,
      0,
    );
}

function romanForm(value: number, form: number): string {
  if (value < 0 || value > 3999 || !Number.isInteger(value)) throw new FormulaValueError("#VALUE!");
  const numerals: [number, string][] =
    form === 0
      ? [
          [1000, "M"],
          [900, "CM"],
          [500, "D"],
          [400, "CD"],
          [100, "C"],
          [90, "XC"],
          [50, "L"],
          [40, "XL"],
          [10, "X"],
          [9, "IX"],
          [5, "V"],
          [4, "IV"],
          [1, "I"],
        ]
      : form === 1
        ? [
            [1000, "M"],
            [900, "CM"],
            [450, "LD"],
            [400, "CD"],
            [90, "XC"],
            [45, "VL"],
            [40, "XL"],
            [9, "IX"],
            [5, "V"],
            [4, "IV"],
            [1, "I"],
          ]
        : form === 2
          ? [
              [1000, "M"],
              [900, "CM"],
              [490, "XD"],
              [400, "CD"],
              [90, "XC"],
              [49, "IL"],
              [40, "XL"],
              [9, "IX"],
              [5, "V"],
              [4, "IV"],
              [1, "I"],
            ]
          : form === 3
            ? [
                [1000, "M"],
                [900, "CM"],
                [495, "VD"],
                [400, "CD"],
                [90, "XC"],
                [49, "IL"],
                [40, "XL"],
                [9, "IX"],
                [5, "V"],
                [4, "IV"],
                [1, "I"],
              ]
            : [
                [1000, "M"],
                [999, "IM"],
                [499, "ID"],
                [99, "IC"],
                [9, "IX"],
                [5, "V"],
                [4, "IV"],
                [1, "I"],
              ];
  let remaining = value;
  return numerals
    .map(([amount, numeral]) => {
      const repeat = Math.floor(remaining / amount);
      remaining %= amount;
      return numeral.repeat(repeat);
    })
    .join("");
}
