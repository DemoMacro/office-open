import { FormulaValueError } from "../errors";
import type { FormulaMatrix, FormulaScalar, FormulaValue } from "../value";
import { isMatrix } from "../value";
import { formatComplex, parseComplex } from "./engineering";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { scalarNumber } from "./utils";

interface ComplexSample {
  real: number;
  imaginary: number;
}

function toComplex(value: FormulaScalar): ComplexSample {
  if (typeof value === "number") return { real: value, imaginary: 0 };
  if (typeof value === "string" && value !== "") return parseComplex(value);
  return { real: 0, imaginary: 0 };
}

function matrix(value: FormulaValue): FormulaMatrix {
  return isMatrix(value) ? value : [[value]];
}

function samples(input: FormulaValue): { rows: ComplexSample[][]; flat: ComplexSample[] } {
  const rows = matrix(input).map((row) => row.map(toComplex));
  return { rows, flat: rows.flat() };
}

function sameShape(input: ComplexSample[][], flat: ComplexSample[]): FormulaMatrix {
  let index = 0;
  return input.map((row) =>
    row.map(() => {
      const sample = flat[index++] ?? { real: 0, imaginary: 0 };
      return formatComplex(sample);
    }),
  );
}

function multiply(left: ComplexSample, right: ComplexSample): ComplexSample {
  return {
    real: left.real * right.real - left.imaginary * right.imaginary,
    imaginary: left.real * right.imaginary + left.imaginary * right.real,
  };
}

function applyWindow(input: ComplexSample[], window: FormulaScalar): ComplexSample[] {
  const mode = typeof window === "string" ? window.toUpperCase() : "RECTANGULAR";
  return input.map((sample, index) => {
    if (mode === "RECTANGULAR" || mode === "") return sample;
    const position = input.length === 1 ? 0 : index / (input.length - 1);
    const factor =
      mode === "HAMMING"
        ? 0.54 - 0.46 * Math.cos(2 * Math.PI * position)
        : mode === "HANNING"
          ? 0.5 - 0.5 * Math.cos(2 * Math.PI * position)
          : NaN;
    if (!Number.isFinite(factor)) throw new FormulaValueError("#VALUE!");
    return { real: sample.real * factor, imaginary: sample.imaginary * factor };
  });
}

function transform(values: ComplexSample[], direction: number): ComplexSample[] {
  const output = [...values];
  const length = output.length;
  for (let index = 1, reversed = 0; index < length; index += 1) {
    let high = length >> 1;
    for (; reversed & high; high >>= 1) reversed ^= high;
    reversed ^= high;
    if (index < reversed) [output[index], output[reversed]] = [output[reversed]!, output[index]!];
  }
  for (let span = 2; span <= length; span <<= 1) {
    const angle = (direction * 2 * Math.PI) / span;
    const step = { real: Math.cos(angle), imaginary: Math.sin(angle) };
    for (let start = 0; start < length; start += span) {
      let rotation = { real: 1, imaginary: 0 };
      for (let offset = start; offset < start + span / 2; offset += 1) {
        const left = output[offset]!;
        const right = multiply(output[offset + span / 2]!, rotation);
        output[offset] = {
          real: left.real + right.real,
          imaginary: left.imaginary + right.imaginary,
        };
        output[offset + span / 2] = {
          real: left.real - right.real,
          imaginary: left.imaginary - right.imaginary,
        };
        rotation = multiply(rotation, step);
      }
    }
  }
  return output;
}

function fourier(
  input: FormulaValue,
  size: FormulaValue,
  inverse: FormulaValue,
  window?: FormulaValue,
): FormulaValue {
  const source = samples(input);
  const requested = size === undefined ? source.flat.length : scalarNumber(size);
  const length = Math.trunc(requested);
  if (length <= 0 || (length & (length - 1)) !== 0) return "#NUM!";
  const normalized: ComplexSample[] = Array.from(
    { length },
    (_, index) =>
      source.flat[index] ?? {
        real: 0,
        imaginary: 0,
      },
  );
  const windowValue =
    window === undefined ? undefined : Array.isArray(window) ? (window.flat()[0] ?? null) : window;
  const values = applyWindow(normalized, windowValue ?? "RECTANGULAR");
  const backward = inverse !== undefined && scalarNumber(inverse) !== 0;
  const direction = backward ? 1 : -1;
  const scale = backward ? 1 / length : 1;
  const transformed = transform(values, direction).map((sample) => ({
    real: sample.real * scale,
    imaginary: sample.imaginary * scale,
  }));
  const output = transformed;
  return sameShape(source.rows, output);
}

export const fourierFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  [
    "FOURIER",
    ([input, size, inverse, _unit, window]) =>
      fourier(input!, size, inverse, window ?? "RECTANGULAR"),
  ],
]);
