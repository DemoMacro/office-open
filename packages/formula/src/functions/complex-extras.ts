import { FormulaValueError } from "../errors";
import type { ComplexNumber } from "./engineering";
import { formatComplex, parseComplex } from "./engineering";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";

function divide(left: ComplexNumber, right: ComplexNumber): ComplexNumber {
  const denominator = right.real ** 2 + right.imaginary ** 2;
  if (denominator === 0) throw new FormulaValueError("#NUM!");
  return {
    real: (left.real * right.real + left.imaginary * right.imaginary) / denominator,
    imaginary: (left.imaginary * right.real - left.real * right.imaginary) / denominator,
  };
}

function hyperbolicCosine(value: ComplexNumber): ComplexNumber {
  return {
    real: Math.cosh(value.real) * Math.cos(value.imaginary),
    imaginary: Math.sinh(value.real) * Math.sin(value.imaginary),
  };
}

function hyperbolicSine(value: ComplexNumber): ComplexNumber {
  return {
    real: Math.sinh(value.real) * Math.cos(value.imaginary),
    imaginary: Math.cosh(value.real) * Math.sin(value.imaginary),
  };
}

function cosine(value: ComplexNumber): ComplexNumber {
  return {
    real: Math.cos(value.real) * Math.cosh(value.imaginary),
    imaginary: -Math.sin(value.real) * Math.sinh(value.imaginary),
  };
}

function sine(value: ComplexNumber): ComplexNumber {
  return {
    real: Math.sin(value.real) * Math.cosh(value.imaginary),
    imaginary: Math.cos(value.real) * Math.sinh(value.imaginary),
  };
}

function formatHyperbolic(
  functionName: "cosh" | "cot" | "csc" | "sec" | "sech" | "csch" | "sinh",
): FormulaFunctionImplementation {
  return ([value]) => {
    const complex = parseComplex(value!);
    const cosh = hyperbolicCosine(complex);
    const sinh = hyperbolicSine(complex);
    switch (functionName) {
      case "cosh":
        return formatComplex(cosh);
      case "sinh":
        return formatComplex(sinh);
      case "cot":
        return formatComplex(divide(cosh, sinh));
      case "sech":
        return formatComplex(divide({ real: 1, imaginary: 0 }, cosh));
      case "csc":
        return formatComplex(divide({ real: 1, imaginary: 0 }, sine(complex)));
      case "sec":
        return formatComplex(divide({ real: 1, imaginary: 0 }, cosine(complex)));
      default:
        return formatComplex(divide({ real: 1, imaginary: 0 }, sinh));
    }
  };
}

export const complexExtraFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  ["IMCOSH", formatHyperbolic("cosh")],
  ["IMCOT", formatHyperbolic("cot")],
  ["IMCSC", formatHyperbolic("csc")],
  ["IMCSCH", formatHyperbolic("csch")],
  ["IMSEC", formatHyperbolic("sec")],
  ["IMSECH", formatHyperbolic("sech")],
  ["IMSINH", formatHyperbolic("sinh")],
]);
