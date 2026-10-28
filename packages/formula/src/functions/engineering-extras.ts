import { FormulaValueError } from "../errors";
import { engineeringFunctions } from "./engineering";
import { mathFunctions } from "./math";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { scalarNumber } from "./utils";

type RoundingDirection = "toward-negative" | "toward-positive";

const bitLimit = 2 ** 48 - 1;

function requireBitNumber(value: number): number {
  if (!Number.isInteger(value) || value < 0 || value > bitLimit)
    throw new FormulaValueError("#NUM!");
  return value;
}

function requireSignificance(value: number): number {
  if (!Number.isFinite(value) || value === 0) throw new FormulaValueError("#NUM!");
  return value;
}

function excelRound(value: number, significance: number, direction: RoundingDirection): number {
  if (value === 0) return 0;
  const quotient = value / significance;
  const multiple =
    direction === "toward-positive"
      ? value > 0
        ? Math.ceil(quotient)
        : Math.trunc(quotient)
      : value > 0
        ? Math.trunc(quotient)
        : Math.floor(quotient);
  return multiple * significance;
}

function binaryOperation(
  operator: (left: bigint, right: bigint) => bigint,
): FormulaFunctionImplementation {
  return ([left, right]) =>
    requireBitNumber(
      Number(
        operator(
          BigInt(requireBitNumber(scalarNumber(left))),
          BigInt(requireBitNumber(scalarNumber(right))),
        ),
      ),
    );
}

function shiftFunction(direction: 1 | -1): FormulaFunctionImplementation {
  return ([value, amount]) => {
    const number = requireBitNumber(scalarNumber(value));
    const shift = requireBitNumber(scalarNumber(amount));
    if (shift > 48) throw new FormulaValueError("#NUM!");
    const shifted =
      shift === 0
        ? BigInt(number)
        : direction === 1
          ? BigInt(number) << BigInt(shift)
          : BigInt(number) >> BigInt(shift);
    if (shifted > bitLimit) throw new FormulaValueError("#NUM!");
    return requireBitNumber(Number(shifted));
  };
}

function delegate(table: FormulaFunctionTable, name: string): FormulaFunctionImplementation {
  const implementation = table.get(name);
  if (!implementation) throw new Error(`missing formula function: ${name}`);
  return (args, context) => implementation(args, context);
}

function aliasFunction(table: FormulaFunctionTable, name: string): FormulaFunctionImplementation {
  const implementation = delegate(table, name);
  return (args, context) => implementation(args, context);
}

export const engineeringExtraFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  ["BITAND", binaryOperation((left, right) => left & right)],
  ["BITLSHIFT", shiftFunction(1)],
  ["BITOR", binaryOperation((left, right) => left | right)],
  ["BITRSHIFT", shiftFunction(-1)],
  ["BITXOR", binaryOperation((left, right) => left ^ right)],
  [
    "CEILING.XCL",
    ([value, significance]) =>
      excelRound(
        scalarNumber(value),
        requireSignificance(scalarNumber(significance)),
        "toward-positive",
      ),
  ],
  ["CEILINGMATH", aliasFunction(mathFunctions, "CEILING.MATH")],
  ["CEILINGPRECISE", aliasFunction(mathFunctions, "CEILING.PRECISE")],
  ["ECMA.CEILING", aliasFunction(mathFunctions, "CEILING.PRECISE")],
  ["ECMACEILING", aliasFunction(mathFunctions, "CEILING.PRECISE")],
  ["ERFCPRECISE", aliasFunction(engineeringFunctions, "ERFC.PRECISE")],
  ["ERFPRECISE", aliasFunction(engineeringFunctions, "ERF.PRECISE")],
  [
    "FLOOR.XCL",
    ([value, significance]) =>
      excelRound(
        scalarNumber(value),
        requireSignificance(scalarNumber(significance)),
        "toward-negative",
      ),
  ],
  ["FLOORMATH", aliasFunction(mathFunctions, "FLOOR.MATH")],
  ["FLOORPRECISE", aliasFunction(mathFunctions, "FLOOR.PRECISE")],
  ["ISO.CEILING", aliasFunction(mathFunctions, "CEILING.PRECISE")],
  ["ISOCEILING", aliasFunction(mathFunctions, "CEILING.PRECISE")],
  ["CONVERTOOO", aliasFunction(engineeringFunctions, "CONVERT")],
  ["CONVERT_OOO", aliasFunction(engineeringFunctions, "CONVERT")],
  ["GCD_EXCEL2003", aliasFunction(mathFunctions, "GCD")],
  ["LCM_EXCEL2003", aliasFunction(mathFunctions, "LCM")],
]);
