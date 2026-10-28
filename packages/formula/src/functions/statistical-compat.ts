import type { FormulaValue } from "../value";
import { coreExtraFunctions } from "./core-extras";
import { statisticalFunctions } from "./statistical";
import { statisticalExtraFunctions } from "./statistical-extras";
import { studentTailProbability } from "./statistical-numeric";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { scalarNumber } from "./utils";

function resolve(name: string): FormulaFunctionImplementation {
  const implementation =
    statisticalFunctions.get(name) ??
    statisticalExtraFunctions.get(name) ??
    coreExtraFunctions.get(name);
  if (!implementation) throw new Error(`Statistical function is unavailable: ${name}`);
  return implementation;
}

function alias(name: string): FormulaFunctionImplementation {
  const implementation = resolve(name);
  return (args, context) => implementation(args, context);
}

function legacyTDist([value, degrees, tails]: Parameters<FormulaFunctionImplementation>[0]):
  | FormulaValue
  | string {
  return studentTailProbability(scalarNumber(value), scalarNumber(degrees), scalarNumber(tails!));
}

export const statisticalCompatFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  ["ERRORTYPE", alias("ERROR.TYPE")],
  ["MODECOMPAT", alias("MODE")],
  ["MODEMULT", alias("MODE.MULT")],
  ["MODESINGL", alias("MODE.SNGL")],
  ["PERCENTILEEXC", alias("PERCENTILE.EXC")],
  ["PERCENTILEINC", alias("PERCENTILE.INC")],
  ["PERCENTRANKEXC", alias("PERCENTRANK.EXC")],
  ["PERCENTRANKINC", alias("PERCENTRANK.INC")],
  ["POISSONDIST", alias("POISSON.DIST")],
  ["QUARTILEEXC", alias("QUARTILE.EXC")],
  ["QUARTILEINC", alias("QUARTILE.INC")],
  ["RANKAVG", alias("RANK.AVG")],
  ["RANKEQ", alias("RANK.EQ")],
  ["TDIST2T", alias("T.DIST.2T")],
  ["TDISTCOMPAT", alias("TDIST")],
  ["TDISTRT", alias("T.DIST.RT")],
  ["LEGACY.CHIDIST", alias("LEGACYCHIDIST")],
  ["LEGACY.CHIINV", alias("LEGACYCHIINV")],
  ["LEGACY.CHITEST", alias("LEGACYCHITEST")],
  ["LEGACY.FDIST", alias("LEGACYFDIST")],
  ["LEGACY.FINV", alias("LEGACYFINV")],
  ["LEGACY.NORMSDIST", alias("LEGACYNORMSDIST")],
  ["LEGACY.NORMSINV", alias("LEGACYNORMSINV")],
  ["LEGACY.TDIST", legacyTDist],
  ["TINV2T", alias("T.INV.2T")],
  ["TINVCOMPAT", alias("TINV")],
  ["TTESTCOMPAT", alias("T.TEST")],
  ["LOGNORMINV", alias("LOGNORM.INV")],
  ["VARS", alias("VAR")],
  ["WEIBULLCOMPAT", alias("WEIBULL")],
  ["WEIBULLDIST", alias("WEIBULL.DIST")],
  ["ZTESTCOMPAT", alias("Z.TEST")],
]);
