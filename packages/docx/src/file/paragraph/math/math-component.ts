/**
 * Math component types module.
 *
 * This module defines the union type of all valid math components
 * that can appear within a Math element.
 *
 * @module
 */
import type { MathBorderBox } from "./border-box";
import type { MathBox } from "./box";
import type {
    MathAngledBrackets,
    MathCurlyBrackets,
    MathRoundBrackets,
    MathSquareBrackets,
} from "./brackets";
import type { MathEqArr } from "./eq-arr";
import type { MathFraction } from "./fraction";
import type { MathFunction } from "./function";
import type { MathGroupChr } from "./group-chr";
import type { MathRun } from "./math-run";
import type { MathMatrix } from "./matrix";
import type { MathIntegral, MathLimitLower, MathLimitUpper, MathSum } from "./n-ary";
import type { MathPhant } from "./phant";
import type { MathRadical } from "./radical";
import type { MathSubScript, MathSubSuperScript, MathSuperScript } from "./script";

/**
 * Union type of all valid math components.
 *
 * MathComponent represents any element that can appear within a Math equation,
 * including runs, fractions, radicals, integrals, sums, scripts, brackets,
 * boxes, matrices, and more.
 */
export type MathComponent =
    | MathRun
    | MathFraction
    | MathSum
    | MathIntegral
    | MathSuperScript
    | MathSubScript
    | MathSubSuperScript
    | MathRadical
    | MathFunction
    | MathRoundBrackets
    | MathCurlyBrackets
    | MathAngledBrackets
    | MathSquareBrackets
    | MathBox
    | MathBorderBox
    | MathEqArr
    | MathGroupChr
    | MathLimitLower
    | MathLimitUpper
    | MathMatrix
    | MathPhant;

// Needed because of: https://github.com/s-panferov/awesome-typescript-loader/issues/432
/**
 * @ignore
 */
export const WORKAROUND4 = "";
