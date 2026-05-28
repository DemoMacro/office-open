/**
 * JSON API types and coercion for math components.
 *
 * Provides a recursive JSON representation of math components and a coercion
 * function that converts JSON objects to class instances.
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import { createMathAccent } from "./accent/math-accent";
import { createMathBar } from "./bar/math-bar";
import { MathBorderBox } from "./border-box/math-border-box";
import type { MathBorderBoxPropertiesOptions } from "./border-box/math-border-box-properties";
import { MathBox } from "./box/math-box";
import type { MathBoxPropertiesOptions } from "./box/math-box-properties";
import {
  MathAngledBrackets,
  MathCurlyBrackets,
  MathRoundBrackets,
  MathSquareBrackets,
} from "./brackets";
import { MathEqArr } from "./eq-arr/math-eq-arr";
import type { MathEqArrPropertiesOptions } from "./eq-arr/math-eq-arr-properties";
import { MathFraction } from "./fraction/math-fraction";
import type { FractionType } from "./fraction/math-fraction-properties";
import { MathFunction } from "./function/math-function";
import type { MathFunctionPropertiesOptions } from "./function/math-function-properties";
import { MathGroupChr } from "./group-chr/math-group-chr";
import type { MathGroupChrPropertiesOptions } from "./group-chr/math-group-chr-properties";
import type { MathComponent } from "./math-component";
import { MathRun } from "./math-run";
import type { MathRunPropertiesOptions } from "./math-run-properties";
import { MathMatrix } from "./matrix/math-matrix";
import type { MathMatrixPropertiesOptions } from "./matrix/math-matrix-properties";
import { MathIntegral, MathLimitLower, MathLimitUpper, MathSum } from "./n-ary";
import type { MathLimitLowPropertiesOptions } from "./n-ary/math-limit-low-properties";
import type { MathLimitUpperPropertiesOptions } from "./n-ary/math-limit-upper-properties";
import { MathPhant } from "./phant/math-phant";
import type { MathPhantPropertiesOptions } from "./phant/math-phant-properties";
import { MathRadical } from "./radical/math-radical";
import { MathSubScript, MathSubSuperScript, MathSuperScript } from "./script";

/**
 * Recursive JSON representation of math components.
 *
 * Each discriminated union member uses a unique key name to identify the component type.
 * Supports mixing JSON objects with class instances.
 *
 * @example
 * ```typescript
 * const children: MathJson[] = [
 *   "x",                                                          // MathRun shorthand
 *   { fraction: { numerator: ["a"], denominator: ["b"] } },       // Fraction
 *   { superScript: { children: ["x"], superScript: ["2"] } },     // x²
 * ];
 * ```
 */
export type MathJson =
  | MathComponent
  | string
  | { readonly text: string; readonly properties?: MathRunPropertiesOptions }
  | {
      readonly fraction: {
        readonly numerator: readonly MathJson[];
        readonly denominator: readonly MathJson[];
        readonly fractionType?: (typeof FractionType)[keyof typeof FractionType];
      };
    }
  | {
      readonly superScript: {
        readonly children: readonly MathJson[];
        readonly superScript: readonly MathJson[];
      };
    }
  | {
      readonly subScript: {
        readonly children: readonly MathJson[];
        readonly subScript: readonly MathJson[];
      };
    }
  | {
      readonly subSuperScript: {
        readonly children: readonly MathJson[];
        readonly subScript: readonly MathJson[];
        readonly superScript: readonly MathJson[];
      };
    }
  | {
      readonly radical: {
        readonly children: readonly MathJson[];
        readonly degree?: readonly MathJson[];
      };
    }
  | {
      readonly sum: {
        readonly children: readonly MathJson[];
        readonly subScript?: readonly MathJson[];
        readonly superScript?: readonly MathJson[];
      };
    }
  | {
      readonly integral: {
        readonly children: readonly MathJson[];
        readonly subScript?: readonly MathJson[];
        readonly superScript?: readonly MathJson[];
      };
    }
  | {
      readonly limitLower: {
        readonly children: readonly MathJson[];
        readonly limit: readonly MathJson[];
        readonly properties?: MathLimitLowPropertiesOptions;
      };
    }
  | {
      readonly limitUpper: {
        readonly children: readonly MathJson[];
        readonly limit: readonly MathJson[];
        readonly properties?: MathLimitUpperPropertiesOptions;
      };
    }
  | {
      readonly function: {
        readonly children: readonly MathJson[];
        readonly name: readonly MathJson[];
        readonly properties?: MathFunctionPropertiesOptions;
      };
    }
  | {
      readonly matrix: {
        readonly rows: readonly (readonly MathJson[])[];
        readonly properties?: MathMatrixPropertiesOptions;
      };
    }
  | { readonly roundBrackets: readonly MathJson[] | { readonly children: readonly MathJson[] } }
  | { readonly curlyBrackets: readonly MathJson[] | { readonly children: readonly MathJson[] } }
  | { readonly angledBrackets: readonly MathJson[] | { readonly children: readonly MathJson[] } }
  | { readonly squareBrackets: readonly MathJson[] | { readonly children: readonly MathJson[] } }
  | {
      readonly borderBox: {
        readonly children: readonly MathJson[];
        readonly properties?: MathBorderBoxPropertiesOptions;
      };
    }
  | {
      readonly box: {
        readonly children: readonly MathJson[];
        readonly properties?: MathBoxPropertiesOptions;
      };
    }
  | {
      readonly groupChr: {
        readonly children: readonly MathJson[];
        readonly properties?: MathGroupChrPropertiesOptions;
      };
    }
  | {
      readonly phant: {
        readonly children: readonly MathJson[];
        readonly properties?: MathPhantPropertiesOptions;
      };
    }
  | {
      readonly eqArr: {
        readonly rows: readonly (readonly MathJson[])[];
        readonly properties?: MathEqArrPropertiesOptions;
      };
    }
  | {
      readonly accent: {
        readonly children: readonly MathJson[];
        readonly accentCharacter?: string;
      };
    }
  | {
      readonly bar: {
        readonly children: readonly MathJson[];
        readonly type: "top" | "bot";
      };
    };

/** Coerce an array of MathJson values to MathComponent instances. */
function coerceArray(items: readonly MathJson[] | undefined): readonly MathComponent[] {
  if (!items) return [];
  return items.map((v) => coerceMathJson(v)) as readonly MathComponent[];
}

/** Coerce bracket shorthand to children array. */
function bracketChildren(
  v: readonly MathJson[] | { readonly children: readonly MathJson[] },
): readonly MathComponent[] {
  if (Array.isArray(v)) return coerceArray(v);
  return coerceArray((v as { readonly children: readonly MathJson[] }).children);
}

/**
 * Coerce a MathJson value to an XmlComponent instance.
 *
 * Handles string → MathRun, class instances pass-through, and recursive
 * conversion of all discriminated JSON object types.
 */
export function coerceMathJson(value: MathJson): XmlComponent {
  // String → MathRun
  if (typeof value === "string") return new MathRun(value);

  // MathComponent instances pass through
  if (value instanceof XmlComponent) return value;

  // Discriminated union: check unique keys (order matters)
  // subSuperScript before superScript/subScript to avoid false match
  if ("subSuperScript" in value) {
    const opts = value.subSuperScript;
    return new MathSubSuperScript({
      children: coerceArray(opts.children),
      subScript: coerceArray(opts.subScript),
      superScript: coerceArray(opts.superScript),
    });
  }

  if ("superScript" in value) {
    const opts = value.superScript;
    return new MathSuperScript({
      children: coerceArray(opts.children),
      superScript: coerceArray(opts.superScript),
    });
  }

  if ("subScript" in value) {
    const opts = value.subScript;
    return new MathSubScript({
      children: coerceArray(opts.children),
      subScript: coerceArray(opts.subScript),
    });
  }

  if ("fraction" in value) {
    const opts = value.fraction;
    return new MathFraction({
      numerator: coerceArray(opts.numerator),
      denominator: coerceArray(opts.denominator),
      ...(opts.fractionType ? { fractionType: opts.fractionType } : {}),
    });
  }

  if ("radical" in value) {
    const opts = value.radical;
    return new MathRadical({
      children: coerceArray(opts.children),
      ...(opts.degree ? { degree: coerceArray(opts.degree) } : {}),
    });
  }

  if ("sum" in value) {
    const opts = value.sum;
    return new MathSum({
      children: coerceArray(opts.children),
      ...(opts.subScript ? { subScript: coerceArray(opts.subScript) } : {}),
      ...(opts.superScript ? { superScript: coerceArray(opts.superScript) } : {}),
    });
  }

  if ("integral" in value) {
    const opts = value.integral;
    return new MathIntegral({
      children: coerceArray(opts.children),
      ...(opts.subScript ? { subScript: coerceArray(opts.subScript) } : {}),
      ...(opts.superScript ? { superScript: coerceArray(opts.superScript) } : {}),
    });
  }

  if ("limitLower" in value) {
    const opts = value.limitLower;
    return new MathLimitLower({
      children: coerceArray(opts.children),
      limit: coerceArray(opts.limit),
      ...(opts.properties ? { properties: opts.properties } : {}),
    });
  }

  if ("limitUpper" in value) {
    const opts = value.limitUpper;
    return new MathLimitUpper({
      children: coerceArray(opts.children),
      limit: coerceArray(opts.limit),
      ...(opts.properties ? { properties: opts.properties } : {}),
    });
  }

  if ("function" in value) {
    const opts = value.function;
    return new MathFunction({
      children: coerceArray(opts.children),
      name: coerceArray(opts.name),
      ...(opts.properties ? { properties: opts.properties } : {}),
    });
  }

  if ("matrix" in value) {
    const opts = value.matrix;
    return new MathMatrix({
      rows: opts.rows.map((row) => coerceArray(row)) as readonly (readonly MathComponent[])[],
      ...(opts.properties ? { properties: opts.properties } : {}),
    });
  }

  // Bracket types: support shorthand (array) or full ({ children })
  if ("roundBrackets" in value) {
    return new MathRoundBrackets({ children: bracketChildren(value.roundBrackets) });
  }

  if ("curlyBrackets" in value) {
    return new MathCurlyBrackets({ children: bracketChildren(value.curlyBrackets) });
  }

  if ("angledBrackets" in value) {
    return new MathAngledBrackets({ children: bracketChildren(value.angledBrackets) });
  }

  if ("squareBrackets" in value) {
    return new MathSquareBrackets({ children: bracketChildren(value.squareBrackets) });
  }

  if ("borderBox" in value) {
    const opts = value.borderBox;
    return new MathBorderBox({
      children: coerceArray(opts.children),
      ...(opts.properties ? { properties: opts.properties } : {}),
    });
  }

  if ("box" in value) {
    const opts = value.box;
    return new MathBox({
      children: coerceArray(opts.children),
      ...(opts.properties ? { properties: opts.properties } : {}),
    });
  }

  if ("groupChr" in value) {
    const opts = value.groupChr;
    return new MathGroupChr({
      children: coerceArray(opts.children),
      ...(opts.properties ? { properties: opts.properties } : {}),
    });
  }

  if ("phant" in value) {
    const opts = value.phant;
    return new MathPhant({
      children: coerceArray(opts.children),
      ...(opts.properties ? { properties: opts.properties } : {}),
    });
  }

  if ("eqArr" in value) {
    const opts = value.eqArr;
    return new MathEqArr({
      rows: opts.rows.map((row) => coerceArray(row)) as readonly (readonly MathComponent[])[],
      ...(opts.properties ? { properties: opts.properties } : {}),
    });
  }

  if ("accent" in value) {
    const opts = value.accent;
    return createMathAccent({
      children: coerceArray(opts.children) as readonly MathComponent[],
      ...(opts.accentCharacter ? { accentCharacter: opts.accentCharacter } : {}),
    });
  }

  if ("bar" in value) {
    const opts = value.bar;
    return createMathBar({
      children: coerceArray(opts.children) as readonly MathComponent[],
      type: opts.type,
    });
  }

  // Fallback: { text: string; properties?: ... } → MathRun
  return new MathRun({ text: value.text, properties: value.properties });
}
