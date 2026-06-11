/**
 * Math types for Office MathML (OMML).
 *
 * Provides the canonical input type for math content and run properties.
 * No XmlComponent dependencies — pure type definitions.
 *
 * @module
 */

// ── MathRunProperties ──

export type MathScriptType =
  | "roman"
  | "script"
  | "fraktur"
  | "double-struck"
  | "sans-serif"
  | "monospace";

export type MathStyleType = "p" | "b" | "i" | "bi";

export interface MathRunPropertiesOptions {
  lit?: boolean;
  normal?: boolean;
  script?: MathScriptType;
  style?: MathStyleType;
  breakAlignment?: number;
  align?: boolean;
}

// ── MathInput ──

/**
 * Recursive input type for math content.
 *
 * Each discriminated union member uses a unique key to identify the component type.
 * Used by `parts/paragraph/math/stringify.ts` for direct XML string generation.
 */
export type MathInput =
  | string
  | { text: string; properties?: MathRunPropertiesOptions }
  | {
      fraction: {
        numerator: MathInput[];
        denominator: MathInput[];
        fractionType?: string;
      };
    }
  | { superScript: { children: MathInput[]; superScript: MathInput[] } }
  | { subScript: { children: MathInput[]; subScript: MathInput[] } }
  | {
      subSuperScript: {
        children: MathInput[];
        subScript: MathInput[];
        superScript: MathInput[];
      };
    }
  | {
      preSubSuperScript: {
        children: MathInput[];
        subScript: MathInput[];
        superScript: MathInput[];
      };
    }
  | { radical: { children: MathInput[]; degree?: MathInput[] } }
  | {
      sum: {
        children: MathInput[];
        subScript?: MathInput[];
        superScript?: MathInput[];
      };
    }
  | {
      integral: {
        children: MathInput[];
        subScript?: MathInput[];
        superScript?: MathInput[];
      };
    }
  | {
      limitLower: {
        children: MathInput[];
        limit: MathInput[];
        properties?: Record<string, unknown>;
      };
    }
  | {
      limitUpper: {
        children: MathInput[];
        limit: MathInput[];
        properties?: Record<string, unknown>;
      };
    }
  | {
      function: {
        children: MathInput[];
        name: MathInput[];
        properties?: Record<string, unknown>;
      };
    }
  | {
      matrix: {
        rows: MathInput[][];
        properties?: Record<string, unknown>;
      };
    }
  | { roundBrackets: MathInput[] | { children: MathInput[] } }
  | { curlyBrackets: MathInput[] | { children: MathInput[] } }
  | { angledBrackets: MathInput[] | { children: MathInput[] } }
  | { squareBrackets: MathInput[] | { children: MathInput[] } }
  | {
      borderBox: {
        children: MathInput[];
        properties?: Record<string, unknown>;
      };
    }
  | {
      box: {
        children: MathInput[];
        properties?: Record<string, unknown>;
      };
    }
  | {
      groupChr: {
        children: MathInput[];
        properties?: Record<string, unknown>;
      };
    }
  | {
      phant: {
        children: MathInput[];
        properties?: Record<string, unknown>;
      };
    }
  | {
      eqArr: {
        rows: MathInput[][];
        properties?: Record<string, unknown>;
      };
    }
  | { accent: { children: MathInput[]; accentCharacter?: string } }
  | { bar: { children: MathInput[]; type: "top" | "bot" } };
