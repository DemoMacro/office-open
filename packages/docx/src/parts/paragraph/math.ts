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

// ── Math structure properties ──

/** Delimiter properties (CT_DPr) — bracket characters, growth, shape. */
export interface MathDelimiterProperties {
  /** Beginning character (m:begChr). */
  beginCharacter?: string;
  /** Ending character (m:endChr). */
  endCharacter?: string;
  /** Separator character between elements (m:sepChr). */
  separatorCharacter?: string;
  /** Whether delimiters grow to content height (m:grow). */
  grow?: boolean;
  /** Delimiter shape — centered or match (m:shp, ST_Shp). */
  shape?: "centered" | "match";
}

/** N-ary operator limit location (ST_LimLoc). */
export type MathNaryLimitLocation = "subSup" | "undOvr";

/** N-ary operator properties (CT_NaryPr beyond chr/subHide/supHide). */
export interface MathNaryProperties {
  /** Limit location relative to the operator (m:limLoc). */
  limitLocation?: MathNaryLimitLocation;
  /** Whether the operator grows to content height (m:grow). */
  grow?: boolean;
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
        /** Argument size scaling for the numerator (m:num/m:argPr/m:argSz). */
        numeratorArgumentSize?: number;
        /** Argument size scaling for the denominator (m:den/m:argPr/m:argSz). */
        denominatorArgumentSize?: number;
      };
    }
  | { superScript: { children: MathInput[]; superScript: MathInput[] } }
  | { subScript: { children: MathInput[]; subScript: MathInput[] } }
  | {
      subSuperScript: {
        children: MathInput[];
        subScript: MathInput[];
        superScript: MathInput[];
        /** Align sub/super scripts (m:sSubSupPr/m:alnScr). */
        alignScript?: boolean;
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
        properties?: MathNaryProperties;
      };
    }
  | {
      integral: {
        children: MathInput[];
        subScript?: MathInput[];
        superScript?: MathInput[];
        properties?: MathNaryProperties;
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
  | {
      roundBrackets: MathInput[] | { children: MathInput[]; properties?: MathDelimiterProperties };
    }
  | {
      curlyBrackets: MathInput[] | { children: MathInput[]; properties?: MathDelimiterProperties };
    }
  | {
      angledBrackets: MathInput[] | { children: MathInput[]; properties?: MathDelimiterProperties };
    }
  | {
      squareBrackets: MathInput[] | { children: MathInput[]; properties?: MathDelimiterProperties };
    }
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
