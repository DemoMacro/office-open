/**
 * Math types for Office MathML (OMML).
 *
 * Provides the canonical input type for math content and run properties.
 * No XmlComponent dependencies — pure type definitions.
 *
 * @module
 */

import type { RunPropertiesOptions } from "./run/properties";

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
  /** Control-character formatting (m:dPr/m:ctrlPr → w:rPr). */
  ctrlPr?: RunPropertiesOptions;
}

/** N-ary operator limit location (ST_LimLoc). */
export type MathNaryLimitLocation = "subSup" | "undOvr";

/** Matrix/eqArr row-spacing rule (ST_SpacingRule, integer 0-4). */
export type MathSpacingRule = 0 | 1 | 2 | 3 | 4;

/** Matrix baseline alignment (s:ST_YAlign). */
export type MathBaselineAlignment = "inline" | "top" | "center" | "bottom" | "inside" | "outside";

/** N-ary operator properties (CT_NaryPr beyond chr/subHide/supHide). */
export interface MathNaryProperties {
  /** Limit location relative to the operator (m:limLoc). */
  limitLocation?: MathNaryLimitLocation;
  /** Whether the operator grows to content height (m:grow). */
  grow?: boolean;
}

/** Matrix column justification (one m:mc entry inside m:mcs). */
export interface MathMatrixColumnOptions {
  /** Number of columns the justification applies to (m:mcPr/m:count). */
  count: number;
  /** Column justification (m:mcPr/m:mcJc). */
  justification: string;
}

/** Matrix properties (CT_MPr). */
export interface MathMatrixProperties {
  /** Row baseline alignment (m:baseJc). */
  baseJc?: MathBaselineAlignment;
  /** Hide placeholder glyphs in empty cells (m:plcHide). */
  plcHide?: boolean;
  /** Row spacing rule (m:rSpRule). */
  rSpRule?: MathSpacingRule;
  /** Column spacing rule (m:cGpRule). */
  cGpRule?: MathSpacingRule;
  /** Row spacing (m:rSp). */
  rSp?: number;
  /** Column spacing (m:cSp). */
  cSp?: number;
  /** Column gap (m:cGp). */
  cGp?: number;
  /** Matrix column justifications (m:mcs). */
  mcs?: MathMatrixColumnOptions[];
}

/** Border-box properties (CT_BorderBoxPr) — which edges/strikes to hide or draw. */
export interface MathBorderBoxProperties {
  /** Omit the top border (m:hideTop). */
  hideTop?: boolean;
  /** Omit the bottom border (m:hideBot). */
  hideBottom?: boolean;
  /** Omit the left border (m:hideLeft). */
  hideLeft?: boolean;
  /** Omit the right border (m:hideRight). */
  hideRight?: boolean;
  /** Draw the horizontal strike (m:strikeH). */
  strikeHorizontal?: boolean;
  /** Draw the vertical strike (m:strikeV). */
  strikeVertical?: boolean;
  /** Draw the bottom-left→top-right diagonal (m:strikeBLTR). */
  strikeDiagonalUp?: boolean;
  /** Draw the top-left→bottom-right diagonal (m:strikeTLBR). */
  strikeDiagonalDown?: boolean;
}

/** Box properties (CT_BoxPr). */
export interface MathBoxProperties {
  /** Emulate an operator for line breaking (m:opEmu). */
  opEmu?: boolean;
  /** Never break inside (m:noBreak). */
  noBreak?: boolean;
  /** Aligned differential, no italics for the first run (m:diff). */
  diff?: boolean;
  /** Align this box with surrounding text (m:aln). */
  aln?: boolean;
}

/** Group-character properties (CT_GroupChrPr). */
export interface MathGroupCharacterProperties {
  /** Grouping character (m:chr). */
  chr?: string;
  /** Character position relative to the base (m:pos). */
  pos?: "top" | "bot";
  /** Base position relative to the character (m:vertJc). */
  vertJc?: "top" | "bot";
}

/** Phantom properties (CT_PhantPr). */
export interface MathPhantomProperties {
  /** Show the phantom content (m:show). */
  show?: boolean;
  /** Zero width (m:zeroWid). */
  zeroWid?: boolean;
  /** Zero ascent (m:zeroAsc). */
  zeroAsc?: boolean;
  /** Zero descent (m:zeroDesc). */
  zeroDesc?: boolean;
  /** Transparent — show without spacing (m:transp). */
  transp?: boolean;
}

/** Equation-array properties (CT_EqArrPr). */
export interface MathEquationArrayProperties {
  /** Row baseline alignment (m:baseJc). */
  baseJc?: MathBaselineAlignment;
  /** Distribute rows to the maximum width (m:maxDist). */
  maxDist?: boolean;
  /** Keep object distance between rows (m:objDist). */
  objDist?: boolean;
  /** Row spacing rule (m:rSpRule). */
  rSpRule?: MathSpacingRule;
  /** Row spacing (m:rSp). */
  rSp?: number;
}

/** Radical properties (CT_RadPr). */
export interface MathRadicalProperties {
  /** Hide the degree argument — square-root display (m:degHide). */
  degHide?: boolean;
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
  | {
      text: string;
      properties?: MathRunPropertiesOptions;
      /** Run text formatting (w:rPr inside m:r). */
      runProperties?: RunPropertiesOptions;
    }
  | {
      fraction: {
        numerator: MathInput[];
        denominator: MathInput[];
        fractionType?: string;
        /** Argument size scaling for the numerator (m:num/m:argPr/m:argSz). */
        numeratorArgumentSize?: number;
        /** Argument size scaling for the denominator (m:den/m:argPr/m:argSz). */
        denominatorArgumentSize?: number;
        /** Control-character formatting (m:fPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      superScript: {
        children: MathInput[];
        superScript: MathInput[];
        /** Control-character formatting (m:sSupPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      subScript: {
        children: MathInput[];
        subScript: MathInput[];
        /** Control-character formatting (m:sSubPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      subSuperScript: {
        children: MathInput[];
        subScript: MathInput[];
        superScript: MathInput[];
        /** Align sub/super scripts (m:sSubSupPr/m:alnScr). */
        alignScript?: boolean;
        /** Control-character formatting (m:sSubSupPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      preSubSuperScript: {
        children: MathInput[];
        subScript: MathInput[];
        superScript: MathInput[];
        /** Control-character formatting (m:sPrePr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      radical: {
        children: MathInput[];
        degree?: MathInput[];
        /** Radical properties (m:radPr). */
        properties?: MathRadicalProperties;
        /** Control-character formatting (m:radPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      sum: {
        children: MathInput[];
        subScript?: MathInput[];
        superScript?: MathInput[];
        properties?: MathNaryProperties;
        /** Control-character formatting (m:naryPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      integral: {
        children: MathInput[];
        subScript?: MathInput[];
        superScript?: MathInput[];
        properties?: MathNaryProperties;
        /** Control-character formatting (m:naryPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      limitLower: {
        children: MathInput[];
        limit: MathInput[];
        /** Control-character formatting (m:limLowPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      limitUpper: {
        children: MathInput[];
        limit: MathInput[];
        /** Control-character formatting (m:limUppPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      function: {
        children: MathInput[];
        name: MathInput[];
        /** Control-character formatting (m:funcPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      matrix: {
        /** Matrix rows; each cell is one m:e (a bare MathInput when the cell holds a single element). */
        rows: (MathInput | MathInput[])[][];
        /** Matrix properties (m:mPr). */
        properties?: MathMatrixProperties;
        /** Control-character formatting (m:mPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      roundBrackets:
        | MathInput[]
        | {
            children?: MathInput[];
            /** Multiple delimited elements (m:e+), separator-split groups. */
            elements?: MathInput[][];
            properties?: MathDelimiterProperties;
          };
    }
  | {
      curlyBrackets:
        | MathInput[]
        | {
            children?: MathInput[];
            /** Multiple delimited elements (m:e+), separator-split groups. */
            elements?: MathInput[][];
            properties?: MathDelimiterProperties;
          };
    }
  | {
      angledBrackets:
        | MathInput[]
        | {
            children?: MathInput[];
            /** Multiple delimited elements (m:e+), separator-split groups. */
            elements?: MathInput[][];
            properties?: MathDelimiterProperties;
          };
    }
  | {
      squareBrackets:
        | MathInput[]
        | {
            children?: MathInput[];
            /** Multiple delimited elements (m:e+), separator-split groups. */
            elements?: MathInput[][];
            properties?: MathDelimiterProperties;
          };
    }
  | {
      borderBox: {
        children: MathInput[];
        /** Border-box properties (m:borderBoxPr). */
        properties?: MathBorderBoxProperties;
        /** Control-character formatting (m:borderBoxPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      box: {
        children: MathInput[];
        /** Box properties (m:boxPr). */
        properties?: MathBoxProperties;
        /** Control-character formatting (m:boxPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      groupChr: {
        children: MathInput[];
        /** Group-character properties (m:groupChrPr). */
        properties?: MathGroupCharacterProperties;
        /** Control-character formatting (m:groupChrPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      phant: {
        children: MathInput[];
        /** Phantom properties (m:phantPr). */
        properties?: MathPhantomProperties;
        /** Control-character formatting (m:phantPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      eqArr: {
        rows: MathInput[][];
        /** Equation-array properties (m:eqArrPr). */
        properties?: MathEquationArrayProperties;
        /** Control-character formatting (m:eqArrPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      accent: {
        children: MathInput[];
        accentCharacter?: string;
        /** Control-character formatting (m:accPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    }
  | {
      bar: {
        children: MathInput[];
        type: "top" | "bot";
        /** Control-character formatting (m:barPr/m:ctrlPr). */
        ctrlPr?: RunPropertiesOptions;
      };
    };
