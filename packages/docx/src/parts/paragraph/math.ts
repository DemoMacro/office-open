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
  literal?: boolean;
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
  controlProperties?: RunPropertiesOptions;
}

/** N-ary operator limit location (ST_LimLoc). */
export type MathNaryLimitLocation = "subSup" | "undOvr";

/** Matrix/eqArr row-spacing rule (ST_SpacingRule, integer 0-4). */
export type MathSpacingRule = 0 | 1 | 2 | 3 | 4;

/** Matrix baseline alignment (s:ST_YAlign). */
export type MathBaselineAlignment = "inline" | "top" | "center" | "bottom" | "inside" | "outside";

/** N-ary operator properties (CT_NaryPr beyond character/subHide/supHide). */
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
  baseline?: MathBaselineAlignment;
  /** Hide placeholder glyphs in empty cells (m:plcHide). */
  hidePlaceholder?: boolean;
  /** Row spacing rule (m:rSpRule). */
  rowSpacingRule?: MathSpacingRule;
  /** Column spacing rule (m:cGpRule). */
  columnGapRule?: MathSpacingRule;
  /** Row spacing (m:rSp). */
  rowSpacing?: number;
  /** Column spacing (m:cSp). */
  columnSpacing?: number;
  /** Column gap (m:cGp). */
  columnGap?: number;
  /** Matrix column justifications (m:mcs). */
  columns?: MathMatrixColumnOptions[];
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
  operatorEmulation?: boolean;
  /** Never break inside (m:noBreak). */
  noBreak?: boolean;
  /** Aligned differential, no italics for the first run (m:diff). */
  differential?: boolean;
  /** Align this box with surrounding text (m:aln). */
  align?: boolean;
}

/** Group-character properties (CT_GroupChrPr). */
export interface MathGroupCharacterProperties {
  /** Grouping character (m:chr). */
  character?: string;
  /** Character position relative to the base (m:pos). */
  position?: "top" | "bot";
  /** Base position relative to the character (m:vertJc). */
  vertical?: "top" | "bot";
}

/** Phantom properties (CT_PhantPr). */
export interface MathPhantomProperties {
  /** Show the phantom content (m:show). */
  show?: boolean;
  /** Zero width (m:zeroWid). */
  zeroWidth?: boolean;
  /** Zero ascent (m:zeroAsc). */
  zeroAscent?: boolean;
  /** Zero descent (m:zeroDesc). */
  zeroDescent?: boolean;
  /** Transparent — show without spacing (m:transp). */
  transparent?: boolean;
}

/** Equation-array properties (CT_EqArrPr). */
export interface MathEquationArrayProperties {
  /** Row baseline alignment (m:baseJc). */
  baseline?: MathBaselineAlignment;
  /** Distribute rows to the maximum width (m:maxDist). */
  distributeRows?: boolean;
  /** Keep object distance between rows (m:objDist). */
  objectDistance?: boolean;
  /** Row spacing rule (m:rSpRule). */
  rowSpacingRule?: MathSpacingRule;
  /** Row spacing (m:rSp). */
  rowSpacing?: number;
}

/** Radical properties (CT_RadPr). */
export interface MathRadicalProperties {
  /** Hide the degree argument — square-root display (m:degHide). */
  hideDegree?: boolean;
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
      /** Trailing m:ctrlPr inside an argument element (m:e/m:num/m:den...,
       *  CT_OMathArg) — carried as a children-array member so its position
       *  round-trips. */
      argumentControlProperties: RunPropertiesOptions;
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
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      superScript: {
        children: MathInput[];
        superScript: MathInput[];
        /** Argument size scaling for the base (m:e/m:argPr/m:argSz). */
        baseArgumentSize?: number;
        /** Argument size scaling for the superscript (m:sup/m:argPr/m:argSz). */
        superScriptArgumentSize?: number;
        /** Control-character formatting (m:sSupPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      subScript: {
        children: MathInput[];
        subScript: MathInput[];
        /** Argument size scaling for the base (m:e/m:argPr/m:argSz). */
        baseArgumentSize?: number;
        /** Argument size scaling for the subscript (m:sub/m:argPr/m:argSz). */
        subScriptArgumentSize?: number;
        /** Control-character formatting (m:sSubPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      subSuperScript: {
        children: MathInput[];
        subScript: MathInput[];
        superScript: MathInput[];
        /** Argument size scaling for the base (m:e/m:argPr/m:argSz). */
        baseArgumentSize?: number;
        /** Argument size scaling for the subscript (m:sub/m:argPr/m:argSz). */
        subScriptArgumentSize?: number;
        /** Argument size scaling for the superscript (m:sup/m:argPr/m:argSz). */
        superScriptArgumentSize?: number;
        /** Align sub/super scripts (m:sSubSupPr/m:alnScr). */
        alignScript?: boolean;
        /** Control-character formatting (m:sSubSupPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      preSubSuperScript: {
        children: MathInput[];
        subScript: MathInput[];
        superScript: MathInput[];
        /** Argument size scaling for the subscript (m:sub/m:argPr/m:argSz). */
        subScriptArgumentSize?: number;
        /** Argument size scaling for the superscript (m:sup/m:argPr/m:argSz). */
        superScriptArgumentSize?: number;
        /** Argument size scaling for the base (m:e/m:argPr/m:argSz). */
        baseArgumentSize?: number;
        /** Control-character formatting (m:sPrePr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      radical: {
        children: MathInput[];
        degree?: MathInput[];
        /** Radical properties (m:radPr). */
        properties?: MathRadicalProperties;
        /** Control-character formatting (m:radPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      sum: {
        children: MathInput[];
        subScript?: MathInput[];
        superScript?: MathInput[];
        properties?: MathNaryProperties;
        /** Control-character formatting (m:naryPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      integral: {
        children: MathInput[];
        subScript?: MathInput[];
        superScript?: MathInput[];
        properties?: MathNaryProperties;
        /** Control-character formatting (m:naryPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      limitLower: {
        children: MathInput[];
        limit: MathInput[];
        /** Control-character formatting (m:limLowPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      limitUpper: {
        children: MathInput[];
        limit: MathInput[];
        /** Control-character formatting (m:limUppPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      function: {
        children: MathInput[];
        name: MathInput[];
        /** Control-character formatting (m:funcPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      matrix: {
        /** Matrix rows; each cell is one m:e (a bare MathInput when the cell holds a single element). */
        rows: (MathInput | MathInput[])[][];
        /** Matrix properties (m:mPr). */
        properties?: MathMatrixProperties;
        /** Control-character formatting (m:mPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
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
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      box: {
        children: MathInput[];
        /** Box properties (m:boxPr). */
        properties?: MathBoxProperties;
        /** Control-character formatting (m:boxPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      groupChr: {
        children: MathInput[];
        /** Group-character properties (m:groupChrPr). */
        properties?: MathGroupCharacterProperties;
        /** Control-character formatting (m:groupChrPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      phant: {
        children: MathInput[];
        /** Phantom properties (m:phantPr). */
        properties?: MathPhantomProperties;
        /** Control-character formatting (m:phantPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      eqArr: {
        rows: MathInput[][];
        /** Equation-array properties (m:eqArrPr). */
        properties?: MathEquationArrayProperties;
        /** Control-character formatting (m:eqArrPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      accent: {
        children: MathInput[];
        accentCharacter?: string;
        /** Control-character formatting (m:accPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    }
  | {
      bar: {
        children: MathInput[];
        type: "top" | "bot";
        /** Control-character formatting (m:barPr/m:ctrlPr). */
        controlProperties?: RunPropertiesOptions;
      };
    };
