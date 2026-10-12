/**
 * Math types for Office MathML (OMML).
 *
 * Provides the canonical input type for math content and run properties.
 * No XmlComponent dependencies — pure type definitions.
 *
 * @module
 */

import type { ChangedProperties } from "@shared/track-revision/track-revision";

import type { RunPropertiesOptions } from "./run/properties";
import type { RunPropertiesChangeOptions } from "./run/properties";
import type { RunOptions } from "./run/run";

// ── MathRunProperties ──

/** Math alphanumerics font class (ST_Script). */
export type MathScriptType =
  | "roman"
  | "script"
  | "fraktur"
  | "double-struck"
  | "sans-serif"
  | "monospace";

/** Math run style (ST_Style). */
export type MathStyleType = "plain" | "bold" | "italic" | "boldItalic";

export interface MathRunPropertiesOptions {
  literal?: boolean;
  normal?: boolean;
  script?: MathScriptType;
  style?: MathStyleType;
  breakAlignment?: number;
  align?: boolean;
}

/** Argument properties (`m:argPr`, CT_OMathArgPr). */
export interface MathArgumentPropertiesOptions {
  /** Argument size scale (`m:argSz/@m:val`; 1 = 100%, 2 = 200%). */
  size?: number;
}

/** Control-character formatting (`m:ctrlPr`, CT_CtrlPr). */
export type MathControlPropertiesOptions =
  | { runProperties?: RunPropertiesOptions }
  | {
      insertion: ChangedProperties & {
        deletion?: RunPropertiesChangeOptions;
        runProperties?: RunPropertiesOptions;
      };
    }
  | {
      deletion: ChangedProperties & {
        runProperties?: RunPropertiesOptions;
      };
    };

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
  controlProperties?: MathControlPropertiesOptions;
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
  /** Hide the subscript placeholder (m:subHide). */
  hideSubscript?: boolean;
  /** Hide the superscript placeholder (m:supHide). */
  hideSuperscript?: boolean;
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
      /** Source `xml:space="preserve"` marker on m:t; round-trip only. */
      preserveSpace?: boolean;
      properties?: MathRunPropertiesOptions;
      /** Run text formatting (w:rPr inside m:r). */
      runProperties?: RunPropertiesOptions;
    }
  | {
      /** Mixed Word run content (w:EG_RunInnerContent mixed with m:t). */
      children: NonNullable<RunOptions["children"]>;
      properties?: MathRunPropertiesOptions;
      /** Run text formatting (w:rPr inside m:r). */
      runProperties?: RunPropertiesOptions;
    }
  | {
      /** Trailing m:ctrlPr inside an argument element (m:e/m:num/m:den...,
       *  CT_OMathArg) — carried as a children-array member so its position
       *  round-trips. */
      argumentControlProperties: MathControlPropertiesOptions;
    }
  | {
      fraction: {
        numerator: MathInput[];
        denominator: MathInput[];
        fractionType?: string;
        /** Numerator argument properties (m:num/m:argPr). */
        numeratorArgumentProperties?: MathArgumentPropertiesOptions;
        /** Denominator argument properties (m:den/m:argPr). */
        denominatorArgumentProperties?: MathArgumentPropertiesOptions;
        /** Control-character formatting (m:fPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      superScript: {
        children: MathInput[];
        superScript: MathInput[];
        /** Base argument properties (m:e/m:argPr). */
        baseArgumentProperties?: MathArgumentPropertiesOptions;
        /** Superscript argument properties (m:sup/m:argPr). */
        superScriptArgumentProperties?: MathArgumentPropertiesOptions;
        /** Control-character formatting (m:sSupPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      subScript: {
        children: MathInput[];
        subScript: MathInput[];
        /** Base argument properties (m:e/m:argPr). */
        baseArgumentProperties?: MathArgumentPropertiesOptions;
        /** Subscript argument properties (m:sub/m:argPr). */
        subScriptArgumentProperties?: MathArgumentPropertiesOptions;
        /** Control-character formatting (m:sSubPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      subSuperScript: {
        children: MathInput[];
        subScript: MathInput[];
        superScript: MathInput[];
        /** Base argument properties (m:e/m:argPr). */
        baseArgumentProperties?: MathArgumentPropertiesOptions;
        /** Subscript argument properties (m:sub/m:argPr). */
        subScriptArgumentProperties?: MathArgumentPropertiesOptions;
        /** Superscript argument properties (m:sup/m:argPr). */
        superScriptArgumentProperties?: MathArgumentPropertiesOptions;
        /** Align sub/super scripts (m:sSubSupPr/m:alnScr). */
        alignScript?: boolean;
        /** Control-character formatting (m:sSubSupPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      preSubSuperScript: {
        children: MathInput[];
        subScript: MathInput[];
        superScript: MathInput[];
        /** Subscript argument properties (m:sub/m:argPr). */
        subScriptArgumentProperties?: MathArgumentPropertiesOptions;
        /** Superscript argument properties (m:sup/m:argPr). */
        superScriptArgumentProperties?: MathArgumentPropertiesOptions;
        /** Base argument properties (m:e/m:argPr). */
        baseArgumentProperties?: MathArgumentPropertiesOptions;
        /** Control-character formatting (m:sPrePr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      radical: {
        children: MathInput[];
        degree?: MathInput[];
        /** Radical properties (m:radPr). */
        properties?: MathRadicalProperties;
        /** Control-character formatting (m:radPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      sum: {
        children: MathInput[];
        subScript?: MathInput[];
        superScript?: MathInput[];
        properties?: MathNaryProperties;
        /** Control-character formatting (m:naryPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      integral: {
        children: MathInput[];
        subScript?: MathInput[];
        superScript?: MathInput[];
        properties?: MathNaryProperties;
        /** Control-character formatting (m:naryPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      /** Any OMML n-ary operator not represented by the sum/integral sugar. */
      nary: {
        /** Exact operator character (m:naryPr/m:chr, for example ∏). */
        character: string;
        children: MathInput[];
        subScript?: MathInput[];
        superScript?: MathInput[];
        properties?: MathNaryProperties;
        /** Control-character formatting (m:naryPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      limitLower: {
        children: MathInput[];
        limit: MathInput[];
        /** Control-character formatting (m:limLowPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      limitUpper: {
        children: MathInput[];
        limit: MathInput[];
        /** Control-character formatting (m:limUppPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      function: {
        children: MathInput[];
        name: MathInput[];
        /** Control-character formatting (m:funcPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      matrix: {
        /** Matrix rows; each cell is one m:e (a bare MathInput when the cell holds a single element). */
        rows: (MathInput | MathInput[])[][];
        /** Matrix properties (m:mPr). */
        properties?: MathMatrixProperties;
        /** Control-character formatting (m:mPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
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
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      box: {
        children: MathInput[];
        /** Box properties (m:boxPr). */
        properties?: MathBoxProperties;
        /** Control-character formatting (m:boxPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      groupChr: {
        children: MathInput[];
        /** Group-character properties (m:groupChrPr). */
        properties?: MathGroupCharacterProperties;
        /** Control-character formatting (m:groupChrPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      phant: {
        children: MathInput[];
        /** Phantom properties (m:phantPr). */
        properties?: MathPhantomProperties;
        /** Control-character formatting (m:phantPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      eqArr: {
        rows: MathInput[][];
        /** Equation-array properties (m:eqArrPr). */
        properties?: MathEquationArrayProperties;
        /** Control-character formatting (m:eqArrPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      accent: {
        children: MathInput[];
        accentCharacter?: string;
        /** Control-character formatting (m:accPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    }
  | {
      bar: {
        children: MathInput[];
        type: "top" | "bot";
        /** Control-character formatting (m:barPr/m:ctrlPr). */
        controlProperties?: MathControlPropertiesOptions;
      };
    };
