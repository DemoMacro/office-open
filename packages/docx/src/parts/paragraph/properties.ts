import type { ShadingProperties } from "@shared/shading";
/**
 * Paragraph properties types for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPparagraphProperties.php
 *
 * @module
 */
import type { ChangedProperties } from "@shared/track-revision/track-revision";

import type { AlignmentType } from "./formatting/alignment";
import type { BordersOptions } from "./formatting/border";
import type { CnfConditionalOptions } from "./formatting/cnf-style";
import type { IndentProperties } from "./formatting/indent";
import type { SpacingProperties } from "./formatting/spacing";
import type { HeadingLevel } from "./formatting/style";
import type { TabStopDefinition } from "./formatting/tab-stop";
import type { FrameOptions } from "./frame/frame-properties";
import type { ParagraphRunOptions } from "./run/run";

/** Numbering applied as a revision (w:numPr/w:ins, CT_TrackChange). */
export interface NumberingInsertionOptions {
  /** Revision ID */
  id: string;
  /** Author of the change */
  author: string;
  /** Date of the change */
  date?: string;
}

/**
 * Vertical text alignment types for paragraphs.
 *
 * Specifies the vertical alignment of text within the paragraph.
 *
 * @publicApi
 */
export const TextAlignmentType = {
  /** Align text to the top */
  TOP: "top",
  /** Align text to the center */
  CENTER: "center",
  /** Align text to the baseline */
  BASELINE: "baseline",
  /** Align text to the bottom */
  BOTTOM: "bottom",
  /** Automatically determine vertical alignment */
  AUTO: "auto",
} as const;

/**
 * Textbox tight wrap types for paragraphs.
 *
 * Specifies how tightly text wraps around a textbox.
 *
 * @publicApi
 */
export const TextboxTightWrapType = {
  /** No tight wrapping */
  NONE: "none",
  /** Tight wrap on all lines */
  ALL_LINES: "allLines",
  /** Tight wrap on first and last lines */
  FIRST_AND_LAST_LINE: "firstAndLastLine",
  /** Tight wrap on first line only */
  FIRST_LINE_ONLY: "firstLineOnly",
  /** Tight wrap on last line only */
  LAST_LINE_ONLY: "lastLineOnly",
} as const;

/**
 * Paragraph style properties for numbering levels.
 *
 * These properties are used when defining paragraph styles within numbering level definitions.
 */
export interface LevelParagraphStylePropertiesOptions {
  /** Paragraph text alignment (left, right, center, justified, etc.) */
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  /** Whether to render text right-to-left for bidirectional languages */
  bidirectional?: boolean;
  /** Whether to insert a page break before this paragraph */
  pageBreakBefore?: boolean;
  /** Custom tab stop positions and alignments */
  tabStops?: TabStopDefinition[];
  /** Whether to display a horizontal line (thematic break) below the paragraph */
  thematicBreak?: boolean;
  /** Whether to prevent single lines at top/bottom of page (widow/orphan control), defaults to true */
  widowControl?: boolean;
  /** Whether to ignore spacing before/after when adjacent paragraphs have the same style */
  contextualSpacing?: boolean;
  /** Position in twips for a right-aligned tab stop */
  rightTabStop?: number;
  /** Position in twips for a left-aligned tab stop */
  leftTabStop?: number;
  /** Indentation settings for the paragraph */
  indent?: IndentProperties;
  /** Spacing before/after paragraph and between lines */
  spacing?: SpacingProperties;
  /**
   * Specifies that the paragraph (or at least part of it) should be rendered on the same page as the next paragraph when possible. If multiple paragraphs are to be kept together but they exceed a page, then the set of paragraphs begin on a new page and page breaks are used thereafter as needed.
   */
  keepNext?: boolean;
  /**
   * Specifies that all lines of the paragraph are to be kept on a single page when possible.
   */
  keepLines?: boolean;
  /** Frame properties for positioning the paragraph */
  frame?: FrameOptions;
  /** Whether to suppress line numbers for this paragraph */
  suppressLineNumbers?: boolean;
  /** Whether to allow word wrapping */
  wordWrap?: boolean;
  /** Whether to allow punctuation to extend beyond text margins */
  overflowPunctuation?: boolean;
  /**
   * This element specifies whether inter-character spacing shall automatically be adjusted between regions of numbers and regions of East Asian text in the current paragraph. These regions shall be determined by the Unicode character values of the text content within the paragraph.
   * This only works in Microsoft Word. It is not part of the ECMA-376 OOXML standard.
   */
  autoSpaceEastAsianText?: boolean;
  /** Whether to prevent text frames from overlapping */
  suppressOverlap?: boolean;
  /** Whether to disable automatic hyphenation for this paragraph */
  suppressAutoHyphens?: boolean;
  /** Whether to automatically adjust right indent for document grid */
  adjustRightInd?: boolean;
  /** Whether to snap the current paragraph to the document grid */
  snapToGrid?: boolean;
  /** Whether to swap left and right indent positions on odd pages for mirrored layouts */
  mirrorIndents?: boolean;
  /** Whether to use Kinsoku forbidden character overflow rules */
  kinsoku?: boolean;
  /** Whether to compress punctuation at the start of a line */
  topLinePunct?: boolean;
  /** Whether to automatically add space between East Asian and Latin text */
  autoSpaceDE?: boolean;
  /** Vertical text alignment within the paragraph */
  textAlignment?: (typeof TextAlignmentType)[keyof typeof TextAlignmentType];
  /** Textbox tight wrap setting */
  textboxTightWrap?: (typeof TextboxTightWrapType)[keyof typeof TextboxTightWrapType];
  /** Text direction for the paragraph (lr, rl, tb, tbV, rlV, lrV) */
  textDirection?: "lr" | "rl" | "tb" | "tbV" | "rlV" | "lrV";
  /** Outline level for table of contents and document outline (0-9) */
  outlineLevel?: number;
  /** HTML div ID reference */
  divId?: number;
  /** Conditional formatting style for table rows/cells */
  cnfStyle?: CnfConditionalOptions;
}

/**
 * Paragraph style properties options.
 *
 * These properties are used when defining paragraph styles and include
 * border, shading, and numbering options in addition to level properties.
 */
export type ParagraphStylePropertiesOptions = {
  /** Referenced style inside the style's own pPr (w:pStyle). */
  style?: string;
  /** Heading shorthand for the referenced style (w:pStyle val="HeadingN"). */
  heading?: (typeof HeadingLevel)[keyof typeof HeadingLevel];
  /** Border settings for the paragraph */
  border?: BordersOptions;
  /** Background shading/fill color for the paragraph */
  shading?: ShadingProperties;
  /** Numbering configuration for lists, or false to remove numbering */
  numbering?:
    | {
        /**
         * numId=0 with a pinned level: cancels the numbering inherited from
         * the style chain while keeping w:ilvl (false covers the level-less
         * form — a bare `<w:numId w:val="0"/>`).
         */
        none: true;
        /** Level in the numbering hierarchy (0-8). */
        level?: number;
      }
    | {
        /** Reference ID of the numbering definition to use */
        reference: string;
        /**
         * Level in the numbering hierarchy (0-8). Omitted when the source
         * w:numPr carried only w:numId (no w:ilvl) — the element then inherits
         * the level instead of pinning one.
         */
        level?: number;
        /** Instance number for multiple lists with same reference */
        instance?: number;
        /**
         * Whether the ListParagraph style is auto-applied (default true).
         * Set false when the paragraph style is authored explicitly (e.g.
         * round-tripped documents whose styles.xml may not define it).
         */
        autoStyle?: boolean;
        /** Numbering change tracking (CT_TrackChangeNumbering) */
        numberingChange?: {
          /** Original numbering value */
          original: string;
          /** Revision ID */
          id: string;
          /** Author of the change */
          author: string;
          /** Date of the change */
          date?: string;
        };
        /** Numbering applied as a revision (w:numPr/w:ins, CT_TrackChange) */
        insertion?: NumberingInsertionOptions;
      }
    | {
        /**
         * Level-only numPr: the source carried w:ilvl with no w:numId — the
         * numbering definition inherits from the style chain while this
         * element overrides the level (CT_NumPr leaves both children
         * optional). Emits `<w:numPr>` with only the w:ilvl.
         */
        levelOnly: true;
        /** Level in the numbering hierarchy (0-8). */
        level?: number;
      }
    | {
        /**
         * Track-change-only numPr: the source carried no w:numId — the whole
         * numbering property set is a tracked insertion whose numbering
         * resolves through the style chain. Emits `<w:numPr>` with only the
         * revision markers (no w:ilvl/w:numId).
         */
        revisionOnly: true;
        /** Numbering change tracking (CT_TrackChangeNumbering) */
        numberingChange?: {
          original: string;
          id: string;
          author: string;
          date?: string;
        };
        /** Numbering applied as a revision (w:numPr/w:ins, CT_TrackChange) */
        insertion?: NumberingInsertionOptions;
      }
    | false;
} & LevelParagraphStylePropertiesOptions;

export type ParagraphPropertiesOptionsBase = {
  /** Heading level (Heading1, Heading2, etc.) - applies predefined heading style */
  heading?: (typeof HeadingLevel)[keyof typeof HeadingLevel];
  /** Style ID to apply to this paragraph */
  style?: string;
  /** Bullet list configuration */
  bullet?: {
    /** Indentation level for the bullet (0-8) */
    level: number;
  };
  /**
   * Run properties to apply to all runs in the paragraph.
   * Reference: ECMA-376, 3rd Edition (June, 2011), Fundamentals and Markup Language Reference § 17.3.1.29.
   */
  run?: ParagraphRunOptions;
} & ParagraphStylePropertiesOptions;

export type ParagraphPropertiesChangeOptions = ChangedProperties & ParagraphPropertiesOptionsBase;

/**
 * Options for configuring paragraph properties.
 *
 * These options control all aspects of paragraph formatting including
 * alignment, spacing, indentation, borders, numbering, and more.
 *
 * Reference: http://officeopenxml.com/WPparagraphProperties.php
 */
export type ParagraphPropertiesOptions = {
  revision?: ParagraphPropertiesChangeOptions;
} & ParagraphPropertiesOptionsBase;
