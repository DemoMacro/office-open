import type { ObjectElementOptions } from "@parts/object";

/**
 * Run module for WordprocessingML documents.
 *
 * A run is a region of text with a common set of properties. It is the primary
 * unit of inline content in a paragraph.
 *
 * Reference: http://officeopenxml.com/WPtext.php
 *
 * @module
 */
import type {
  AnnotationReference,
  CarriageReturn,
  ContinuationSeparator,
  DayLong,
  DayShort,
  EndnoteReference,
  FootnoteReferenceElement,
  LastRenderedPageBreak,
  MonthLong,
  MonthShort,
  NoBreakHyphen,
  PageNumberElement,
  Separator,
  SoftHyphen,
  Tab,
  YearLong,
  YearShort,
} from "./empty-children";
import type { ParagraphRunPropertiesOptions, RunPropertiesOptions } from "./properties";

interface RunOptionsBase {
  children?: (
    | (typeof PageNumber)[keyof typeof PageNumber]
    | string
    | AnnotationReference
    | CarriageReturn
    | ContinuationSeparator
    | DayLong
    | DayShort
    | EndnoteReference
    | FootnoteReferenceElement
    | LastRenderedPageBreak
    | MonthLong
    | MonthShort
    | NoBreakHyphen
    | PageNumberElement
    | Separator
    | SoftHyphen
    | Tab
    | YearLong
    | YearShort
    | { object: ObjectElementOptions }
    | Record<string, unknown>
  )[];
  break?: number;
  text?: string;
}

/**
 * Options for creating a Run element.
 *
 * The run element specifies a region of text with a common set of properties.
 * The children property can contain various inline content elements.
 *
 * @see {@link Run}
 */
export type RunOptions = RunOptionsBase &
  RunPropertiesOptions & {
    /** Revision save ID for the run (w:rsidR, hex string e.g. "00123456"). */
    rsid?: string;
    /** Revision save ID for run properties (w:rsidRPr, hex string). */
    runPropertiesRsid?: string;
    /** Revision save ID when run was deleted (w:rsidDel, hex string). */
    deletionRsid?: string;
  };

export type ParagraphRunOptions = RunOptionsBase & ParagraphRunPropertiesOptions;

/**
 * Constants for page number field types.
 *
 * These values are used to insert dynamic page number fields into a document.
 *
 * Reference: http://officeopenxml.com/WPfields.php
 *
 * @publicApi
 */
export const PageNumber = {
  /** Inserts the current page number */
  CURRENT: "CURRENT",
  /** Inserts the total number of pages in the document */
  TOTAL_PAGES: "TOTAL_PAGES",
  /** Inserts the total number of pages in the current section */
  TOTAL_PAGES_IN_SECTION: "TOTAL_PAGES_IN_SECTION",
  /** Inserts the current section number */
  CURRENT_SECTION: "SECTION",
} as const;
