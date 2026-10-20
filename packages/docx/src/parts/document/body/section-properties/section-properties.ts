import type { LongHexNumber } from "@office-open/core";
/**
 * Section properties module for WordprocessingML documents.
 *
 * Section properties define page layout including page size, margins,
 * headers/footers, columns, and page numbering.
 *
 * Reference: http://officeopenxml.com/WPsection.php
 *
 * @module
 */
import type { ChangedProperties } from "@shared/track-revision/track-revision";
import type { SectionVerticalAlign } from "@shared/vertical-align";

import type { ColumnsProperties } from "./properties/columns";
import type { DocGridProperties } from "./properties/doc-grid";
import type {
  EndnotePropertiesOptions,
  FootnotePropertiesOptions,
} from "./properties/footnote-endnote-properties";
import type { LineNumberProperties } from "./properties/line-number";
import type { PageBordersOptions } from "./properties/page-borders";
import type { PageMarginProperties } from "./properties/page-margin";
import type { PageNumberTypeProperties } from "./properties/page-number";
import { PageOrientation } from "./properties/page-size";
import type { PageSizeProperties } from "./properties/page-size";
import { PageTextDirectionType } from "./properties/page-text-direction";
import type { SectionType } from "./properties/section-type";

/**
 * Header/footer group for specifying different headers/footers
 * for default, first, and even pages.
 */
export interface HeaderFooterGroup<T> {
  default?: T;
  first?: T;
  even?: T;
  odd?: T;
}

export interface SectionPropertiesOptionsBase {
  runPropertiesRsid?: LongHexNumber;
  deletionRsid?: LongHexNumber;
  additionRsid?: LongHexNumber;
  sectionRsid?: LongHexNumber;
  /**
   * Page size (w:pgSz). Three states: omitted (fresh generation defaults to
   * A4 portrait); a PageSizeProperties object (emits provided values); or
   * false (a parsed source with no w:pgSz is preserved by emitting nothing).
   */
  pageSize?: PageSizeProperties | false;
  /**
   * Page margins in twips (w:pgMar). Three states: omitted (fresh generation
   * defaults to Word's standard margins); a PageMarginProperties object; or
   * false (a parsed source with no w:pgMar is preserved by emitting nothing).
   */
  pageMargin?: PageMarginProperties | false;
  /** Page numbering (w:pgNumType). */
  pageNumberType?: PageNumberTypeProperties;
  /** Page borders (w:pgBorders). */
  pageBorders?: PageBordersOptions;
  /** Section text flow direction (w:textDirection). */
  textDirection?: (typeof PageTextDirectionType)[keyof typeof PageTextDirectionType];
  /**
   * Document grid. Omitted = fresh generation emits Word's CJK default line
   * grid (linePitch 312, type "lines"); a DocGridProperties object emits the
   * provided values; false = explicit off (emit nothing).
   */
  grid?: DocGridProperties | false;
  lineNumberType?: LineNumberProperties;
  titlePage?: boolean;
  verticalAlign?: SectionVerticalAlign;
  columns?: ColumnsProperties;
  /** Section-break placement: "nextPage"/"nextColumn" start there, "continuous" flows on, "evenPage"/"oddPage" next even/odd page. */
  type?: (typeof SectionType)[keyof typeof SectionType];
  /**
   * Tolerated misplaced settings element (`w:mirrorMargins`) inside a source
   * section; round-trip only — author document settings instead.
   */
  mirrorMargins?: boolean;
  /** Printer settings binary part path, e.g. word/printerSettings/printerSettings1.bin. */
  printerSettingsPath?: string;
  noEndnote?: boolean;
  formProtection?: boolean;
  bidi?: boolean;
  rtlGutter?: boolean;
  paperSrc?: {
    first?: number;
    other?: number;
  };
  /** Vendor-defined gutter marker (w:tmGutter, round-trip only). */
  tmGutter?: number;
  /** Vendor-defined section marker (w:tmSection, round-trip only). */
  tmSection?: TmSectionPropertiesOptions;
  footnoteProperties?: FootnotePropertiesOptions;
  endnoteProperties?: EndnotePropertiesOptions;
  /** Number of columns for the section footnotes (w15:footnoteColumns, Word 2013+). */
  footnoteColumns?: number;
  /**
   * Source header/footer reference emission order (round-trip only). Word may
   * interleave w:headerReference and w:footerReference; absent = canonical
   * header group (default/first/even) then footer group.
   */
  headerFooterReferenceOrder?: HeaderFooterReferenceSlot[];
}

/** One w:headerReference/w:footerReference slot identified by kind and type. */
export type HeaderFooterReferenceSlot =
  | "header-default"
  | "header-odd"
  | "header-first"
  | "header-even"
  | "footer-default"
  | "footer-odd"
  | "footer-first"
  | "footer-even";

/** Vendor extension attributes for w:tmSection (round-trip only). */
export interface TmSectionPropertiesOptions {
  /** Vendor-defined section marker height (w:h, round-trip only). */
  height?: number;
}

export type SectionPropertiesChangeOptions = ChangedProperties & SectionPropertiesOptionsBase;

export type SectionPropertiesOptions = {
  revision?: SectionPropertiesChangeOptions;
} & SectionPropertiesOptionsBase;

export const sectionMarginDefaults = {
  TOP: 1440,
  RIGHT: 1800,
  BOTTOM: 1440,
  LEFT: 1800,
  HEADER: 851,
  FOOTER: 992,
  GUTTER: 0,
};

export const sectionPageSizeDefaults = {
  WIDTH: 11_906,
  HEIGHT: 16_838,
  ORIENTATION: PageOrientation.PORTRAIT,
};
