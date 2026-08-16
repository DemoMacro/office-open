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
}

export interface SectionPropertiesOptionsBase {
  runPropertiesRsid?: string;
  deletionRsid?: string;
  rsid?: string;
  sectionRsid?: string;
  /** Page size (w:pgSz). Defaults to A4 portrait. */
  pageSize?: PageSizeProperties;
  /** Page margins in twips (w:pgMar). Defaults to Word's standard margins. */
  pageMargin?: PageMarginProperties;
  /** Page numbering (w:pgNumType). */
  pageNumberType?: PageNumberTypeProperties;
  /** Page borders (w:pgBorders). */
  pageBorders?: PageBordersOptions;
  /** Section text flow direction (w:textDirection). */
  textDirection?: (typeof PageTextDirectionType)[keyof typeof PageTextDirectionType];
  /**
   * Document grid. Three states: omitted (fresh generation emits Word's CJK
   * default line grid — linePitch 312, type "lines"); a DocGridProperties object
   * (emits provided values, e.g. from a parsed source); or false (explicit off —
   * a parsed source with no w:docGrid is preserved by emitting nothing).
   */
  grid?: DocGridProperties | false;
  lineNumberType?: LineNumberProperties;
  titlePage?: boolean;
  verticalAlign?: SectionVerticalAlign;
  columns?: ColumnsProperties;
  type?: (typeof SectionType)[keyof typeof SectionType];
  noEndnote?: boolean;
  formProtection?: boolean;
  bidi?: boolean;
  rtlGutter?: boolean;
  paperSrc?: {
    first?: number;
    other?: number;
  };
  footnoteProperties?: FootnotePropertiesOptions;
  endnoteProperties?: EndnotePropertiesOptions;
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
