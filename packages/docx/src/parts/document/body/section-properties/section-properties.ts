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
import type { HeaderFooterEntry } from "@parts/header-footer";
import type { ChangedAttributesProperties } from "@shared/track-revision/track-revision";
import type { SectionVerticalAlign } from "@shared/vertical-align";

import type { ColumnsAttributes } from "./properties/columns";
import type { DocGridAttributesProperties } from "./properties/doc-grid";
import type {
  EndnotePropertiesOptions,
  FootnotePropertiesOptions,
} from "./properties/footnote-endnote-properties";
import type { LineNumberAttributes } from "./properties/line-number";
import type { PageBordersOptions } from "./properties/page-borders";
import type { PageMarginAttributes } from "./properties/page-margin";
import type { PageNumberTypeAttributes } from "./properties/page-number";
import { PageOrientation } from "./properties/page-size";
import type { PageSizeAttributes } from "./properties/page-size";
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
  rsidRPr?: string;
  rsidDel?: string;
  rsidR?: string;
  rsidSect?: string;
  page?: {
    size?: PageSizeAttributes;
    margin?: PageMarginAttributes;
    pageNumbers?: PageNumberTypeAttributes;
    borders?: PageBordersOptions;
    textDirection?: (typeof PageTextDirectionType)[keyof typeof PageTextDirectionType];
  };
  grid?: DocGridAttributesProperties;
  headerWrapperGroup?: HeaderFooterGroup<HeaderFooterEntry>;
  footerWrapperGroup?: HeaderFooterGroup<HeaderFooterEntry>;
  lineNumbers?: LineNumberAttributes;
  titlePage?: boolean;
  verticalAlign?: SectionVerticalAlign;
  column?: ColumnsAttributes;
  type?: (typeof SectionType)[keyof typeof SectionType];
  noEndnote?: boolean;
  formProtection?: boolean;
  bidi?: boolean;
  rtlGutter?: boolean;
  paperSrc?: {
    first?: number;
    other?: number;
  };
  footnotePr?: FootnotePropertiesOptions;
  endnotePr?: EndnotePropertiesOptions;
  printerSettingsId?: string;
}

export type SectionPropertiesChangeOptions = ChangedAttributesProperties &
  SectionPropertiesOptionsBase;

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
