import type { DocumentOptions, SectionOptions } from "@office-open/docx";
import type { PackageOptions, StyleColumnsOptions, StylePropertyOptions } from "@office-open/odf";

import type { OdtCalculationSettingsOptions } from "./calculation-settings";
import type { OdtEmbeddedObjectOptions } from "./embedded-objects";

/** Canonical document model plus the unified ODF package overlay. */
export interface OdtDocumentOptions extends DocumentOptions, PackageOptions {
  embeddedObjects?: OdtEmbeddedObjectOptions[];
  calculationSettings?: OdtCalculationSettingsOptions;
  indexAutoMarkFile?: OdtIndexAutoMarkFileOptions;
  sectionSources?: OdtSectionSourceOptions[];
  /** All ODF page layouts; the first applicable layout also maps to section properties. */
  pageLayouts?: OdtPageLayoutOptions[];
  /** ODF master pages and their canonical header/footer regions. */
  masterPages?: OdtMasterPageOptions[];
}

/** ODF page layout expressed with canonical section page properties. */
export interface OdtPageLayoutOptions {
  name: string;
  properties?: SectionOptions["properties"];
  /** Exact ODF column descriptor; canonical section columns are projected from it. */
  columns?: StyleColumnsOptions;
  /** Typed ODF page-layout properties without a canonical section equivalent. */
  pageLayoutProperties?: Extract<StylePropertyOptions, { name: "style:page-layout-properties" }>;
}

/** ODF master page expressed with canonical section header/footer regions. */
export interface OdtMasterPageOptions {
  name: string;
  pageLayoutName?: string;
  headers?: SectionOptions["headers"];
  footers?: SectionOptions["footers"];
}

/** External alphabetical-index auto-mark file linked from ODT body declarations. */
export interface OdtIndexAutoMarkFileOptions {
  href: string;
}

/** External source metadata attached to a named ODT text section. */
export interface OdtSectionSourceOptions {
  href: string;
  filterName?: string;
  show?: "embed";
  sectionName?: string;
  targetSectionName?: string;
}

export type { DocumentOptions, OdtCalculationSettingsOptions };
