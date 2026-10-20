import type { DocumentOptions } from "@office-open/docx";
import type { PackageOptions } from "@office-open/odf";

import type { OdtEmbeddedObjectOptions } from "./embedded-objects";

/** Canonical document model plus the unified ODF package overlay. */
export interface OdtDocumentOptions extends DocumentOptions, PackageOptions {
  embeddedObjects?: OdtEmbeddedObjectOptions[];
  indexAutoMarkFile?: OdtIndexAutoMarkFileOptions;
  sectionSources?: OdtSectionSourceOptions[];
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

export type { DocumentOptions };
