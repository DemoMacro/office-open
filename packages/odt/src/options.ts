import type { DocumentOptions } from "@office-open/docx";
import type { PackageOptions } from "@office-open/odf";

import type { OdtEmbeddedObjectOptions } from "./embedded-objects";

/** Canonical document model plus the unified ODF package overlay. */
export interface OdtDocumentOptions extends DocumentOptions, PackageOptions {
  embeddedObjects?: OdtEmbeddedObjectOptions[];
}

export type { DocumentOptions };
