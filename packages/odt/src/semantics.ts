import type { DocumentOptions } from "@office-open/docx";
import type { OdfPackageOverlayOptions } from "@office-open/odf";

import type { OdtEmbeddedObjectOptions } from "./package";

/** Canonical document model plus the unified ODF package overlay. */
export interface OdtDocumentOptions extends DocumentOptions, OdfPackageOverlayOptions {
  embeddedObjects?: OdtEmbeddedObjectOptions[];
}

export type { DocumentOptions };
