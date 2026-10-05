import type { DocumentOptions } from "@office-open/docx";
import type { OcfManifestOptions } from "@office-open/odf";

import type { OdtAutomaticStyleOverlay } from "./automatic-styles";
import type { OdtEmbeddedObjectOptions, OdtPackageMemberOptions } from "./package";

export interface OdtDocumentOptions extends DocumentOptions {
  /** OCF manifest overlay; model-derived entries are appended when missing. */
  packageManifest?: OcfManifestOptions;
  /** Standalone ODF subdocuments retained where docx has no equivalent run. */
  embeddedObjects?: OdtEmbeddedObjectOptions[];
  /** Source-only OCF members retained outside modeled package content. */
  packageMembers?: OdtPackageMemberOptions[];
  /** Source automatic-style overlays keyed by stable name. */
  styleOverlays?: OdtAutomaticStyleOverlay[];
}

export type { DocumentOptions };
