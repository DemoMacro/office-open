import type { DocumentOptions } from "@office-open/docx";
import type { OcfManifestOptions } from "@office-open/odf";

import type { OdtAutomaticStyleOverlay } from "./automatic-styles";

export interface OdtDocumentOptions extends DocumentOptions {
  /** OCF manifest overlay; model-derived entries are appended when missing. */
  packageManifest?: OcfManifestOptions;
  /** Source automatic-style overlays keyed by stable name. */
  styleOverlays?: OdtAutomaticStyleOverlay[];
}

export type { DocumentOptions };
