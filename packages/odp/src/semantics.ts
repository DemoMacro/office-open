import type { OcfManifestOptions } from "@office-open/odf";
import type { PresentationOptions } from "@office-open/pptx";

export interface OdpDocumentOptions extends PresentationOptions {
  /** OCF manifest overlay; model-derived entries are appended when missing. */
  packageManifest?: OcfManifestOptions;
  /** Source automatic-style overlays keyed by stable name. */
  styleOverlays?: OdpStyleOverlay[];
}

export interface OdpStyleOverlay {
  /** Stable source style name. */
  name: string;
  /** RNG style family. */
  family: string;
  /** RNG-valid property elements not represented by canonical slide options. */
  properties: { name: string; attributes: Record<string, string | number | boolean> }[];
}

export type { PresentationOptions, SlideOptions } from "@office-open/pptx";
