import type { TextBodyOptions } from "@office-open/core";
import type { PackageOptions } from "@office-open/odf";
import type {
  MasterDefinition,
  PresentationOptions,
  SlideOptions,
  SlideSize,
} from "@office-open/pptx";

import type { OdpEmbeddedObjectOptions } from "./embedded-objects";
import type { OdpSlideNodeOptions } from "./slide-nodes";
import type { OdpTransitionOptions } from "./transition";

/** Canonical presentation model plus the unified ODF package overlay. */
export interface OdpPresentationOptions
  extends Omit<PresentationOptions, "masters">, PackageOptions {
  embeddedObjects?: OdpEmbeddedObjectOptions[];
  masters?: OdpMasterDefinition[];
  slides?: OdpSlideOptions[];
  /** All ODF page layouts; the first referenced layout also maps to presentation size. */
  pageLayouts?: OdpPageLayoutOptions[];
}

/** ODF page layout expressed with the canonical presentation size type. */
export interface OdpPageLayoutOptions {
  name: string;
  size?: SlideSize;
}

/** Typed ODF master-page header/footer content without a PPTX equivalent. */
export interface OdpMasterHeaderFooterOptions {
  display?: boolean;
  paragraphs?: TextBodyOptions["paragraphs"];
}

/** Canonical presentation master plus ODF-only master-page text regions. */
export interface OdpMasterDefinition extends MasterDefinition {
  /** ODF page-layout referenced by this master page. */
  pageLayoutName?: string;
  header?: OdpMasterHeaderFooterOptions;
  headerLeft?: OdpMasterHeaderFooterOptions;
  headerFirst?: OdpMasterHeaderFooterOptions;
  footer?: OdpMasterHeaderFooterOptions;
  footerLeft?: OdpMasterHeaderFooterOptions;
  footerFirst?: OdpMasterHeaderFooterOptions;
}

/** PPTX slide model plus fully typed ODF-only nodes in source order. */
export interface OdpSlideOptions extends Omit<SlideOptions, "transition"> {
  animations?: SlideOptions["animations"];
  transition?: OdpTransitionOptions | string;
  drawingPageStyleName?: string;
  /** ODF presentation page layout referenced directly by this slide. */
  presentationPageLayoutName?: string;
  odpNodes?: OdpSlideNodeOptions[];
}

export type { PresentationOptions, SlideOptions } from "@office-open/pptx";
