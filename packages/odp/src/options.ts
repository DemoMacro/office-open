import type { PackageOptions } from "@office-open/odf";
import type { PresentationOptions, SlideOptions } from "@office-open/pptx";

import type { OdpEmbeddedObjectOptions } from "./embedded-objects";
import type { OdpOpaqueSlideNodeOptions } from "./opaque-node";

/** Canonical presentation model plus the unified ODF package overlay. */
export interface OdpPresentationOptions extends PresentationOptions, PackageOptions {
  embeddedObjects?: OdpEmbeddedObjectOptions[];
  slides?: OdpSlideOptions[];
}

/** PPTX slide model plus ODP-only nodes that have no canonical equivalent. */
export interface OdpSlideOptions extends SlideOptions {
  opaqueNodes?: OdpOpaqueSlideNodeOptions[];
}

export type { PresentationOptions, SlideOptions } from "@office-open/pptx";
