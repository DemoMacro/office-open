import type { PackageOptions } from "@office-open/odf";
import type { PresentationOptions, SlideOptions } from "@office-open/pptx";

import type { OdpEmbeddedObjectOptions } from "./embedded-objects";
import type { OdpSlideNodeOptions } from "./slide-nodes";

/** Canonical presentation model plus the unified ODF package overlay. */
export interface OdpPresentationOptions extends PresentationOptions, PackageOptions {
  embeddedObjects?: OdpEmbeddedObjectOptions[];
  slides?: OdpSlideOptions[];
}

/** PPTX slide model plus fully typed ODF-only nodes in source order. */
export interface OdpSlideOptions extends SlideOptions {
  odpNodes?: OdpSlideNodeOptions[];
}

export type { PresentationOptions, SlideOptions } from "@office-open/pptx";
