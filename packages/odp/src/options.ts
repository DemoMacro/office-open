import type { PackageOptions } from "@office-open/odf";
import type { PresentationOptions } from "@office-open/pptx";

import type { OdpEmbeddedObjectOptions } from "./embedded-objects";

/** Canonical presentation model plus the unified ODF package overlay. */
export interface OdpPresentationOptions extends PresentationOptions, PackageOptions {
  embeddedObjects?: OdpEmbeddedObjectOptions[];
}

export type { PresentationOptions, SlideOptions } from "@office-open/pptx";
