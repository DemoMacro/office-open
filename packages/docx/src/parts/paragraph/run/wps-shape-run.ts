/**
 * WPS shape run types for WordprocessingML documents.
 *
 * @module
 */
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { WpsShapeCoreOptions } from "@parts/drawing/inline/graphic/graphic-data/wps";
import type { MediaTransformation } from "@shared/media";

import type { Floating } from "../../drawing";

export * from "@parts/drawing/inline/graphic/graphic-data/wps/body-properties";

interface CoreShapeOptions {
  transformation: MediaTransformation;
  floating?: Floating;
  altText?: DocPropertiesOptions;
}

/**
 * @publicApi
 */
export type IWpsShapeOptions = WpsShapeCoreOptions & { type: "wps" } & CoreShapeOptions;
