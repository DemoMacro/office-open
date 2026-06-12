/**
 * WPG group run types for WordprocessingML documents.
 *
 * @module
 */
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { GroupChildMediaData, MediaTransformation } from "@shared/media";

import type { Floating } from "../../drawing";

export * from "@parts/drawing/inline/graphic/graphic-data/wps/body-properties";

interface CoreGroupOptions {
  children: GroupChildMediaData[];
  transformation: MediaTransformation;
  floating?: Floating;
  altText?: DocPropertiesOptions;
}

/**
 * @publicApi
 */
export type WpgGroupRunOptions = { type: "wpg" } & CoreGroupOptions;
