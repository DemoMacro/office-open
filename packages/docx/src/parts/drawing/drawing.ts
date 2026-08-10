/**
 * Drawing module for WordprocessingML documents.
 *
 * @module
 */
import type {
  BlipEffectsOptions,
  EffectListOptions,
  FillOptions,
  OutlineOptions,
  TileOptions,
} from "@office-open/core/drawingml";

import type { DocPropertiesOptions } from "./doc-properties/doc-properties";
import type { Floating } from "./floating";

/**
 * Distance options for drawing elements.
 *
 * Specifies the margins around a drawing element.
 */
export interface Distance {
  distT?: number;
  distB?: number;
  distL?: number;
  distR?: number;
}

/**
 * Options for configuring a drawing element.
 *
 * @see {@link Drawing}
 */
export interface DrawingOptions {
  floating?: Floating;
  docProperties?: DocPropertiesOptions;
  outline?: OutlineOptions;
  fill?: FillOptions;
  effects?: EffectListOptions;
  blipEffects?: BlipEffectsOptions;
  tile?: TileOptions;
}
