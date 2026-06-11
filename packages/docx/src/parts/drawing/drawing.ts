/**
 * Drawing module for WordprocessingML documents.
 *
 * @module
 */
import type { FillOptions } from "@office-open/core/drawingml";

import type { DocPropertiesOptions } from "./doc-properties/doc-properties";
import type { Floating } from "./floating";
import type { BlipEffectsOptions } from "./inline/graphic/graphic-data/pic/blip/blip-effects";
import type { TileOptions } from "./inline/graphic/graphic-data/pic/blip/tile";
import type { EffectListOptions } from "./inline/graphic/graphic-data/pic/effects/effect-list";
import type { OutlineOptions } from "./inline/graphic/graphic-data/pic/outline/outline";

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
