import type { FillOptions } from "@office-open/core/drawingml";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
/**
 * WPS shape types for WordprocessingML documents.
 *
 * Provides type definitions for WPS shape elements in drawingML.
 *
 * @module
 */
import type { MediaDataTransformation } from "@shared/media";

import type { CustomGeometryOptions } from "../pic/custom-geometry/custom-geometry";
import type { EffectDagOptions } from "../pic/effects/effect-dag";
import type { EffectListOptions } from "../pic/effects/effect-list";
import type { OutlineOptions } from "../pic/outline/outline";
import type { Scene3DOptions } from "../pic/three-d/scene-3d";
import type { Shape3DOptions } from "../pic/three-d/shape-3d";
import type { BodyPropertiesOptions } from "./body-properties";
import type { NonVisualShapePropertiesOptions } from "./non-visual-shape-properties";

export interface WpsShapeCoreOptions {
  children: (ParagraphOptions | string)[];
  nonVisualProperties?: NonVisualShapePropertiesOptions;
  bodyProperties?: BodyPropertiesOptions;
  outline?: OutlineOptions;
  fill?: FillOptions;
  customGeometry?: CustomGeometryOptions;
  effectDag?: EffectDagOptions;
  effects?: EffectListOptions;
  scene3d?: Scene3DOptions;
  shape3d?: Shape3DOptions;
}

export type WpsShapeOptions = WpsShapeCoreOptions & {
  transformation: MediaDataTransformation;
};
