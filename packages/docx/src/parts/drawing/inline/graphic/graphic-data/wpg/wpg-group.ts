import type { FillOptions } from "@office-open/core/drawingml";
/**
 * WPG group shape types for WordprocessingML documents.
 *
 * Provides type definitions for group shape elements in drawingML.
 *
 * @module
 */
import type { MediaDataTransformation } from "@shared/media";

import type { EffectListOptions } from "../pic/effects/effect-list";

export type GroupChild = unknown;

/**
 * Child coordinate offset (CT_Point2D).
 */
export interface ChildOffset {
  x: number;
  y: number;
}

/**
 * Child coordinate extent (CT_PositiveSize2D).
 */
export interface ChildExtent {
  cx: number;
  cy: number;
}

export interface WpgGroupCoreOptions {
  children: GroupChild[];
}

export type WpgGroupOptions = WpgGroupCoreOptions & {
  transformation: MediaDataTransformation;
  /** Child coordinate offset (chOff) */
  childOffset?: ChildOffset;
  /** Child coordinate extent (chExt) */
  childExtent?: ChildExtent;
  /** Group fill */
  fill?: FillOptions;
  /** Group effects */
  effects?: EffectListOptions;
};
