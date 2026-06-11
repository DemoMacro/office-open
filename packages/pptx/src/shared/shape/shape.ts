import type { ShapeLockingOptions, UniversalMeasure } from "@office-open/core";
/**
 * Shape options type for PPTX.
 *
 * @module
 */
import type { AnimationOptions } from "@shared/animation/types";
import type { EffectsOptions } from "@shared/drawingml/effects";
import type { OutlineOptions } from "@shared/drawingml/outline";
import type { ShapePropertiesOptions } from "@shared/drawingml/shape-properties";

import type { TextBodyOptions } from "./text-body";

export interface ShapeStyleOptions {
  lineReference?: { index: number; color?: string };
  fillReference?: { index: number; color?: string };
  effectReference?: { index: number; color?: string };
  fontReference?: { index: number; color?: string };
}

export interface ShapeOptions {
  id?: number;
  name?: string;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  geometry?: string;
  fill?: ShapePropertiesOptions["fill"];
  outline?: OutlineOptions;
  effects?: EffectsOptions;
  flipHorizontal?: boolean;
  rotation?: number;
  textBody?: TextBodyOptions;
  animation?: AnimationOptions;
  locking?: ShapeLockingOptions;
  placeholder?: "title" | "body" | "subTitle" | "sldNum" | "dt" | "ftr" | "hdr" | "obj";
  placeholderIndex?: number;
  useBackgroundFill?: boolean;
  isPhoto?: boolean;
  userDrawn?: boolean;
  hasCustomPrompt?: boolean;
  style?: ShapeStyleOptions;
  blackWhiteMode?:
    | "clr"
    | "auto"
    | "gray"
    | "ltGray"
    | "invGray"
    | "grayWhite"
    | "blackGray"
    | "blackWhite"
    | "black"
    | "white"
    | "hidden";
}
