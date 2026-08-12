// ── Shared: cross-part types ──

// Presentation-level types
export {
  type PresentationOptions,
  type ShowOptions,
  type SlideOptions,
  type SlideCommentOptions,
  type MasterDefinition,
  type LayoutDefinition,
  type LayoutPlaceholderOptions,
  type MasterChild,
  type SlideSize,
} from "./file";

// DrawingML — re-exports from core
export { createOutline, type OutlineOptions } from "@office-open/core/drawingml";
export {
  createGradientFill,
  createGradientStop,
  type GradientFillOptions as CoreGradientFillOptions,
  type GradientStop,
  type PathShade,
  type TileFlipMode,
} from "@office-open/core/drawingml";
export {
  type LineCap,
  type CompoundLine,
  type PenAlignment,
  type PresetDash,
  type LineJoin,
} from "@office-open/core/drawingml";
export { createScene3D, type Scene3DOptions } from "@office-open/core/drawingml";
export {
  createShape3D,
  type Shape3DOptions,
  type PresetMaterial,
} from "@office-open/core/drawingml";
export {
  createBevel,
  createBottomBevel,
  type BevelOptions,
  type BevelPreset,
} from "@office-open/core/drawingml";
export { createEffectList, type EffectListOptions } from "@office-open/core/drawingml";
export { createColorElement } from "@office-open/core/drawingml";
export { createColorTransforms, type ColorTransformOptions } from "@office-open/core/drawingml";

// DrawingML — fill API
export {
  buildFill,
  extractBlipFillMedia,
  type BlipFillConfigOptions,
  type BlipFillMediaData,
  type FillOptions,
  type GradientStopOptions,
} from "./drawingml/fill";

// DrawingML — local aliases + core re-exports
export type { Transform2DOptions } from "./drawingml/transform-2d";
export { stringifyPresetGeometry } from "@office-open/core/drawingml";
export type { ShapePropertiesOptions } from "@office-open/core/drawingml";

// Shape types — text/run types re-exported from core DrawingML
export type { ShapeOptions } from "./shape/shape";
export type {
  TextBodyOptions,
  ParagraphDescriptorOptions,
  RunOptions,
} from "@office-open/core/drawingml";
export {
  type UnderlineStyle,
  type StrikeStyle,
  type TextCapitalization,
  type RunPropertiesOptions,
  type HyperlinkOptions,
} from "@office-open/core/drawingml";
export type { TextAlignment, ParagraphPropertiesOptions } from "@office-open/core/drawingml";
export type { GroupOptions, GroupShapeOptions } from "./shape/group-shape";
export type { LineShapeOptions, ConnectorOptions, ConnectorShapeOptions } from "./shape/line-shape";

// Media
export { Media } from "@office-open/core";
export { createTransformation, type MediaTransformation } from "./media/media";
export type { MediaData, MediaDataTransformation } from "./media/data";
export type { VideoFrameOptions, VideoType, PosterType } from "./media/video-frame";
export type { AudioFrameOptions, AudioType } from "./media/audio-frame";

// Table
export type { TableOptions } from "./table/table-frame";
export type { TableRowOptions } from "./table/table-row";
export type { VerticalAlignment, TableCellOptions } from "./table/table-cell";
export type { CellBorderOptions } from "./table/table-cell-properties";

// Theme
export {
  createThemeXml,
  type ThemeOptions,
  type ColorSchemeOptions,
  type FontSchemeOptions,
} from "./theme";

// Header-footer
export type { SlideHeaderFooterOptions } from "./header-footer";

// Picture
export type { PictureOptions } from "./picture";

// Background — re-export from parts
export { type BackgroundOptions } from "@parts/background";

// Transition
export {
  buildTransition,
  type TransitionOptions,
  type TransitionType,
  type TransitionDirection,
} from "./transition";

// Animation
export {
  type AnimationType,
  type AnimationTrigger,
  type AnimationDirection,
  type AnimationClass,
  type EmphasisType,
  type PathAnimationType,
  type MediaAnimationType,
  type AnimationCalcMode,
  type AnimationValueType,
  type AnimationOptions,
} from "./animation/types";

// Constants
export { Relationships } from "@office-open/core";
export { type RelationshipType } from "@office-open/core";
export { ChartCollection, type ChartData } from "@office-open/core/chart";
export type { ChartSpaceOptions, ChartSeriesData, ChartType } from "@office-open/core/chart";
export { chartSpaceDesc } from "@office-open/core/chart";
export { SmartArtCollection, type SmartArtData } from "@office-open/core/smartart";
export { createDataModel } from "@office-open/core/smartart";
export type { TreeNode } from "@office-open/core/smartart";

// Placeholder inheritance
export {
  resolvePlaceholder,
  extractPlaceholderDefinition,
  PLACEHOLDER_TYPE_TO_KEY,
  type PlaceholderDefinition,
  type PlaceholderFacets,
  type PlaceholderPosition,
  type ResolvedPlaceholder,
} from "./placeholder";

// Slide types — re-export from parts
export type { SlideChild } from "@parts/slide/slide-child";
