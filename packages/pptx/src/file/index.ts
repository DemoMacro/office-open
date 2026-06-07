export {
  File,
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
export type { ViewPropertiesOptions } from "./view-properties";
export type {
  PresentationOptions as IPresentationXmlOptions,
  PhotoAlbumOptions,
  ModifyVerifierOptions,
  EmbeddedFontOptions,
  CustomShowOptions,
  KinsokuOptions,
  CustomerDataOptions,
  ViewWrapper,
} from "./presentation";
export { Shape, type ShapeOptions } from "./shape/shape";
export { TextBody, type TextBodyOptions } from "./shape/text-body";
export { Paragraph, type ParagraphOptions } from "./shape/paragraph/paragraph";
export { TextRun, type RunOptions } from "./shape/paragraph/run";
export {
  RunProperties,
  UnderlineStyle,
  StrikeStyle,
  TextCapitalization,
  type RunPropertiesOptions,
  type HyperlinkOptions,
} from "./shape/paragraph/run-properties";
export {
  ParagraphProperties,
  type TextAlignment,
  type ParagraphPropertiesOptions,
} from "./shape/paragraph/paragraph-properties";

// DrawingML — re-exports from core
export {
  createOutline,
  type OutlineOptions as CoreOutlineOptions,
} from "@office-open/core/drawingml";
export {
  createGradientFill,
  createGradientStop,
  type GradientFillOptions as CoreGradientFillOptions,
  type GradientStop,
  PathShadeType,
  TileFlipMode,
} from "@office-open/core/drawingml";
export {
  LineCap,
  CompoundLine,
  PenAlignment,
  PresetDash,
  LineJoin,
} from "@office-open/core/drawingml";
export { createScene3D, type Scene3DOptions } from "@office-open/core/drawingml";
export {
  createShape3D,
  type Shape3DOptions,
  PresetMaterialType,
} from "@office-open/core/drawingml";
export {
  createBevel,
  createBottomBevel,
  type BevelOptions,
  BevelPresetType,
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
export { createOutlineCompat, type OutlineOptions } from "./drawingml/outline";

// DrawingML — local implementations (pptx-specific)
export { Transform2D, type ITransform2DOptions } from "./drawingml/transform-2d";
export { PresetGeometry } from "@office-open/core/drawingml";
export { ShapeProperties, type ShapePropertiesOptions } from "./drawingml/shape-properties";
export {
  createPptxEffectList,
  ReflectionAlignment,
  type EffectsOptions,
  type ShadowOptions,
  type GlowOptions,
  type ReflectionOptions,
  type SoftEdgeOptions,
} from "./drawingml/effects";

// Shape types (classes stripped, interfaces only)
export type { GroupShapeOptions } from "./shape/group-shape";
export type { LineShapeOptions, ConnectorShapeOptions } from "./shape/line-shape";
export { Media } from "./media/media";
export { createTransformation, type MediaTransformation } from "./media/media";
export type { IMediaData, MediaDataTransformation } from "./media/data";
export type { VideoFrameOptions, VideoType, PosterType } from "./media/video-frame";
export type { AudioFrameOptions, AudioType } from "./media/audio-frame";
export type { CorePropertiesOptions } from "./core-properties";
export { AppProperties } from "@office-open/core";
export { ContentTypes } from "./content-types";
export { Relationships } from "@office-open/core";
export { type RelationshipType } from "@office-open/core";
export { type SlideChild } from "./slide/slide-child";
export {
  DefaultTheme,
  type ThemeOptions,
  type ColorSchemeOptions,
  type FontSchemeOptions,
} from "./theme";
export {
  DefaultSlideMaster,
  type SlideMasterOptions,
  type MasterPlaceholderOptions,
  type MasterPlaceholderPosition,
} from "./slide-master";
export { DefaultSlideLayout, SlideLayout, type SlideLayoutType } from "./slide-layout";
export { DefaultNotesMaster } from "./notes-master";
export { DefaultHandoutMaster } from "./handout-master";
export { buildNotesSlideXml, type NotesSlideOptions } from "./notes-slide";
export type { SlideHeaderFooterOptions } from "./header-footer";
export type { PictureOptions } from "./picture";
export { Background, type BackgroundOptions } from "./background";
export type { TableOptions } from "./table/table-frame";
export type { TableRowOptions } from "./table/table-row";
export type { VerticalAlignment, TableCellOptions } from "./table/table-cell";
export type { CellBorderOptions } from "./table/table-cell-properties";
export type { ChartOptions } from "./chart-frame";
export { ChartCollection, type ChartData } from "@office-open/core/chart";
export { ChartSpace, type ChartSpaceOptions, type ChartSeriesData } from "@office-open/core/chart";
export type { ChartType } from "@office-open/core/chart";
export { SmartArtCollection, type SmartArtData } from "@office-open/core/smartart";
export { createDataModel } from "@office-open/core/smartart";
export type { SmartArtOptions } from "./smartart";
export type { TreeNode } from "@office-open/core/smartart";
export type { LockedCanvasFrameOptions } from "./locked-canvas-frame";
export {
  buildTransition,
  type TransitionOptions,
  type TransitionType,
  type TransitionDirection,
} from "./transition";
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
