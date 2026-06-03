export {
  File,
  type PresentationOptions,
  type SlideOptions,
  type SlideCommentOptions,
  type MasterDefinition,
  type LayoutDefinition,
  type LayoutPlaceholderOptions,
  type MasterChild,
  type SlideSize,
} from "./file";
export { Presentation } from "./presentation/presentation";
export type { PresentationOptions as IPresentationXmlOptions } from "./presentation/presentation";
export { Shape, type ShapeOptions } from "./shape/shape";
export { TextBody, type TextBodyOptions } from "./shape/text-body";
export { Paragraph, type ParagraphOptions } from "./shape/paragraph/paragraph";
export { TextRun, type RunOptions } from "./shape/paragraph/run";
export { Text } from "./shape/paragraph/text";
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
export { EndParagraphRunProperties } from "./shape/paragraph/end-paragraph-run";
export { Field, SlideNumberField, DateTimeField } from "./shape/paragraph/field";

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
export { BlipFill } from "./drawingml/blip-fill";
export { Transform2D, type ITransform2DOptions } from "./drawingml/transform-2d";
export { PresetGeometry } from "./drawingml/preset-geometry";
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
export { NonVisualDrawingProperties } from "./drawingml/non-visual-drawing-props";
export { NonVisualShapeProperties } from "./drawingml/non-visual-shape-props";
export { NonVisualPictureProperties } from "./drawingml/non-visual-picture-props";
export { GroupShapeProperties } from "./drawingml/group-shape-properties";
export { GroupTransform2D, type IGroupTransform2DOptions } from "./drawingml/group-transform-2d";

export { GroupShape, type GroupShapeOptions } from "./shape/group-shape";
export { LineShape, type LineShapeOptions } from "./shape/line-shape";
export { ConnectorShape, type ConnectorShapeOptions } from "./shape/line-shape";
export { Media } from "./media/media";
export { createTransformation, type MediaTransformation } from "./media/media";
export type { IMediaData, MediaDataTransformation } from "./media/data";
export {
  VideoFrame,
  type VideoFrameOptions,
  type VideoType,
  type PosterType,
} from "./media/video-frame";
export { AudioFrame, type AudioFrameOptions, type AudioType } from "./media/audio-frame";
export { CoreProperties, type CorePropertiesOptions } from "./core-properties/properties";
export { AppProperties } from "./app-properties/app-properties";
export { ContentTypes } from "./content-types/content-types";
export { Relationships } from "./relationships/relationships";
export { type RelationshipType } from "./relationships/relationship/relationship";
export { Slide } from "./slide/slide";
export { type SlideChild } from "./slide/slide-child";
export { coerceChild } from "./slide/coerce";
export { ShapeTree } from "./shape-tree/shape-tree";
export {
  DefaultTheme,
  type ThemeOptions,
  type ColorSchemeOptions,
  type FontSchemeOptions,
} from "./theme/theme";
export {
  DefaultSlideMaster,
  type SlideMasterOptions,
  type MasterPlaceholderOptions,
  type MasterPlaceholderPosition,
} from "./slide-master/slide-master";
export { DefaultSlideLayout, SlideLayout, type SlideLayoutType } from "./slide-layout/slide-layout";
export { DefaultNotesMaster } from "./notes-master/notes-master";
export { DefaultHandoutMaster } from "./handout-master/handout-master";
export { NotesSlide, type NotesSlideOptions } from "./notes/notes-slide";
export { HeaderFooter, type SlideHeaderFooterOptions } from "./header-footer/header-footer";
export { PresentationWrapper, type ViewWrapper } from "./presentation/presentation-wrapper";
export { Picture, type PictureOptions } from "./picture/picture";
export { Background, type BackgroundOptions } from "./background/background";
export { Table, type TableOptions } from "./table/table-frame";
export { DrawingTable } from "./table/table";
export { TableRow, type TableRowOptions } from "./table/table-row";
export { TableCell, type VerticalAlignment, type TableCellOptions } from "./table/table-cell";
export { TableProperties } from "./table/table-properties";
export { TableCellProperties, type CellBorderOptions } from "./table/table-cell-properties";
export { Chart, type ChartOptions } from "./chart/chart-frame";
export { ChartCollection, type ChartData } from "./chart/chart-collection";
export { ChartSpace, type ChartSpaceOptions, type ChartSeriesData } from "./chart/chart-space";
export type { ChartType } from "./chart/chart-types/create-chart-type";
export { SmartArt, type SmartArtOptions, type TreeNode } from "./smartart";
export {
  Transition,
  type TransitionOptions,
  type TransitionType,
  type TransitionDirection,
} from "./transition/transition";
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
