// Color
export { createColorElement, createSolidFill } from "./color/solid-fill";
export type { SolidFillOptions } from "./color/solid-fill";
export { createRgbColor } from "./color/rgb-color";
export type { RgbColorOptions } from "./color/rgb-color";
export { createSchemeColor, SchemeColor } from "./color/scheme-color";
export type { SchemeColorOptions } from "./color/scheme-color";
export { createHslColor } from "./color/hsl-color";
export type { HslColorOptions } from "./color/hsl-color";
export { createSystemColor, SystemColor } from "./color/system-color";
export type { SystemColorOptions } from "./color/system-color";
export { createScRgbColor } from "./color/sc-rgb-color";
export type { ScRgbColorOptions } from "./color/sc-rgb-color";
export { createPresetColor, PresetColor } from "./color/preset-color";
export type { PresetColorOptions } from "./color/preset-color";
export { createColorTransforms } from "./color/color-transform";
export type { ColorTransformOptions } from "./color/color-transform";

// Fill
export {
  buildFill,
  extractBlipFillMedia,
  type BlipFillConfigOptions,
  type BlipFillMediaData,
  type FillOptions,
  type GradientStopOptions,
} from "./fill/fill-options";
export { createNoFill } from "./fill/no-fill";
export { createGradientFill, createGradientStop } from "./fill/gradient-fill";
export type {
  GradientFillOptions,
  GradientShadeOptions,
  LinearShadeOptions,
  PathShadeOptions,
  GradientStop,
  RelativeRect,
} from "./fill/gradient-fill";
export { PathShadeType, TileFlipMode } from "./fill/gradient-fill";
export { createPatternFill, PresetPattern } from "./fill/pattern-fill";
export type { PatternFillOptions } from "./fill/pattern-fill";
export { createGroupFill } from "./fill/group-fill";

// Outline
export { createOutline } from "./outline/outline";
export type { OutlineOptions, OutlineFillProperties } from "./outline/outline";
export { LineCap, CompoundLine, PenAlignment, PresetDash, LineJoin } from "./outline/outline";
export { createCustomDash } from "./outline/custom-dash";
export type { DashStop } from "./outline/custom-dash";
export { createLineEnd } from "./outline/line-end";
export type { LineEndOptions } from "./outline/line-end";
export { LineEndType, LineEndWidth, LineEndLength } from "./outline/line-end";

// Effects
export { createEffectList } from "./effects/effect-list";
export type { EffectListOptions, BlurEffectOptions } from "./effects/effect-list";
export { createEffectDag } from "./effects/effect-dag";
export type { EffectDagOptions, EffectContainerType } from "./effects/effect-dag";
export { createGlowEffect } from "./effects/glow";
export type { GlowEffectOptions } from "./effects/glow";
export { createOuterShadowEffect } from "./effects/outer-shadow";
export type { OuterShadowEffectOptions } from "./effects/outer-shadow";
export { createInnerShadowEffect } from "./effects/inner-shadow";
export type { InnerShadowEffectOptions } from "./effects/inner-shadow";
export { createPresetShadowEffect, PresetShadowVal } from "./effects/preset-shadow";
export type { PresetShadowEffectOptions } from "./effects/preset-shadow";
export { createReflectionEffect } from "./effects/reflection";
export type { ReflectionEffectOptions } from "./effects/reflection";
export { createSoftEdgeEffect } from "./effects/soft-edge";
export { createFillOverlayEffect, BlendMode } from "./effects/fill-overlay";
export type { FillOverlayEffectOptions } from "./effects/fill-overlay";
export { RectAlignment } from "./effects/outer-shadow";

// 3D
export { createScene3D } from "./three-d/scene-3d";
export type {
  Scene3DOptions,
  CameraOptions,
  LightRigOptions,
  BackdropOptions,
  SphereCoords,
  Point3D,
  Vector3D,
} from "./three-d/scene-3d";
export { createShape3D } from "./three-d/shape-3d";
export type { Shape3DOptions } from "./three-d/shape-3d";
export { PresetMaterialType } from "./three-d/shape-3d";
export { createBevel, createBottomBevel } from "./three-d/bevel";
export type { BevelOptions } from "./three-d/bevel";
export { BevelPresetType } from "./three-d/bevel";

// Geometry
export { PresetGeometry } from "./geometry/preset-geometry";
export type { PresetGeometryOptions } from "./geometry/preset-geometry";
export { createAdjustmentValues } from "./geometry/adjustment-values";
export type { GeometryGuide } from "./geometry/adjustment-values";
export { createCustomGeometry } from "./geometry/custom-geometry";
export type {
  CustomGeometryOptions,
  PathOptions,
  PathCommand,
  PathFillMode,
  ConnectionSite,
  GeomRect,
} from "./geometry/custom-geometry";

// Blip
export { createBlip } from "./blip/blip";
export type { BlipOptions } from "./blip/blip";
export { createBlipFill } from "./blip/blip-fill";
export type { BlipFillOptions } from "./blip/blip-fill";
export { createBlipEffects } from "./blip/blip-effects";
export type { BlipEffectsOptions } from "./blip/blip-effects";
export { createExtentionList } from "./blip/blip-extentions";
export { createSourceRectangle } from "./blip/source-rectangle";
export type { SourceRectangleOptions } from "./blip/source-rectangle";
export { Stretch } from "./blip/stretch";
export { createTileInfo } from "./blip/tile";
export type { TileOptions } from "./blip/tile";
export { TileAlignment } from "./blip/tile";

// Transform
export { createTransform2D, createGroupTransform2D } from "./transform";
export type { Transform2DOptions, GroupTransform2DOptions } from "./transform";

// Media transformation
export { createTransformation } from "./media/transformation";
export type { MediaTransformation, MediaDataTransformation } from "./media/transformation";

// Table Style
export {
  createTableStyle,
  createTableStyleList,
  type TableStyleOptions,
  type TableStyleListOptions,
  type TablePartStyleOptions,
  type TableTextStyleOptions,
  type TableCellStyleOptions,
  type TableCellBorderOptions,
  type ThemeableLineStyleOptions,
  type StyleMatrixReferenceOptions,
  type TableStyleRegion,
  type OnOffStyleType,
} from "./table-style";

// Locking
export {
  createShapeLocking,
  createPictureLocking,
  createGroupLocking,
  createGraphicFrameLocking,
  type ShapeLockingOptions,
  type PictureLockingOptions,
  type GroupLockingOptions,
  type GraphicFrameLockingOptions,
} from "./locking";

// Locked Canvas
export { LockedCanvas, type LockedCanvasOptions } from "./locked-canvas";

// Chart Drawing (cdr: elements)
export {
  createCdrX,
  createCdrY,
  createCdrFrom,
  createCdrTo,
  createCdrExt,
  createCdrCNvPr,
  createCdrCNvSpPr,
  createCdrCNvCxnSpPr,
  createCdrCNvPicPr,
  createCdrCNvGrpSpPr,
  createCdrCNvGraphicFramePr,
  createCdrNvSpPr,
  createCdrNvCxnSpPr,
  createCdrNvPicPr,
  createCdrNvGrpSpPr,
  createCdrNvGraphicFramePr,
  createCdrSp,
  createCdrCxnSp,
  createCdrPic,
  createCdrGrpSp,
  createCdrGraphicFrame,
  createCdrSpPr,
  createCdrGrpSpPr,
  createCdrStyle,
  createCdrTxBody,
  createCdrBlipFill,
  createCdrXfrm,
  createCdrRelSizeAnchor,
  createCdrAbsSizeAnchor,
} from "./chart-drawing";

// Spreadsheet Drawing (xdr: elements)
export {
  createXdrCol,
  createXdrColOff,
  createXdrRow,
  createXdrRowOff,
  createXdrFrom,
  createXdrTo,
  createXdrCNvPr,
  createXdrCNvSpPr,
  createXdrCNvCxnSpPr,
  createXdrCNvPicPr,
  createXdrCNvGrpSpPr,
  createXdrNvSpPr,
  createXdrNvCxnSpPr,
  createXdrNvPicPr,
  createXdrNvGrpSpPr,
  createXdrSp,
  createXdrCxnSp,
  createXdrPic,
  createXdrGrpSp,
  createXdrSpPr,
  createXdrGrpSpPr,
  createXdrStyle,
  createXdrTxBody,
  createXdrBlipFill,
  createXdrContentPart,
  createXdrTwoCellAnchor,
  createXdrOneCellAnchor,
} from "./spreadsheet-drawing";
export type {
  XdrMarkerOptions,
  XdrShapeAttributes,
  XdrTwoCellAnchorOptions,
} from "./spreadsheet-drawing";

// Wordprocessing Drawing (wp: elements)
export {
  createWpCNvPr,
  createWpCNvSpPr,
  createWpCNvCnPr,
  createWpCNvFrPr,
  createWpCNvGrpSpPr,
  createWpCNvContentPartPr,
  createWpSp,
  createWpSpPr,
  createWpGrpSpPr,
  createWpStyle,
  createWpXfrm,
  createWpBodyPr,
  createWpTxbx,
  createWpTxbxContent,
  createWpLinkedTxbx,
  createWpExtLst,
  createWpGraphicFrame,
  createWpContentPart,
  createWpNvContentPartPr,
  createWpGrpSp,
  createWpNestedGrpSp,
  createWpCanvas,
  createWpBg,
  createWpWhole,
} from "./wordprocessing-drawing";
export type {
  WpNonVisualDrawingPropsOptions,
  WpShapeOptions,
  WpTextboxOptions,
  WpLinkedTextboxOptions,
  WpGraphicFrameOptions,
  WpContentPartOptions,
  WpGroupOptions,
  WpCanvasOptions,
} from "./wordprocessing-drawing";

// Diagram (SmartArt dgm: elements)
export {
  // Layout variable property set
  createAdj,
  createAdjLst,
  createAnimLvl,
  createAnimOne,
  createChMax,
  createChPref,
  createOrgChart,
  createHierBranch,
  createPresLayoutVars,
  AnimLevelValue,
  AnimOneValue,
  HierBranchStyle,
  // Definition headers
  createColorsDefHdr,
  createColorsDefHdrLst,
  createLayoutDefHdr,
  createLayoutDefHdrLst,
  createStyleDefHdr,
  createStyleDefHdrLst,
  // Style and color lists
  createDiagramStyle,
  createFillClrLst,
  createLinClrLst,
  createEffectClrLst,
  createTxFillClrLst,
  createTxLinClrLst,
  createTxEffectClrLst,
  createStyleLbl,
  ColorMethod,
  HueDirection,
  StyleMatrixIndex,
  FontCollectionIndex,
  // Relationship IDs
  createDiagramRelIds,
  // Extension list, sp3d, txPr
  createDiagramExtLst,
  createDiagramSp3d,
  createDiagramTxPr,
} from "./diagram";
export type {
  AdjOptions,
  AdjLstOptions,
  AnimLvlOptions,
  AnimOneOptions,
  ChMaxOptions,
  ChPrefOptions,
  OrgChartOptions,
  HierBranchOptions,
  PresLayoutVarsOptions,
  DiagramNameOptions,
  DiagramDescriptionOptions,
  DiagramCategoryOptions,
  ColorsDefHdrOptions,
  ColorsDefHdrLstOptions,
  LayoutDefHdrOptions,
  LayoutDefHdrLstOptions,
  StyleDefHdrOptions,
  StyleDefHdrLstOptions,
  DiagramStyleOptions,
  ColorListOptions,
  DiagramStyleLblOptions,
  DiagramRelIdsOptions,
  DiagramExtLstOptions,
  DiagramExtensionOptions,
  DiagramTextPropsOptions,
} from "./diagram";
