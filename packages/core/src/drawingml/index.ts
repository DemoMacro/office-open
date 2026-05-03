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
    IGradientStop,
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
export { PresetGeometryAttributes } from "./geometry/preset-geometry-attributes";
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

// Media transformation
export { createTransformation } from "./media/transformation";
export type { IMediaTransformation, IMediaDataTransformation } from "./media/transformation";
