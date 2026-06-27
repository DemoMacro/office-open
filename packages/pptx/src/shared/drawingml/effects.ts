import { convertPointsToEmu, xsdRectAlignment } from "@office-open/core";
import {
  createEffectList,
  createScene3D,
  createShape3D,
  type EffectListOptions,
  type Scene3DOptions,
  type Shape3DOptions,
  type BevelOptions,
} from "@office-open/core/drawingml";

export type EffectType = "outerShadow" | "innerShadow" | "glow" | "reflection" | "softEdge";

export const ReflectionAlignment = {
  TOP_LEFT: "topLeft",
  TOP: "top",
  TOP_RIGHT: "topRight",
  LEFT: "left",
  CENTER: "center",
  RIGHT: "right",
  BOTTOM_LEFT: "bottomLeft",
  BOTTOM: "bottom",
  BOTTOM_RIGHT: "bottomRight",
} as const;

export interface ShadowOptions {
  blur?: number;
  distance?: number;
  direction?: number;
  color?: string;
  alpha?: number;
  rotateWithShape?: boolean;
}

export interface GlowOptions {
  radius?: number;
  color?: string;
  alpha?: number;
}

export interface ReflectionOptions {
  blurRadius?: number;
  distance?: number;
  direction?: number;
  startAlpha?: number;
  startPosition?: number;
  endAlpha?: number;
  endPosition?: number;
  fadeDirection?: number;
  scaleX?: number;
  scaleY?: number;
  skewX?: number;
  skewY?: number;
  alignment?: (typeof ReflectionAlignment)[keyof typeof ReflectionAlignment];
  rotateWithShape?: boolean;
}

export interface SoftEdgeOptions {
  radius?: number;
}

export interface PPTXBevelOptions {
  width?: number;
  height?: number;
  preset?: string;
}

export interface Rotation3DOptions {
  x?: number;
  y?: number;
  z?: number;
  perspective?: number;
}

export interface EffectsOptions {
  outerShadow?: ShadowOptions;
  innerShadow?: ShadowOptions;
  glow?: GlowOptions;
  reflection?: ReflectionOptions;
  softEdge?: SoftEdgeOptions;
  rotation3D?: Rotation3DOptions;
  bevelTop?: PPTXBevelOptions;
  bevelBottom?: PPTXBevelOptions;
  extrusionH?: number;
  /** Shape depth (a:sp3d/@z) in EMUs. */
  depth?: number;
  /** Contour width (a:sp3d/@contourW) in EMUs. */
  contourWidth?: number;
  material?: "plastic" | "metal" | "matte" | "warmMatte" | "softEdge" | "flat" | "powder";
  lighting?:
    | "flat"
    | "legacyFlat1"
    | "legacyFlat2"
    | "legacyFlat3"
    | "legacyFlat4"
    | "legacyHarsh1"
    | "legacyHarsh2"
    | "legacyHarsh3"
    | "legacyHarsh4"
    | "legacyNormal1"
    | "legacyNormal2"
    | "legacyNormal3"
    | "legacyNormal4"
    | "threePt"
    | "balanced"
    | "soft"
    | "harsh"
    | "flood"
    | "contrasting"
    | "morning"
    | "sunrise"
    | "sunset"
    | "chilly"
    | "freezing"
    | "twoPt"
    | "brightRoom"
    | "gallery";
}

/** Convert PPTX simple color to core SolidFillOptions. */
function toColor(color?: string, alpha?: number) {
  if (!color) return { value: "000000", alpha: (alpha ?? 40) * 1000 };
  return { value: color.replace("#", ""), alpha: (alpha ?? 40) * 1000 };
}

/** Convert PPTX ShadowOptions to core OuterShadowEffectOptions. */
function toOuterShadow(opts: ShadowOptions) {
  return {
    blurRadius: opts.blur,
    distance: opts.distance,
    direction: opts.direction,
    rotWithShape: opts.rotateWithShape === false ? false : undefined,
    color: toColor(opts.color, opts.alpha),
  };
}

/** Convert PPTX ShadowOptions to core InnerShadowEffectOptions. */
function toInnerShadow(opts: ShadowOptions) {
  return {
    blurRadius: opts.blur,
    distance: opts.distance,
    direction: opts.direction,
    color: toColor(opts.color, opts.alpha),
  };
}

/** Convert PPTX GlowOptions to core GlowEffectOptions. */
function toGlow(opts: GlowOptions) {
  return {
    radius: opts.radius ?? 152400,
    color: toColor(opts.color, opts.alpha),
  };
}

/** Convert PPTX ReflectionOptions to core ReflectionEffectOptions. */
function toReflection(opts: ReflectionOptions) {
  const result: Record<string, number | string> = {};
  if (opts.blurRadius !== undefined) result.blurRadius = opts.blurRadius;
  if (opts.distance !== undefined) result.distance = opts.distance;
  if (opts.direction !== undefined) result.direction = opts.direction;
  if (opts.startAlpha !== undefined) result.startAlpha = opts.startAlpha * 1000;
  if (opts.startPosition !== undefined) result.startPosition = opts.startPosition * 1000;
  if (opts.endAlpha !== undefined) result.endAlpha = opts.endAlpha * 1000;
  if (opts.endPosition !== undefined) result.endPosition = opts.endPosition * 1000;
  if (opts.fadeDirection !== undefined) result.fadeDirection = opts.fadeDirection * 60000;
  if (opts.scaleX !== undefined) result.scaleX = opts.scaleX * 1000;
  if (opts.scaleY !== undefined) result.scaleY = opts.scaleY * 1000;
  if (opts.skewX !== undefined) result.skewX = opts.skewX * 60000;
  if (opts.skewY !== undefined) result.skewY = opts.skewY * 60000;
  if (opts.alignment !== undefined) result.alignment = xsdRectAlignment.to(opts.alignment);
  if (opts.rotateWithShape === false) result.rotWithShape = 0;
  return result;
}

/** Convert PPTX PPTXBevelOptions to core BevelOptions. */
function toBevel(opts: PPTXBevelOptions): BevelOptions {
  return {
    ...(opts.width !== undefined && { w: convertPointsToEmu(opts.width) }),
    ...(opts.height !== undefined && { h: convertPointsToEmu(opts.height) }),
    ...(opts.preset !== undefined && { prst: opts.preset as BevelOptions["prst"] }),
  };
}

/** Map PPTX EffectsOptions to core EffectListOptions. */
export function toEffectListOptions(opts: EffectsOptions): EffectListOptions | undefined {
  const hasEffects =
    opts.outerShadow || opts.innerShadow || opts.glow || opts.reflection || opts.softEdge;
  if (!hasEffects) return undefined;

  return {
    outerShadow: opts.outerShadow ? toOuterShadow(opts.outerShadow) : undefined,
    innerShadow: opts.innerShadow ? toInnerShadow(opts.innerShadow) : undefined,
    glow: opts.glow ? toGlow(opts.glow) : undefined,
    reflection: opts.reflection ? toReflection(opts.reflection) : undefined,
    softEdge: opts.softEdge ? (opts.softEdge.radius ?? 50800) : undefined,
  };
}

/** Map PPTX EffectsOptions to core Scene3DOptions, or null if not needed. */
export function buildScene3D(options: EffectsOptions): ReturnType<typeof createScene3D> | null {
  if (!options.rotation3D && !options.lighting) return null;

  const cameraPreset = options.rotation3D?.perspective
    ? "legacyPerspectiveFront"
    : "orthographicFront";
  const cameraOpts = {
    preset: cameraPreset,
    ...(options.rotation3D?.perspective && { fov: options.rotation3D.perspective }),
    ...(options.rotation3D && {
      rotation: {
        lat: (options.rotation3D.x ?? 0) * 60000,
        lon: (options.rotation3D.y ?? 0) * 60000,
        rev: (options.rotation3D.z ?? 0) * 60000,
      },
    }),
  };

  return createScene3D({
    camera: cameraOpts as Scene3DOptions["camera"],
    lightRig: { rig: options.lighting ?? "threePt", direction: "t" },
  });
}

/** Map PPTX EffectsOptions to core Shape3DOptions, or null if not needed. */
export function buildShape3D(options: EffectsOptions): ReturnType<typeof createShape3D> | null {
  if (
    !options.extrusionH &&
    !options.bevelTop &&
    !options.bevelBottom &&
    !options.material &&
    !options.depth &&
    !options.contourWidth
  )
    return null;

  const shape3dOpts: Shape3DOptions = {
    ...(options.bevelTop ? { bevelT: toBevel(options.bevelTop) } : {}),
    ...(options.bevelBottom ? { bevelB: toBevel(options.bevelBottom) } : {}),
    ...(options.extrusionH !== undefined ? { extrusionH: options.extrusionH } : {}),
    ...(options.material ? { prstMaterial: options.material } : {}),
    ...(options.depth !== undefined ? { z: options.depth } : {}),
    ...(options.contourWidth !== undefined ? { contourW: options.contourWidth } : {}),
  };

  return createShape3D(shape3dOpts);
}

/** Create a:effectLst from PPTX simplified EffectsOptions. */
export function createPptxEffectList(options: EffectsOptions) {
  const effectListOpts = toEffectListOptions(options);
  return effectListOpts ? createEffectList(effectListOpts) : null;
}

// Re-export core types for advanced users
export type { EffectListOptions, Scene3DOptions, Shape3DOptions, BevelOptions };
