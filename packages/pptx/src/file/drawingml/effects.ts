import {
    createEffectList,
    createScene3D,
    createShape3D,
    createBevel,
    createBottomBevel,
    type EffectListOptions,
    type Scene3DOptions,
    type Shape3DOptions,
    type BevelOptions,
} from "@office-open/core/drawingml";

export type EffectType = "outerShadow" | "innerShadow" | "glow" | "reflection" | "softEdge";

export const ReflectionAlignment = {
    TOP_LEFT: "tl",
    TOP: "t",
    TOP_RIGHT: "tr",
    LEFT: "l",
    CENTER: "ctr",
    RIGHT: "r",
    BOTTOM_LEFT: "bl",
    BOTTOM: "b",
    BOTTOM_RIGHT: "br",
} as const;

export interface IShadowOptions {
    readonly blur?: number;
    readonly dist?: number;
    readonly direction?: number;
    readonly color?: string;
    readonly alpha?: number;
    readonly rotateWithShape?: boolean;
}

export interface IGlowOptions {
    readonly radius?: number;
    readonly color?: string;
    readonly alpha?: number;
}

export interface IReflectionOptions {
    readonly blurRadius?: number;
    readonly dist?: number;
    readonly direction?: number;
    readonly startAlpha?: number;
    readonly startPosition?: number;
    readonly endAlpha?: number;
    readonly endPosition?: number;
    readonly fadeDirection?: number;
    readonly scaleX?: number;
    readonly scaleY?: number;
    readonly skewX?: number;
    readonly skewY?: number;
    readonly alignment?: keyof typeof ReflectionAlignment;
    readonly rotateWithShape?: boolean;
}

export interface ISoftEdgeOptions {
    readonly radius?: number;
}

export interface IBevelOptions {
    readonly width?: number;
    readonly height?: number;
}

export interface IRotation3DOptions {
    readonly x?: number;
    readonly y?: number;
    readonly z?: number;
    readonly perspective?: number;
}

export interface IEffectsOptions {
    readonly outerShadow?: IShadowOptions;
    readonly innerShadow?: IShadowOptions;
    readonly glow?: IGlowOptions;
    readonly reflection?: IReflectionOptions;
    readonly softEdge?: ISoftEdgeOptions;
    readonly rotation3D?: IRotation3DOptions;
    readonly bevelTop?: IBevelOptions;
    readonly bevelBottom?: IBevelOptions;
    readonly extrusionH?: number;
    readonly material?:
        | "plastic"
        | "metal"
        | "matte"
        | "warmMatte"
        | "softEdge"
        | "flat"
        | "powder";
    readonly lighting?:
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

/** Convert PPTX IShadowOptions to core OuterShadowEffectOptions. */
function toOuterShadow(opts: IShadowOptions) {
    return {
        blurRad: opts.blur,
        dist: opts.dist,
        dir: opts.direction,
        rotWithShape: opts.rotateWithShape === false ? false : undefined,
        color: toColor(opts.color, opts.alpha),
    };
}

/** Convert PPTX IShadowOptions to core InnerShadowEffectOptions. */
function toInnerShadow(opts: IShadowOptions) {
    return {
        blurRad: opts.blur,
        dist: opts.dist,
        dir: opts.direction,
        color: toColor(opts.color, opts.alpha),
    };
}

/** Convert PPTX IGlowOptions to core GlowEffectOptions. */
function toGlow(opts: IGlowOptions) {
    return {
        rad: opts.radius ?? 152400,
        color: toColor(opts.color, opts.alpha),
    };
}

/** Convert PPTX IReflectionOptions to core ReflectionEffectOptions. */
function toReflection(opts: IReflectionOptions) {
    const result: Record<string, number> = {};
    if (opts.blurRadius !== undefined) result.blurRad = opts.blurRadius;
    if (opts.dist !== undefined) result.dist = opts.dist;
    if (opts.direction !== undefined) result.dir = opts.direction;
    if (opts.startAlpha !== undefined) result.stA = opts.startAlpha * 1000;
    if (opts.startPosition !== undefined) result.stPos = opts.startPosition * 1000;
    if (opts.endAlpha !== undefined) result.endA = opts.endAlpha * 1000;
    if (opts.endPosition !== undefined) result.endPos = opts.endPosition * 1000;
    if (opts.fadeDirection !== undefined) result.fadeDir = opts.fadeDirection * 60000;
    if (opts.scaleX !== undefined) result.sx = opts.scaleX * 1000;
    if (opts.scaleY !== undefined) result.sy = opts.scaleY * 1000;
    if (opts.skewX !== undefined) result.kx = opts.skewX * 60000;
    if (opts.skewY !== undefined) result.ky = opts.skewY * 60000;
    if (opts.alignment !== undefined) result.algn = ReflectionAlignment[opts.alignment];
    if (opts.rotateWithShape === false) result.rotWithShape = 0;
    return result;
}

/** Convert PPTX IBevelOptions to core BevelOptions. */
function toBevel(opts: IBevelOptions): BevelOptions {
    return {
        ...(opts.width !== undefined && { w: opts.width * 12700 }),
        ...(opts.height !== undefined && { h: opts.height * 12700 }),
    };
}

/** Map PPTX IEffectsOptions to core EffectListOptions. */
function toEffectListOptions(opts: IEffectsOptions): EffectListOptions | undefined {
    const hasEffects = opts.outerShadow || opts.innerShadow || opts.glow || opts.reflection || opts.softEdge;
    if (!hasEffects) return undefined;

    return {
        outerShadow: opts.outerShadow ? toOuterShadow(opts.outerShadow) : undefined,
        innerShadow: opts.innerShadow ? toInnerShadow(opts.innerShadow) : undefined,
        glow: opts.glow ? toGlow(opts.glow) : undefined,
        reflection: opts.reflection ? toReflection(opts.reflection) : undefined,
        softEdge: opts.softEdge ? opts.softEdge.radius ?? 50800 : undefined,
    };
}

/** Map PPTX IEffectsOptions to core Scene3DOptions, or null if not needed. */
export function buildScene3D(options: IEffectsOptions): ReturnType<typeof createScene3D> | null {
    if (!options.rotation3D && !options.lighting) return null;

    const cameraPreset = options.rotation3D?.perspective ? "legacyPerspectiveFront" : "orthographicFront";
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

/** Map PPTX IEffectsOptions to core Shape3DOptions, or null if not needed. */
export function buildShape3D(options: IEffectsOptions): ReturnType<typeof createShape3D> | null {
    if (!options.extrusionH && !options.bevelTop && !options.bevelBottom && !options.material) return null;

    const shape3dOpts: Shape3DOptions = {};
    if (options.bevelTop) shape3dOpts.bevelT = toBevel(options.bevelTop);
    if (options.bevelBottom) shape3dOpts.bevelB = toBevel(options.bevelBottom);
    if (options.extrusionH !== undefined) shape3dOpts.extrusionH = options.extrusionH;
    if (options.material) shape3dOpts.prstMaterial = options.material;

    return createShape3D(shape3dOpts);
}

/** Create a:effectLst from PPTX simplified IEffectsOptions. */
export function createPptxEffectList(options: IEffectsOptions) {
    const effectListOpts = toEffectListOptions(options);
    return effectListOpts ? createEffectList(effectListOpts) : null;
}

// Re-export core types for advanced users
export { EffectListOptions, Scene3DOptions, Shape3DOptions, BevelOptions };
