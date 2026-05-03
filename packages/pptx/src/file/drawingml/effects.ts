import { BuilderElement, XmlComponent } from "@file/xml-components";

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

function buildShadowElement(tag: string, options: IShadowOptions): XmlComponent {
    const attrs: Record<
        string,
        { readonly key: string; readonly value: string | number | boolean }
    > = {};
    if (options.blur !== undefined) attrs.blurRad = { key: "blurRad", value: options.blur };
    if (options.dist !== undefined) attrs.dist = { key: "dist", value: options.dist };
    if (options.direction !== undefined) attrs.dir = { key: "dir", value: options.direction };
    if (options.rotateWithShape !== undefined)
        attrs.rotWithShape = { key: "rotWithShape", value: options.rotateWithShape ? 1 : 0 };

    const color = options.color ?? "000000";
    const alpha = options.alpha ?? 40;

    return new BuilderElement({
        name: `a:${tag}`,
        attributes: attrs,
        children: [
            new BuilderElement({
                name: "a:srgbClr",
                attributes: { val: { key: "val", value: color } },
                children: [
                    new BuilderElement({
                        name: "a:alpha",
                        attributes: { val: { key: "val", value: alpha * 1000 } },
                    }),
                ],
            }),
        ],
    });
}

function buildGlowElement(options: IGlowOptions): XmlComponent {
    const radius = options.radius ?? 152400;
    const color = options.color ?? "4472C4";
    const alpha = options.alpha ?? 40;

    return new BuilderElement({
        name: "a:glow",
        attributes: { rad: { key: "rad", value: radius } },
        children: [
            new BuilderElement({
                name: "a:srgbClr",
                attributes: { val: { key: "val", value: color } },
                children: [
                    new BuilderElement({
                        name: "a:alpha",
                        attributes: { val: { key: "val", value: alpha * 1000 } },
                    }),
                ],
            }),
        ],
    });
}

function buildReflectionElement(options: IReflectionOptions): XmlComponent {
    const attrs: Record<string, { readonly key: string; readonly value: string | number }> = {};
    if (options.blurRadius !== undefined)
        attrs.blurRad = { key: "blurRad", value: options.blurRadius };
    if (options.startAlpha !== undefined)
        attrs.stA = { key: "stA", value: options.startAlpha * 1000 };
    if (options.startPosition !== undefined)
        attrs.stPos = { key: "stPos", value: options.startPosition * 1000 };
    if (options.endAlpha !== undefined)
        attrs.endA = { key: "endA", value: options.endAlpha * 1000 };
    if (options.endPosition !== undefined)
        attrs.endPos = { key: "endPos", value: options.endPosition * 1000 };
    if (options.dist !== undefined) attrs.dist = { key: "dist", value: options.dist };
    if (options.direction !== undefined) attrs.dir = { key: "dir", value: options.direction };
    if (options.fadeDirection !== undefined)
        attrs.fadeDir = { key: "fadeDir", value: options.fadeDirection * 60000 };
    if (options.scaleX !== undefined) attrs.sx = { key: "sx", value: options.scaleX * 1000 };
    if (options.scaleY !== undefined) attrs.sy = { key: "sy", value: options.scaleY * 1000 };
    if (options.skewX !== undefined) attrs.kx = { key: "kx", value: options.skewX * 60000 };
    if (options.skewY !== undefined) attrs.ky = { key: "ky", value: options.skewY * 60000 };
    if (options.alignment !== undefined)
        attrs.algn = { key: "algn", value: ReflectionAlignment[options.alignment] };
    if (options.rotateWithShape !== undefined)
        attrs.rotWithShape = { key: "rotWithShape", value: options.rotateWithShape ? 1 : 0 };

    return new BuilderElement({ name: "a:reflection", attributes: attrs });
}

function buildSoftEdgeElement(options: ISoftEdgeOptions): XmlComponent {
    const radius = options.radius ?? 50800;
    return new BuilderElement({
        name: "a:softEdge",
        attributes: { rad: { key: "rad", value: radius } },
    });
}

/**
 * a:effectLst — Visual effects for a shape (shadow, glow, reflection, etc.).
 * XSD order: blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge
 */
export class EffectList extends XmlComponent {
    public constructor(options: IEffectsOptions) {
        super("a:effectLst");

        if (options.glow) {
            this.root.push(buildGlowElement(options.glow));
        }
        if (options.innerShadow) {
            this.root.push(buildShadowElement("innerShdw", options.innerShadow));
        }
        if (options.outerShadow) {
            this.root.push(buildShadowElement("outerShdw", options.outerShadow));
        }
        if (options.reflection) {
            this.root.push(buildReflectionElement(options.reflection));
        }
        if (options.softEdge) {
            this.root.push(buildSoftEdgeElement(options.softEdge));
        }
    }
}

/**
 * a:scene3d — 3D scene for a shape (camera + lighting rig).
 */
export function buildScene3D(options: IEffectsOptions): XmlComponent | null {
    if (!options.rotation3D && !options.lighting) return null;

    const rotAttrs: Record<string, { readonly key: string; readonly value: number }> = {
        lat: { key: "lat", value: (options.rotation3D?.x ?? 0) * 60000 },
        lon: { key: "lon", value: (options.rotation3D?.y ?? 0) * 60000 },
        rev: { key: "rev", value: (options.rotation3D?.z ?? 0) * 60000 },
    };

    const cameraAttrs: Record<string, { readonly key: string; readonly value: string | number }> = {
        prst: {
            key: "prst",
            value: options.rotation3D?.perspective ? "legacyPerspectiveFront" : "orthographicFront",
        },
    };
    if (options.rotation3D?.perspective) {
        cameraAttrs.fov = { key: "fov", value: options.rotation3D.perspective };
    }

    const camera = new BuilderElement({
        name: "a:camera",
        attributes: cameraAttrs,
        children: options.rotation3D
            ? [new BuilderElement({ name: "a:rot", attributes: rotAttrs })]
            : undefined,
    });

    const lightRig = new BuilderElement({
        name: "a:lightRig",
        attributes: {
            rig: { key: "rig", value: options.lighting ?? "threePt" },
            dir: { key: "dir", value: "t" },
        },
        children: [
            new BuilderElement({
                name: "a:rot",
                attributes: {
                    lat: { key: "lat", value: 0 },
                    lon: { key: "lon", value: 0 },
                    rev: { key: "rev", value: 0 },
                },
            }),
        ],
    });

    return new BuilderElement({
        name: "a:scene3d",
        children: [camera, lightRig],
    });
}

/**
 * a:sp3d — 3D shape properties (extrusion, bevel, material).
 * XSD: extrusionH/prstMaterial/contourW/z are attributes; children are bevelT → bevelB → extrusionClr → contourClr → extLst
 */
export function buildShape3D(options: IEffectsOptions): XmlComponent | null {
    if (!options.extrusionH && !options.bevelTop && !options.bevelBottom && !options.material)
        return null;

    const children: BuilderElement<{}>[] = [];

    if (options.bevelTop) {
        children.push(buildBevelElement("a:bevelT", options.bevelTop));
    }
    if (options.bevelBottom) {
        children.push(buildBevelElement("a:bevelB", options.bevelBottom));
    }

    const attrs: Record<string, { readonly key: string; readonly value: string | number }> = {};
    if (options.extrusionH !== undefined)
        attrs.extrusionH = { key: "extrusionH", value: options.extrusionH };
    if (options.material) attrs.prstMaterial = { key: "prstMaterial", value: options.material };

    return new BuilderElement({
        name: "a:sp3d",
        attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
        children: children.length > 0 ? children : undefined,
    });
}

function buildBevelElement(name: string, options: IBevelOptions): BuilderElement<{}> {
    const attrs: Record<string, { readonly key: string; readonly value: number }> = {};
    if (options.width !== undefined) attrs.w = { key: "w", value: options.width * 12700 };
    if (options.height !== undefined) attrs.h = { key: "h", value: options.height * 12700 };
    return new BuilderElement({
        name,
        attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
    });
}
