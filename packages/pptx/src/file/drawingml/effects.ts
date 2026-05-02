import { BuilderElement, XmlComponent } from "@file/xml-components";

export type EffectType = "outerShadow" | "innerShadow" | "glow" | "reflection" | "softEdge";

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
    readonly fadeStart?: number;
    readonly fadeEnd?: number;
    readonly blurRect?: string;
    readonly stRect?: string;
    readonly rotation?: number;
}

export interface ISoftEdgeOptions {
    readonly radius?: number;
}

export interface IEffectsOptions {
    readonly outerShadow?: IShadowOptions;
    readonly innerShadow?: IShadowOptions;
    readonly glow?: IGlowOptions;
    readonly reflection?: IReflectionOptions;
    readonly softEdge?: ISoftEdgeOptions;
}

function buildShadowElement(tag: string, options: IShadowOptions): XmlComponent {
    const attrs: Record<string, { readonly key: string; readonly value: string | number | boolean }> = {};
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
    if (options.blurRadius !== undefined) attrs.blurRad = { key: "blurRad", value: options.blurRadius };
    if (options.dist !== undefined) attrs.dist = { key: "dist", value: options.dist };
    if (options.direction !== undefined) attrs.dir = { key: "dir", value: options.direction };
    if (options.fadeStart !== undefined) attrs.fadeStart = { key: "fadeStart", value: options.fadeStart * 1000 };
    if (options.fadeEnd !== undefined) attrs.fadeEnd = { key: "fadeEnd", value: options.fadeEnd * 1000 };
    if (options.blurRect !== undefined) attrs.blurRect = { key: "blurRect", value: options.blurRect };
    if (options.stRect !== undefined) attrs.stRect = { key: "stRect", value: options.stRect };
    if (options.rotation !== undefined) attrs.rot = { key: "rot", value: options.rotation };

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
 */
export class EffectList extends XmlComponent {
    public constructor(options: IEffectsOptions) {
        super("a:effectLst");

        if (options.outerShadow) {
            this.root.push(buildShadowElement("outerShdw", options.outerShadow));
        }
        if (options.innerShadow) {
            this.root.push(buildShadowElement("innerShdw", options.innerShadow));
        }
        if (options.glow) {
            this.root.push(buildGlowElement(options.glow));
        }
        if (options.reflection) {
            this.root.push(buildReflectionElement(options.reflection));
        }
        if (options.softEdge) {
            this.root.push(buildSoftEdgeElement(options.softEdge));
        }
    }
}
