/**
 * Effect container (effectDag) for DrawingML shapes.
 *
 * Provides CT_EffectContainer — a directed acyclic graph (DAG) of effects
 * supporting 28 effect types including alpha/color operations, nested containers,
 * and all effects from CT_EffectList.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_EffectContainer, EG_Effect
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { createGradientFill } from "../fill/gradient-fill";
import type { GradientFillOptions } from "../fill/gradient-fill";
import { createGroupFill } from "../fill/group-fill";
import { createPatternFill } from "../fill/pattern-fill";
import type { PatternFillOptions } from "../fill/pattern-fill";
import { createNoFill } from "../outline/no-fill";
import { createSolidFill } from "../outline/solid-fill";
import type { SolidFillOptions } from "../outline/solid-fill";
import { createColorElement } from "../outline/solid-fill";
import type { BlurEffectOptions } from "./effect-list";
import type { FillOverlayEffectOptions } from "./fill-overlay";
import { createFillOverlayEffect } from "./fill-overlay";
import type { GlowEffectOptions } from "./glow";
import { createGlowEffect } from "./glow";
import type { InnerShadowEffectOptions } from "./inner-shadow";
import { createInnerShadowEffect } from "./inner-shadow";
import type { OuterShadowEffectOptions } from "./outer-shadow";
import { createOuterShadowEffect } from "./outer-shadow";
import type { PresetShadowEffectOptions } from "./preset-shadow";
import { createPresetShadowEffect } from "./preset-shadow";
import type { ReflectionEffectOptions } from "./reflection";
import { createReflectionEffect } from "./reflection";
import { createSoftEdgeEffect } from "./soft-edge";

// ─── Effect Container Type ──────────────────────────────────────────────────

/**
 * Effect container type (ST_EffectContainerType).
 *
 * Controls how child effects are combined:
 * - `sib`: sibling effects are applied in parallel
 * - `tree`: effects are applied sequentially in tree order
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_EffectContainerType">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="sib"/>
 *     <xsd:enumeration value="tree"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const EffectContainerType = {
    SIB: "sib",
    TREE: "tree",
} as const;

// ─── New Effect Options ─────────────────────────────────────────────────────

/** Alpha bi-level effect — clips alpha to threshold (CT_AlphaBiLevelEffect). */
export interface AlphaBiLevelEffectOptions {
    /** Alpha threshold (fixed percentage, required) */
    readonly thresh: number;
}

/** Alpha inverse effect — inverts alpha, optionally with color (CT_AlphaInverseEffect). */
export interface AlphaInverseEffectOptions {
    /** Optional color to apply inverse to */
    readonly color?: SolidFillOptions;
}

/** Alpha modulate fixed effect (CT_AlphaModulateFixedEffect). */
export interface AlphaModulateFixedEffectOptions {
    /** Amount percentage (default 100%) */
    readonly amt?: number;
}

/** Alpha outset effect (CT_AlphaOutsetEffect). */
export interface AlphaOutsetEffectOptions {
    /** Outset radius */
    readonly rad?: number;
}

/** Alpha replace effect — replaces alpha value (CT_AlphaReplaceEffect). */
export interface AlphaReplaceEffectOptions {
    /** New alpha value (fixed percentage, required) */
    readonly a: number;
}

/** Blend effect — blends with nested container (CT_BlendEffect). */
export interface BlendEffectOptions {
    /** Blend mode (required) */
    readonly blend: (typeof import("./fill-overlay").BlendMode)[keyof typeof import("./fill-overlay").BlendMode];
    /** Nested effect container */
    readonly container: EffectDagOptions;
}

/** Color change effect — maps one color to another (CT_ColorChangeEffect). */
export interface ColorChangeEffectOptions {
    /** Source color to replace */
    readonly from: SolidFillOptions;
    /** Target color */
    readonly to: SolidFillOptions;
    /** Whether to use alpha channel (default true) */
    readonly useA?: boolean;
}

/** Duotone effect — two-color tone mapping (CT_DuotoneEffect). */
export interface DuotoneEffectOptions {
    /** First color */
    readonly color1: SolidFillOptions;
    /** Second color */
    readonly color2: SolidFillOptions;
}

/** Fill effect — applies a fill as an effect (CT_FillEffect). */
export interface FillEffectOptions {
    /** Solid fill */
    readonly solidFill?: SolidFillOptions;
    /** Gradient fill */
    readonly gradientFill?: GradientFillOptions;
    /** Pattern fill */
    readonly patternFill?: PatternFillOptions;
    /** Group fill (inherit from parent) */
    readonly groupFill?: boolean;
    /** No fill */
    readonly noFill?: boolean;
}

/** HSL effect — adjusts hue, saturation, luminance (CT_HSLEffect). */
export interface HSLEffectOptions {
    /** Hue angle in 60,000ths of a degree (default 0) */
    readonly hue?: number;
    /** Saturation percentage (default 0%) */
    readonly sat?: number;
    /** Luminance percentage (default 0%) */
    readonly lum?: number;
}

/** Luminance effect — adjusts brightness and contrast (CT_LuminanceEffect). */
export interface LuminanceEffectOptions {
    /** Brightness percentage (default 0%) */
    readonly bright?: number;
    /** Contrast percentage (default 0%) */
    readonly contrast?: number;
}

/** Tint effect — applies tint with hue and amount (CT_TintEffect). */
export interface TintEffectOptions {
    /** Hue angle in 60,000ths of a degree (default 0) */
    readonly hue?: number;
    /** Tint amount percentage (default 0%) */
    readonly amt?: number;
}

/** Relative offset effect (CT_RelativeOffsetEffect). */
export interface RelativeOffsetEffectOptions {
    /** Horizontal offset percentage (default 0%) */
    readonly tx?: number;
    /** Vertical offset percentage (default 0%) */
    readonly ty?: number;
}

/** Transform effect — scale, skew, translate (CT_TransformEffect). */
export interface TransformEffectOptions {
    /** Horizontal scale percentage (default 100%) */
    readonly sx?: number;
    /** Vertical scale percentage (default 100%) */
    readonly sy?: number;
    /** Horizontal skew angle (default 0) */
    readonly kx?: number;
    /** Vertical skew angle (default 0) */
    readonly ky?: number;
    /** Horizontal translation */
    readonly tx?: number;
    /** Vertical translation */
    readonly ty?: number;
}

/** Effect reference — references an effect by ID (CT_EffectReference). */
export interface EffectReferenceOptions {
    /** Reference ID (required) */
    readonly ref: string;
}

// ─── Effect Container Options ───────────────────────────────────────────────

/**
 * Options for an effect container (CT_EffectContainer).
 *
 * Supports all 28 effect types from EG_Effect, including recursive nesting.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_EffectContainer">
 *   <xsd:group ref="EG_Effect" minOccurs="0" maxOccurs="unbounded"/>
 *   <xsd:attribute name="type" type="ST_EffectContainerType" default="sib"/>
 *   <xsd:attribute name="name" type="xsd:token"/>
 * </xsd:complexType>
 * ```
 */
export interface EffectDagOptions {
    /** Container type: "sib" (parallel) or "tree" (sequential) */
    readonly type?: (typeof EffectContainerType)[keyof typeof EffectContainerType];
    /** Container name */
    readonly name?: string;

    // ── Existing effects (also in CT_EffectList) ──
    readonly blur?: BlurEffectOptions;
    readonly fillOverlay?: FillOverlayEffectOptions;
    readonly glow?: GlowEffectOptions;
    readonly innerShadow?: InnerShadowEffectOptions;
    readonly outerShadow?: OuterShadowEffectOptions;
    readonly presetShadow?: PresetShadowEffectOptions;
    readonly reflection?: ReflectionEffectOptions | true;
    readonly softEdge?: number;

    // ── New effects for CT_EffectContainer ──
    /** Nested effect containers */
    readonly containers?: readonly EffectDagOptions[];
    /** Effect references */
    readonly effectRefs?: readonly EffectReferenceOptions[];
    /** Alpha bi-level effect */
    readonly alphaBiLevel?: AlphaBiLevelEffectOptions;
    /** Alpha ceiling effect (empty element) */
    readonly alphaCeiling?: boolean;
    /** Alpha floor effect (empty element) */
    readonly alphaFloor?: boolean;
    /** Alpha inverse effect */
    readonly alphaInverse?: AlphaInverseEffectOptions;
    /** Alpha modulate effect (requires nested container) */
    readonly alphaModulate?: EffectDagOptions;
    /** Alpha modulate fixed effect */
    readonly alphaModulateFixed?: AlphaModulateFixedEffectOptions;
    /** Alpha outset effect */
    readonly alphaOutset?: AlphaOutsetEffectOptions;
    /** Alpha replace effect */
    readonly alphaReplace?: AlphaReplaceEffectOptions;
    /** Bi-level effect (threshold) */
    readonly biLevel?: { readonly thresh: number };
    /** Blend effect (requires nested container) */
    readonly blend?: BlendEffectOptions;
    /** Color change effect */
    readonly colorChange?: ColorChangeEffectOptions;
    /** Color replace effect */
    readonly colorReplace?: SolidFillOptions;
    /** Duotone effect */
    readonly duotone?: DuotoneEffectOptions;
    /** Fill effect */
    readonly fill?: FillEffectOptions;
    /** Grayscale effect (empty element) */
    readonly grayscale?: boolean;
    /** HSL effect */
    readonly hsl?: HSLEffectOptions;
    /** Luminance effect */
    readonly luminance?: LuminanceEffectOptions;
    /** Tint effect */
    readonly tint?: TintEffectOptions;
    /** Relative offset effect */
    readonly relativeOffset?: RelativeOffsetEffectOptions;
    /** Transform effect */
    readonly transform?: TransformEffectOptions;
}

// ─── Internal Helpers ───────────────────────────────────────────────────────

const buildOptionalAttributes = (
    options: Record<string, unknown>,
): Record<string, { readonly key: string; readonly value: string | number }> | undefined => {
    const attrs: Record<string, { readonly key: string; readonly value: string | number }> = {};
    for (const [key, value] of Object.entries(options)) {
        if (value !== undefined) {
            attrs[key] = { key, value: value as string | number };
        }
    }
    return Object.keys(attrs).length > 0 ? attrs : undefined;
};

// ─── New Effect Factories ───────────────────────────────────────────────────

const createAlphaBiLevelEffect = (options: AlphaBiLevelEffectOptions): XmlComponent =>
    new BuilderElement<{ readonly thresh: number }>({
        name: "a:alphaBiLevel",
        attributes: { thresh: { key: "thresh", value: options.thresh } },
    });

const createAlphaCeilingEffect = (): XmlComponent => new BuilderElement({ name: "a:alphaCeiling" });

const createAlphaFloorEffect = (): XmlComponent => new BuilderElement({ name: "a:alphaFloor" });

const createAlphaInverseEffect = (options?: AlphaInverseEffectOptions): XmlComponent =>
    new BuilderElement({
        name: "a:alphaInv",
        children: options?.color ? [createColorElement(options.color)] : undefined,
    });

const createAlphaModulateFixedEffect = (
    options?: AlphaModulateFixedEffectOptions,
): XmlComponent => {
    if (options?.amt === undefined) {
        return new BuilderElement({ name: "a:alphaModFix" });
    }
    return new BuilderElement<{ readonly amt: number }>({
        name: "a:alphaModFix",
        attributes: { amt: { key: "amt", value: options.amt } },
    });
};

const createAlphaOutsetEffect = (options?: AlphaOutsetEffectOptions): XmlComponent => {
    if (options?.rad === undefined) {
        return new BuilderElement({ name: "a:alphaOutset" });
    }
    return new BuilderElement<{ readonly rad: number }>({
        name: "a:alphaOutset",
        attributes: { rad: { key: "rad", value: options.rad } },
    });
};

const createAlphaReplaceEffect = (options: AlphaReplaceEffectOptions): XmlComponent =>
    new BuilderElement<{ readonly a: number }>({
        name: "a:alphaRepl",
        attributes: { a: { key: "a", value: options.a } },
    });

const createBiLevelEffect = (thresh: number): XmlComponent =>
    new BuilderElement<{ readonly thresh: number }>({
        name: "a:biLevel",
        attributes: { thresh: { key: "thresh", value: thresh } },
    });

const createBlendEffect = (options: BlendEffectOptions): XmlComponent =>
    new BuilderElement<{ readonly blend: string }>({
        name: "a:blend",
        attributes: { blend: { key: "blend", value: options.blend } },
        children: [createEffectContainer(options.container)],
    });

const createColorChangeEffect = (options: ColorChangeEffectOptions): XmlComponent => {
    const children: XmlComponent[] = [
        new BuilderElement({
            name: "a:clrFrom",
            children: [createColorElement(options.from)],
        }),
        new BuilderElement({
            name: "a:clrTo",
            children: [createColorElement(options.to)],
        }),
    ];

    if (options.useA === false) {
        return new BuilderElement<{ readonly useA: number }>({
            name: "a:clrChange",
            attributes: { useA: { key: "useA", value: 0 } },
            children,
        });
    }

    return new BuilderElement({ name: "a:clrChange", children });
};

const createColorReplaceEffect = (color: SolidFillOptions): XmlComponent =>
    new BuilderElement({
        name: "a:clrRepl",
        children: [createColorElement(color)],
    });

const createDuotoneEffect = (options: DuotoneEffectOptions): XmlComponent =>
    new BuilderElement({
        name: "a:duotone",
        children: [createColorElement(options.color1), createColorElement(options.color2)],
    });

const createFillEffect = (options: FillEffectOptions): XmlComponent => {
    let fillElement: XmlComponent;

    if (options.noFill) {
        fillElement = createNoFill();
    } else if (options.solidFill) {
        fillElement = createSolidFill(options.solidFill);
    } else if (options.gradientFill) {
        fillElement = createGradientFill(options.gradientFill);
    } else if (options.patternFill) {
        fillElement = createPatternFill(options.patternFill);
    } else if (options.groupFill) {
        fillElement = createGroupFill();
    } else {
        fillElement = createNoFill();
    }

    return new BuilderElement({
        name: "a:fill",
        children: [fillElement],
    });
};

const createGrayscaleEffect = (): XmlComponent => new BuilderElement({ name: "a:grayscl" });

const createHSLEffect = (options?: HSLEffectOptions): XmlComponent => {
    const attrs = buildOptionalAttributes({
        hue: options?.hue,
        sat: options?.sat,
        lum: options?.lum,
    });
    return new BuilderElement({
        name: "a:hsl",
        attributes: attrs,
    });
};

const createLuminanceEffect = (options?: LuminanceEffectOptions): XmlComponent => {
    const attrs = buildOptionalAttributes({
        bright: options?.bright,
        contrast: options?.contrast,
    });
    return new BuilderElement({
        name: "a:lum",
        attributes: attrs,
    });
};

const createTintEffect = (options?: TintEffectOptions): XmlComponent => {
    const attrs = buildOptionalAttributes({
        hue: options?.hue,
        amt: options?.amt,
    });
    return new BuilderElement({
        name: "a:tint",
        attributes: attrs,
    });
};

const createRelativeOffsetEffect = (options?: RelativeOffsetEffectOptions): XmlComponent => {
    const attrs = buildOptionalAttributes({
        tx: options?.tx,
        ty: options?.ty,
    });
    return new BuilderElement({
        name: "a:relOff",
        attributes: attrs,
    });
};

const createTransformEffect = (options?: TransformEffectOptions): XmlComponent => {
    const attrs = buildOptionalAttributes({
        sx: options?.sx,
        sy: options?.sy,
        kx: options?.kx,
        ky: options?.ky,
        tx: options?.tx,
        ty: options?.ty,
    });
    return new BuilderElement({
        name: "a:xfrm",
        attributes: attrs,
    });
};

const createEffectReference = (options: EffectReferenceOptions): XmlComponent =>
    new BuilderElement<{ readonly ref: string }>({
        name: "a:effect",
        attributes: { ref: { key: "ref", value: options.ref } },
    });

// ─── Blur Effect (from effect-list, re-implemented for DAG) ─────────────────

const createBlurEffect = (options: BlurEffectOptions): XmlComponent => {
    const hasAttributes = options.rad !== undefined || options.grow === false;

    if (!hasAttributes) {
        return new BuilderElement({ name: "a:blur" });
    }

    const attributePayload = {
        ...(options.rad !== undefined && {
            rad: { key: "rad", value: options.rad },
        }),
        ...(options.grow === false && {
            grow: { key: "grow", value: 0 },
        }),
    };

    return new BuilderElement({
        attributes: attributePayload as never,
        name: "a:blur",
    });
};

// ─── Effect Container ───────────────────────────────────────────────────────

/**
 * Creates an effect container element (a:effectDag or a:cont).
 *
 * This is the CT_EffectContainer type — a recursive DAG of effects.
 * All 28 EG_Effect types are supported.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_EffectContainer">
 *   <xsd:group ref="EG_Effect" minOccurs="0" maxOccurs="unbounded"/>
 *   <xsd:attribute name="type" type="ST_EffectContainerType" default="sib"/>
 *   <xsd:attribute name="name" type="xsd:token" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @param options - Container options with effects
 * @param elementName - Element name, defaults to "a:effectDag" for top-level
 */
const createEffectContainer = (options: EffectDagOptions, elementName = "a:cont"): XmlComponent => {
    const attrs = buildOptionalAttributes({
        type: options.type,
        name: options.name,
    });

    const children: XmlComponent[] = [];

    // ── Existing effects (also in CT_EffectList) ──
    if (options.blur) {
        children.push(createBlurEffect(options.blur));
    }
    if (options.fillOverlay) {
        children.push(createFillOverlayEffect(options.fillOverlay));
    }
    if (options.glow) {
        children.push(createGlowEffect(options.glow));
    }
    if (options.innerShadow) {
        children.push(createInnerShadowEffect(options.innerShadow));
    }
    if (options.outerShadow) {
        children.push(createOuterShadowEffect(options.outerShadow));
    }
    if (options.presetShadow) {
        children.push(createPresetShadowEffect(options.presetShadow));
    }
    if (options.reflection) {
        children.push(
            createReflectionEffect(options.reflection === true ? undefined : options.reflection),
        );
    }
    if (options.softEdge !== undefined) {
        children.push(createSoftEdgeEffect(options.softEdge));
    }

    // ── New effects for CT_EffectContainer ──
    if (options.containers) {
        for (const nested of options.containers) {
            children.push(createEffectContainer(nested));
        }
    }
    if (options.effectRefs) {
        for (const ref of options.effectRefs) {
            children.push(createEffectReference(ref));
        }
    }
    if (options.alphaBiLevel) {
        children.push(createAlphaBiLevelEffect(options.alphaBiLevel));
    }
    if (options.alphaCeiling) {
        children.push(createAlphaCeilingEffect());
    }
    if (options.alphaFloor) {
        children.push(createAlphaFloorEffect());
    }
    if (options.alphaInverse) {
        children.push(createAlphaInverseEffect(options.alphaInverse));
    }
    if (options.alphaModulate) {
        children.push(
            new BuilderElement({
                name: "a:alphaMod",
                children: [createEffectContainer(options.alphaModulate)],
            }),
        );
    }
    if (options.alphaModulateFixed) {
        children.push(createAlphaModulateFixedEffect(options.alphaModulateFixed));
    }
    if (options.alphaOutset) {
        children.push(createAlphaOutsetEffect(options.alphaOutset));
    }
    if (options.alphaReplace) {
        children.push(createAlphaReplaceEffect(options.alphaReplace));
    }
    if (options.biLevel) {
        children.push(createBiLevelEffect(options.biLevel.thresh));
    }
    if (options.blend) {
        children.push(createBlendEffect(options.blend));
    }
    if (options.colorChange) {
        children.push(createColorChangeEffect(options.colorChange));
    }
    if (options.colorReplace) {
        children.push(createColorReplaceEffect(options.colorReplace));
    }
    if (options.duotone) {
        children.push(createDuotoneEffect(options.duotone));
    }
    if (options.fill) {
        children.push(createFillEffect(options.fill));
    }
    if (options.grayscale) {
        children.push(createGrayscaleEffect());
    }
    if (options.hsl) {
        children.push(createHSLEffect(options.hsl));
    }
    if (options.luminance) {
        children.push(createLuminanceEffect(options.luminance));
    }
    if (options.tint) {
        children.push(createTintEffect(options.tint));
    }
    if (options.relativeOffset) {
        children.push(createRelativeOffsetEffect(options.relativeOffset));
    }
    if (options.transform) {
        children.push(createTransformEffect(options.transform));
    }

    return new BuilderElement({
        name: elementName,
        attributes: attrs as never,
        children: children.length > 0 ? children : undefined,
    });
};

/**
 * Creates an effect DAG element (a:effectDag).
 *
 * This is the EG_EffectProperties choice alternative to effectLst,
 * supporting all 28 effect types with recursive nesting.
 *
 * @example
 * ```typescript
 * // Simple effect DAG with glow and outer shadow
 * createEffectDag({
 *   glow: { rad: 50800, color: { value: "FF0000" } },
 *   outerShadow: { color: { value: "000000" }, blurRad: 76200 },
 * });
 *
 * // DAG with alpha effects and nested container
 * createEffectDag({
 *   type: "tree",
 *   alphaBiLevel: { thresh: 50000 },
 *   containers: [{
 *     glow: { color: { value: "FF0000" } },
 *   }],
 * });
 * ```
 */
export const createEffectDag = (options: EffectDagOptions): XmlComponent =>
    createEffectContainer(options, "a:effectDag");
