/**
 * Blip image adjustment effects for DrawingML.
 *
 * These effects are applied directly to the image data within the `<a:blip>` element,
 * corresponding to Word's "Picture Format > Adjust" features (brightness, contrast,
 * grayscale, tint, duotone, etc.).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_Blip children
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";
import { createColorElement } from "../color/solid-fill";
import type { SolidFillOptions } from "../color/solid-fill";

// ─── Effect option interfaces ────────────────────────────────────────

/**
 * Options for luminance (brightness/contrast) effect.
 *
 * ## XSD: CT_LuminanceEffect
 * ```xml
 * <xsd:attribute name="bright" type="ST_FixedPercentage" default="0%"/>
 * <xsd:attribute name="contrast" type="ST_FixedPercentage" default="0%"/>
 * ```
 */
export interface LuminanceEffectOptions {
    /** Brightness adjustment in percentage points (-100 to 100). Default 0. */
    readonly bright?: number;
    /** Contrast adjustment in percentage points (-100 to 100). Default 0. */
    readonly contrast?: number;
}

/**
 * Options for HSL effect.
 *
 * ## XSD: CT_HSLEffect
 * ```xml
 * <xsd:attribute name="hue" type="ST_PositiveFixedAngle" default="0"/>
 * <xsd:attribute name="sat" type="ST_FixedPercentage" default="0%"/>
 * <xsd:attribute name="lum" type="ST_FixedPercentage" default="0%"/>
 * ```
 */
export interface HSLEffectOptions {
    /** Hue rotation in 60000ths of a degree. Default 0. */
    readonly hue?: number;
    /** Saturation adjustment in percentage points (-100 to 100). Default 0. */
    readonly sat?: number;
    /** Luminance adjustment in percentage points (-100 to 100). Default 0. */
    readonly lum?: number;
}

/**
 * Options for tint effect.
 *
 * ## XSD: CT_TintEffect
 * ```xml
 * <xsd:attribute name="hue" type="ST_PositiveFixedAngle" default="0"/>
 * <xsd:attribute name="amt" type="ST_FixedPercentage" default="0%"/>
 * ```
 */
export interface TintEffectOptions {
    /** Hue in 60000ths of a degree. Default 0. */
    readonly hue?: number;
    /** Amount in percentage points (0 to 100). Default 0. */
    readonly amt?: number;
}

/**
 * Options for duotone effect.
 *
 * ## XSD: CT_DuotoneEffect
 * ```xml
 * <xsd:group ref="EG_ColorChoice" minOccurs="2" maxOccurs="2"/>
 * ```
 */
export interface DuotoneEffectOptions {
    /** First color (dark). */
    readonly color1: SolidFillOptions;
    /** Second color (light). */
    readonly color2: SolidFillOptions;
}

/**
 * Options for biLevel (black & white threshold) effect.
 *
 * ## XSD: CT_BiLevelEffect
 * ```xml
 * <xsd:attribute name="thresh" type="ST_PositiveFixedPercentage" use="required"/>
 * ```
 */
export interface BiLevelEffectOptions {
    /** Threshold percentage (0 to 100). Pixels above are white, below are black. */
    readonly thresh: number;
}

/**
 * Options for alpha replace effect.
 *
 * ## XSD: CT_AlphaReplaceEffect
 * ```xml
 * <xsd:attribute name="a" type="ST_PositiveFixedPercentage" use="required"/>
 * ```
 */
export interface AlphaReplaceEffectOptions {
    /** Alpha percentage (0 = fully transparent, 100 = fully opaque). */
    readonly amount: number;
}

/**
 * Options for alpha bi-level effect.
 *
 * ## XSD: CT_AlphaBiLevelEffect
 * ```xml
 * <xsd:attribute name="thresh" type="ST_PositiveFixedPercentage" use="required"/>
 * ```
 */
export interface AlphaBiLevelEffectOptions {
    /** Threshold percentage (0 to 100). Alpha above is fully opaque, below is fully transparent. */
    readonly thresh: number;
}

/**
 * Options for alpha modulate fixed effect.
 *
 * ## XSD: CT_AlphaModulateFixedEffect
 * ```xml
 * <xsd:attribute name="amt" type="ST_PositivePercentage" default="100%"/>
 * ```
 */
export interface AlphaModulateFixedEffectOptions {
    /** Alpha amount in percentage (0 to 100). Default 100. */
    readonly amount?: number;
}

/**
 * Options for color change effect.
 *
 * ## XSD: CT_ColorChangeEffect
 * ```xml
 * <xsd:element name="clrFrom" type="CT_Color"/>
 * <xsd:element name="clrTo" type="CT_Color"/>
 * <xsd:attribute name="useA" type="xsd:boolean" default="true"/>
 * ```
 */
export interface ColorChangeEffectOptions {
    /** Source color to change from. */
    readonly from: SolidFillOptions;
    /** Target color to change to. */
    readonly to: SolidFillOptions;
    /** Whether to use alpha channel. Default true. */
    readonly useAlpha?: boolean;
}

/**
 * Options for color replace effect.
 *
 * ## XSD: CT_ColorReplaceEffect
 * ```xml
 * <xsd:group ref="EG_ColorChoice" minOccurs="1" maxOccurs="1"/>
 * ```
 */
export interface ColorReplaceEffectOptions {
    /** Replacement color. */
    readonly color: SolidFillOptions;
}

/**
 * Options for blur effect on blip.
 *
 * ## XSD: CT_BlurEffect
 * ```xml
 * <xsd:attribute name="rad" type="ST_PositiveCoordinate" default="0"/>
 * <xsd:attribute name="grow" type="xsd:boolean" default="true"/>
 * ```
 */
export interface BlipBlurEffectOptions {
    /** Blur radius in EMUs. Default 0. */
    readonly rad?: number;
    /** Whether to grow the blur boundary. Default true. */
    readonly grow?: boolean;
}

/**
 * All blip image adjustment effects.
 */
export interface BlipEffectsOptions {
    /** Grayscale effect (no options needed). */
    readonly grayscale?: boolean;
    /** Luminance (brightness/contrast) effect. */
    readonly luminance?: LuminanceEffectOptions;
    /** HSL adjustment effect. */
    readonly hsl?: HSLEffectOptions;
    /** Tint effect. */
    readonly tint?: TintEffectOptions;
    /** Duotone effect (two-color mapping). */
    readonly duotone?: DuotoneEffectOptions;
    /** Black & white threshold effect. */
    readonly biLevel?: BiLevelEffectOptions;
    /** Alpha ceiling effect (no options needed). */
    readonly alphaCeiling?: boolean;
    /** Alpha floor effect (no options needed). */
    readonly alphaFloor?: boolean;
    /** Alpha inverse effect (with optional color). */
    readonly alphaInverse?: SolidFillOptions;
    /** Alpha modulate fixed effect. */
    readonly alphaModFix?: AlphaModulateFixedEffectOptions;
    /** Alpha replace effect. */
    readonly alphaRepl?: AlphaReplaceEffectOptions;
    /** Alpha bi-level effect. */
    readonly alphaBiLevel?: AlphaBiLevelEffectOptions;
    /** Color change effect. */
    readonly colorChange?: ColorChangeEffectOptions;
    /** Color replace effect. */
    readonly colorRepl?: ColorReplaceEffectOptions;
    /** Blur effect on blip. */
    readonly blur?: BlipBlurEffectOptions;
}

// ─── Factory functions ───────────────────────────────────────────────

/**
 * Creates blip effect elements from BlipEffectsOptions.
 *
 * @returns Array of XML components representing blip effects
 */
export const createBlipEffects = (options: BlipEffectsOptions): XmlComponent[] => {
    const children: XmlComponent[] = [];

    if (options.grayscale) {
        children.push(new BuilderElement({ name: "a:grayscl" }));
    }

    if (options.luminance) {
        const attrs: Record<string, { readonly key: string; readonly value: string }> = {};
        if (options.luminance.bright !== undefined) {
            attrs.bright = { key: "bright", value: `${options.luminance.bright}%` };
        }
        if (options.luminance.contrast !== undefined) {
            attrs.contrast = { key: "contrast", value: `${options.luminance.contrast}%` };
        }
        children.push(new BuilderElement({ attributes: attrs as never, name: "a:lum" }));
    }

    if (options.hsl) {
        const attrs: Record<string, { readonly key: string; readonly value: string }> = {};
        if (options.hsl.hue !== undefined) {
            attrs.hue = { key: "hue", value: String(options.hsl.hue) };
        }
        if (options.hsl.sat !== undefined) {
            attrs.sat = { key: "sat", value: `${options.hsl.sat}%` };
        }
        if (options.hsl.lum !== undefined) {
            attrs.lum = { key: "lum", value: `${options.hsl.lum}%` };
        }
        children.push(new BuilderElement({ attributes: attrs as never, name: "a:hsl" }));
    }

    if (options.tint) {
        const attrs: Record<string, { readonly key: string; readonly value: string }> = {};
        if (options.tint.hue !== undefined) {
            attrs.hue = { key: "hue", value: String(options.tint.hue) };
        }
        if (options.tint.amt !== undefined) {
            attrs.amt = { key: "amt", value: `${options.tint.amt}%` };
        }
        children.push(new BuilderElement({ attributes: attrs as never, name: "a:tint" }));
    }

    if (options.duotone) {
        children.push(
            new BuilderElement({
                children: [
                    createColorElement(options.duotone.color1),
                    createColorElement(options.duotone.color2),
                ],
                name: "a:duotone",
            }),
        );
    }

    if (options.biLevel) {
        children.push(
            new BuilderElement({
                attributes: { thresh: { key: "thresh", value: `${options.biLevel.thresh}%` } },
                name: "a:biLevel",
            }),
        );
    }

    if (options.alphaCeiling) {
        children.push(new BuilderElement({ name: "a:alphaCeiling" }));
    }

    if (options.alphaFloor) {
        children.push(new BuilderElement({ name: "a:alphaFloor" }));
    }

    if (options.alphaInverse !== undefined) {
        if (typeof options.alphaInverse === "boolean") {
            children.push(new BuilderElement({ name: "a:alphaInv" }));
        } else {
            children.push(
                new BuilderElement({
                    children: [createColorElement(options.alphaInverse)],
                    name: "a:alphaInv",
                }),
            );
        }
    }

    if (options.alphaModFix) {
        const amt = options.alphaModFix.amount ?? 100;
        children.push(
            new BuilderElement({
                attributes: { amt: { key: "amt", value: `${amt}%` } },
                name: "a:alphaModFix",
            }),
        );
    }

    if (options.alphaRepl) {
        children.push(
            new BuilderElement({
                attributes: { a: { key: "a", value: `${options.alphaRepl.amount}%` } },
                name: "a:alphaRepl",
            }),
        );
    }

    if (options.alphaBiLevel) {
        children.push(
            new BuilderElement({
                attributes: { thresh: { key: "thresh", value: `${options.alphaBiLevel.thresh}%` } },
                name: "a:alphaBiLevel",
            }),
        );
    }

    if (options.colorChange) {
        const attrs: Record<string, { readonly key: string; readonly value: string }> = {};
        if (options.colorChange.useAlpha === false) {
            attrs.useA = { key: "useA", value: "0" };
        }
        children.push(
            new BuilderElement({
                attributes: attrs as never,
                children: [
                    new BuilderElement({
                        children: [createColorElement(options.colorChange.from)],
                        name: "a:clrFrom",
                    }),
                    new BuilderElement({
                        children: [createColorElement(options.colorChange.to)],
                        name: "a:clrTo",
                    }),
                ],
                name: "a:clrChange",
            }),
        );
    }

    if (options.colorRepl) {
        children.push(
            new BuilderElement({
                children: [createColorElement(options.colorRepl.color)],
                name: "a:clrRepl",
            }),
        );
    }

    if (options.blur) {
        const attrs: Record<string, { readonly key: string; readonly value: string | number }> = {};
        if (options.blur.rad !== undefined) {
            attrs.rad = { key: "rad", value: options.blur.rad };
        }
        if (options.blur.grow === false) {
            attrs.grow = { key: "grow", value: 0 };
        }
        children.push(new BuilderElement({ attributes: attrs as never, name: "a:blur" }));
    }

    return children;
};
