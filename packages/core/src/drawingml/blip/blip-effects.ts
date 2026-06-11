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
import { element } from "@office-open/xml";

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
  bright?: number;
  /** Contrast adjustment in percentage points (-100 to 100). Default 0. */
  contrast?: number;
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
  hue?: number;
  /** Saturation adjustment in percentage points (-100 to 100). Default 0. */
  saturation?: number;
  /** Luminance adjustment in percentage points (-100 to 100). Default 0. */
  luminance?: number;
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
  hue?: number;
  /** Amount in percentage points (0 to 100). Default 0. */
  amount?: number;
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
  color1: SolidFillOptions;
  /** Second color (light). */
  color2: SolidFillOptions;
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
  threshold: number;
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
  amount: number;
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
  threshold: number;
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
  amount?: number;
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
  from: SolidFillOptions;
  /** Target color to change to. */
  to: SolidFillOptions;
  /** Whether to use alpha channel. Default true. */
  useAlpha?: boolean;
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
  color: SolidFillOptions;
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
  radius?: number;
  /** Whether to grow the blur boundary. Default true. */
  grow?: boolean;
}

/**
 * All blip image adjustment effects.
 */
export interface BlipEffectsOptions {
  /** Grayscale effect (no options needed). */
  grayscale?: boolean;
  /** Luminance (brightness/contrast) effect. */
  luminance?: LuminanceEffectOptions;
  /** HSL adjustment effect. */
  hsl?: HSLEffectOptions;
  /** Tint effect. */
  tint?: TintEffectOptions;
  /** Duotone effect (two-color mapping). */
  duotone?: DuotoneEffectOptions;
  /** Black & white threshold effect. */
  biLevel?: BiLevelEffectOptions;
  /** Alpha ceiling effect (no options needed). */
  alphaCeiling?: boolean;
  /** Alpha floor effect (no options needed). */
  alphaFloor?: boolean;
  /** Alpha inverse effect (with optional color). */
  alphaInverse?: SolidFillOptions;
  /** Alpha modulate fixed effect. */
  alphaModFix?: AlphaModulateFixedEffectOptions;
  /** Alpha replace effect. */
  alphaRepl?: AlphaReplaceEffectOptions;
  /** Alpha bi-level effect. */
  alphaBiLevel?: AlphaBiLevelEffectOptions;
  /** Color change effect. */
  colorChange?: ColorChangeEffectOptions;
  /** Color replace effect. */
  colorRepl?: ColorReplaceEffectOptions;
  /** Blur effect on blip. */
  blur?: BlipBlurEffectOptions;
}

// ─── Factory functions ───────────────────────────────────────────────

/**
 * Creates blip effect elements from BlipEffectsOptions.
 *
 * @returns Array of XML strings representing blip effects
 */
export const createBlipEffects = (options: BlipEffectsOptions): string[] => {
  const children: string[] = [];

  if (options.grayscale) {
    children.push(`<a:grayscl/>`);
  }

  if (options.luminance) {
    const attrs: Record<string, string | number | undefined> = {};
    if (options.luminance.bright !== undefined) {
      attrs.bright = `${options.luminance.bright}%`;
    }
    if (options.luminance.contrast !== undefined) {
      attrs.contrast = `${options.luminance.contrast}%`;
    }
    children.push(element("a:lum", attrs));
  }

  if (options.hsl) {
    const attrs: Record<string, string | number | undefined> = {};
    if (options.hsl.hue !== undefined) {
      attrs.hue = String(options.hsl.hue);
    }
    if (options.hsl.saturation !== undefined) {
      attrs.sat = `${options.hsl.saturation}%`;
    }
    if (options.hsl.luminance !== undefined) {
      attrs.lum = `${options.hsl.luminance}%`;
    }
    children.push(element("a:hsl", attrs));
  }

  if (options.tint) {
    const attrs: Record<string, string | number | undefined> = {};
    if (options.tint.hue !== undefined) {
      attrs.hue = String(options.tint.hue);
    }
    if (options.tint.amount !== undefined) {
      attrs.amt = `${options.tint.amount}%`;
    }
    children.push(element("a:tint", attrs));
  }

  if (options.duotone) {
    children.push(
      element("a:duotone", undefined, [
        createColorElement(options.duotone.color1),
        createColorElement(options.duotone.color2),
      ]),
    );
  }

  if (options.biLevel) {
    children.push(`<a:biLevel thresh="${options.biLevel.threshold}%"/>`);
  }

  if (options.alphaCeiling) {
    children.push(`<a:alphaCeiling/>`);
  }

  if (options.alphaFloor) {
    children.push(`<a:alphaFloor/>`);
  }

  if (options.alphaInverse !== undefined) {
    if (typeof options.alphaInverse === "boolean") {
      children.push(`<a:alphaInv/>`);
    } else {
      children.push(element("a:alphaInv", undefined, [createColorElement(options.alphaInverse)]));
    }
  }

  if (options.alphaModFix) {
    const amt = options.alphaModFix.amount ?? 100;
    children.push(`<a:alphaModFix amt="${amt}%"/>`);
  }

  if (options.alphaRepl) {
    children.push(`<a:alphaRepl a="${options.alphaRepl.amount}%"/>`);
  }

  if (options.alphaBiLevel) {
    children.push(`<a:alphaBiLevel thresh="${options.alphaBiLevel.threshold}%"/>`);
  }

  if (options.colorChange) {
    const attrs: Record<string, string | number | undefined> = {};
    if (options.colorChange.useAlpha === false) {
      attrs.useA = "0";
    }
    children.push(
      element("a:clrChange", attrs, [
        element("a:clrFrom", undefined, [createColorElement(options.colorChange.from)]),
        element("a:clrTo", undefined, [createColorElement(options.colorChange.to)]),
      ]),
    );
  }

  if (options.colorRepl) {
    children.push(element("a:clrRepl", undefined, [createColorElement(options.colorRepl.color)]));
  }

  if (options.blur) {
    const attrs: Record<string, string | number | undefined> = {};
    if (options.blur.radius !== undefined) {
      attrs.rad = options.blur.radius;
    }
    if (options.blur.grow === false) {
      attrs.grow = 0;
    }
    children.push(element("a:blur", attrs));
  }

  return children;
};
