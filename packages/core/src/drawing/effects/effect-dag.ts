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
import { element } from "@office-open/xml";

import { xsdBlendMode, xsdEffectContainer } from "../../util/mappings";
import { createSolidFill } from "../color/solid-fill";
import type { SolidFillOptions } from "../color/solid-fill";
import { createColorElement } from "../color/solid-fill";
import { createGradientFill } from "../fill/gradient-fill";
import type { GradientFillOptions } from "../fill/gradient-fill";
import { createGroupFill } from "../fill/group-fill";
import { createNoFill } from "../fill/no-fill";
import { createPatternFill } from "../fill/pattern-fill";
import type { PatternFillOptions } from "../fill/pattern-fill";
import type { BlurEffectOptions } from "./effect-list";
import { appendCommonEffects } from "./effect-list";
import type { FillOverlayEffectOptions } from "./fill-overlay";
import { BlendMode } from "./fill-overlay";
import type { GlowEffectOptions } from "./glow";
import type { InnerShadowEffectOptions } from "./inner-shadow";
import type { OuterShadowEffectOptions } from "./outer-shadow";
import type { PresetShadowEffectOptions } from "./preset-shadow";
import type { ReflectionEffectOptions } from "./reflection";

// ─── Effect Container Type ──────────────────────────────────────────────────

/**
 * Effect container type (ST_EffectContainerType).
 *
 * Controls how child effects are combined:
 * - `sib`: sibling effects are applied in parallel
 * - `tree`: effects are applied sequentially in tree order
 *
 * Note: the `sib` token maps to the `"sibling"` value accepted here.
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
export type EffectContainer = "sibling" | "tree";

// ─── Effect Options ─────────────────────────────────────────────────────────
// Shared CT_* effect option types live in blip-effects.ts (single home);
// this module re-exports them for the effect-DAG path.

import type {
  AlphaBiLevelEffectOptions,
  AlphaModulateFixedEffectOptions,
  AlphaReplaceEffectOptions,
  ColorChangeEffectOptions,
  DuotoneEffectOptions,
  HSLEffectOptions,
  LuminanceEffectOptions,
  TintEffectOptions,
} from "../blip/blip-effects";

export type {
  AlphaBiLevelEffectOptions,
  AlphaModulateFixedEffectOptions,
  AlphaReplaceEffectOptions,
  ColorChangeEffectOptions,
  DuotoneEffectOptions,
  HSLEffectOptions,
  LuminanceEffectOptions,
  TintEffectOptions,
} from "../blip/blip-effects";

/** Alpha inverse effect — inverts alpha, optionally with color (CT_AlphaInverseEffect). */
export interface AlphaInverseEffectOptions {
  /** Optional color to apply inverse to */
  color?: SolidFillOptions;
}

/** Alpha outset effect (CT_AlphaOutsetEffect). */
export interface AlphaOutsetEffectOptions {
  /** Outset radius */
  radius?: number;
}

/** Blend effect — blends with nested container (CT_BlendEffect). */
export interface BlendEffectOptions {
  /** Blend mode (required) */
  blend: (typeof BlendMode)[keyof typeof BlendMode];
  /** Nested effect container */
  container: EffectDagOptions;
}

/** Fill effect — applies a fill as an effect (CT_FillEffect). */
export interface FillEffectOptions {
  solidFill?: SolidFillOptions;
  gradientFill?: GradientFillOptions;
  patternFill?: PatternFillOptions;
  /** Group fill (inherit from parent) */
  groupFill?: boolean;
  noFill?: boolean;
}

/** Relative offset effect (CT_RelativeOffsetEffect). */
export interface RelativeOffsetEffectOptions {
  /** Horizontal offset percentage (default 0%) */
  translateX?: number;
  /** Vertical offset percentage (default 0%) */
  translateY?: number;
}

/** Transform effect — scale, skew, translate (CT_TransformEffect). */
export interface TransformEffectOptions {
  /** Horizontal scale percentage (default 100%) */
  scaleX?: number;
  /** Vertical scale percentage (default 100%) */
  scaleY?: number;
  /** Horizontal skew angle in degrees (default 0). */
  skewX?: number;
  /** Vertical skew angle in degrees (default 0). */
  skewY?: number;
  /** Horizontal translation */
  translateX?: number;
  /** Vertical translation */
  translateY?: number;
}

/** Effect reference — references an effect by ID (CT_EffectReference). */
export interface EffectReferenceOptions {
  /** Reference ID (required) */
  ref: string;
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
  type?: EffectContainer;
  /** Container name */
  name?: string;

  // ── Existing effects (also in CT_EffectList) ──
  blur?: BlurEffectOptions;
  fillOverlay?: FillOverlayEffectOptions;
  glow?: GlowEffectOptions;
  innerShadow?: InnerShadowEffectOptions;
  outerShadow?: OuterShadowEffectOptions;
  presetShadow?: PresetShadowEffectOptions;
  reflection?: ReflectionEffectOptions | true;
  softEdge?: number;

  // ── New effects for CT_EffectContainer ──
  /** Nested effect containers */
  containers?: readonly EffectDagOptions[];
  /** Effect references */
  effectRefs?: readonly EffectReferenceOptions[];
  /** Alpha bi-level effect */
  alphaBiLevel?: AlphaBiLevelEffectOptions;
  /** Alpha ceiling effect (empty element) */
  alphaCeiling?: boolean;
  /** Alpha floor effect (empty element) */
  alphaFloor?: boolean;
  /** Alpha inverse effect */
  alphaInverse?: AlphaInverseEffectOptions;
  /** Alpha modulate effect (requires nested container) */
  alphaModulate?: EffectDagOptions;
  /** Alpha modulate fixed effect */
  alphaModulateFixed?: AlphaModulateFixedEffectOptions;
  /** Alpha outset effect */
  alphaOutset?: AlphaOutsetEffectOptions;
  /** Alpha replace effect */
  alphaReplace?: AlphaReplaceEffectOptions;
  /** Bi-level effect (threshold) */
  biLevel?: { threshold: number };
  /** Blend effect (requires nested container) */
  blend?: BlendEffectOptions;
  /** Color change effect */
  colorChange?: ColorChangeEffectOptions;
  /** Color replace effect */
  colorReplace?: SolidFillOptions;
  /** Duotone effect */
  duotone?: DuotoneEffectOptions;
  /** Fill effect */
  fill?: FillEffectOptions;
  /** Grayscale effect (empty element) */
  grayscale?: boolean;
  /** HSL effect */
  hsl?: HSLEffectOptions;
  /** Luminance effect */
  luminance?: LuminanceEffectOptions;
  /** Tint effect */
  tint?: TintEffectOptions;
  /** Relative offset effect */
  relativeOffset?: RelativeOffsetEffectOptions;
  /** Transform effect */
  transform?: TransformEffectOptions;
}

// ─── New Effect Factories ───────────────────────────────────────────────────

const createAlphaBiLevelEffect = (options: AlphaBiLevelEffectOptions): string =>
  `<a:alphaBiLevel thresh="${options.threshold}"/>`;

const createAlphaCeilingEffect = (): string => "<a:alphaCeiling/>";

const createAlphaFloorEffect = (): string => "<a:alphaFloor/>";

const createAlphaInverseEffect = (options?: AlphaInverseEffectOptions): string =>
  element(
    "a:alphaInv",
    undefined,
    options?.color ? [createColorElement(options.color)] : undefined,
  );

const createAlphaModulateFixedEffect = (options?: AlphaModulateFixedEffectOptions): string =>
  options?.amount === undefined
    ? "<a:alphaModFix/>"
    : `<a:alphaModFix amt="${Math.round(options.amount * 1000)}"/>`;

const createAlphaOutsetEffect = (options?: AlphaOutsetEffectOptions): string =>
  options?.radius === undefined ? "<a:alphaOutset/>" : `<a:alphaOutset rad="${options.radius}"/>`;

const createAlphaReplaceEffect = (options: AlphaReplaceEffectOptions): string =>
  `<a:alphaRepl a="${Math.round(options.alpha * 1000)}"/>`;

const createBiLevelEffect = (thresh: number): string =>
  `<a:biLevel thresh="${Math.round(thresh * 1000)}"/>`;

const createBlendEffect = (options: BlendEffectOptions): string =>
  element("a:blend", { blend: xsdBlendMode.to(options.blend) }, [
    createEffectContainer(options.container),
  ]);

const createColorChangeEffect = (options: ColorChangeEffectOptions): string => {
  const children: string[] = [
    element("a:clrFrom", undefined, [createColorElement(options.from)]),
    element("a:clrTo", undefined, [createColorElement(options.to)]),
  ];

  if (options.useAlpha === false) {
    return element("a:clrChange", { useA: 0 }, children);
  }

  return element("a:clrChange", undefined, children);
};

const createColorReplaceEffect = (color: SolidFillOptions): string =>
  element("a:clrRepl", undefined, [createColorElement(color)]);

const createDuotoneEffect = (options: DuotoneEffectOptions): string =>
  element("a:duotone", undefined, [
    createColorElement(options.color1),
    createColorElement(options.color2),
  ]);

const createFillEffect = (options: FillEffectOptions): string => {
  let fillElement: string;

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

  return element("a:fill", undefined, [fillElement]);
};

const createGrayscaleEffect = (): string => "<a:grayscl/>";

const createHSLEffect = (options?: HSLEffectOptions): string => {
  const attrs: Record<string, number> = {};
  if (options?.hue !== undefined) attrs.hue = Math.round(options.hue * 60000);
  if (options?.saturation !== undefined) attrs.sat = options.saturation;
  if (options?.luminance !== undefined) attrs.lum = options.luminance;
  return element("a:hsl", Object.keys(attrs).length > 0 ? attrs : undefined);
};

const createLuminanceEffect = (options?: LuminanceEffectOptions): string => {
  const attrs: Record<string, number> = {};
  if (options?.bright !== undefined) attrs.bright = Math.round(options.bright * 1000);
  if (options?.contrast !== undefined) attrs.contrast = Math.round(options.contrast * 1000);
  return element("a:lum", Object.keys(attrs).length > 0 ? attrs : undefined);
};

const createTintEffect = (options?: TintEffectOptions): string => {
  const attrs: Record<string, number> = {};
  if (options?.hue !== undefined) attrs.hue = Math.round(options.hue * 60000);
  if (options?.amount !== undefined) attrs.amt = Math.round(options.amount * 1000);
  return element("a:tint", Object.keys(attrs).length > 0 ? attrs : undefined);
};

const createRelativeOffsetEffect = (options?: RelativeOffsetEffectOptions): string => {
  const attrs: Record<string, number> = {};
  if (options?.translateX !== undefined) attrs.tx = Math.round(options.translateX * 1000);
  if (options?.translateY !== undefined) attrs.ty = Math.round(options.translateY * 1000);
  return element("a:relOff", Object.keys(attrs).length > 0 ? attrs : undefined);
};

const createTransformEffect = (options?: TransformEffectOptions): string => {
  const attrs: Record<string, number> = {};
  if (options?.scaleX !== undefined) attrs.sx = Math.round(options.scaleX * 1000);
  if (options?.scaleY !== undefined) attrs.sy = Math.round(options.scaleY * 1000);
  if (options?.skewX !== undefined) attrs.kx = Math.round(options.skewX * 60000);
  if (options?.skewY !== undefined) attrs.ky = Math.round(options.skewY * 60000);
  if (options?.translateX !== undefined) attrs.tx = options.translateX;
  if (options?.translateY !== undefined) attrs.ty = options.translateY;
  return element("a:xfrm", Object.keys(attrs).length > 0 ? attrs : undefined);
};

const createEffectReference = (options: EffectReferenceOptions): string =>
  `<a:effect ref="${options.ref}"/>`;

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
 * @param elementName - Element name, defaults to "a:cont" for top-level
 */
const createEffectContainer = (options: EffectDagOptions, elementName = "a:cont"): string => {
  const attrs: Record<string, string | number> = {};
  if (options.type) attrs.type = xsdEffectContainer.to(options.type);
  if (options.name) attrs.name = options.name;

  const children: string[] = [];

  // The eight CT_EffectList effects, then the DAG-only ones below.
  appendCommonEffects(options, children);

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
    children.push(element("a:alphaMod", undefined, [createEffectContainer(options.alphaModulate)]));
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
    children.push(createBiLevelEffect(options.biLevel.threshold));
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

  return element(
    elementName,
    Object.keys(attrs).length > 0 ? attrs : undefined,
    children.length > 0 ? children : undefined,
  );
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
 *   glow: { radius: 50800, color: { value: "FF0000" } },
 *   outerShadow: { color: { value: "000000" }, blurRadius: 76200 },
 * });
 *
 * // DAG with alpha effects and nested container
 * createEffectDag({
 *   type: "tree",
 *   alphaBiLevel: { threshold: 50 },
 *   containers: [{
 *     glow: { color: { value: "FF0000" } },
 *   }],
 * });
 * ```
 */
export const createEffectDag = (options: EffectDagOptions): string =>
  createEffectContainer(options, "a:effectDag");
