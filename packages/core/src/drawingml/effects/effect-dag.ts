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
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";
import { xsdBlendMode, xsdEffectContainer } from "../../xsd-mappings";
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
import type { FillOverlayEffectOptions } from "./fill-overlay";
import { BlendMode } from "./fill-overlay";
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
  SIB: "sibling",
  TREE: "tree",
} as const;

// ─── New Effect Options ─────────────────────────────────────────────────────

/** Alpha bi-level effect — clips alpha to threshold (CT_AlphaBiLevelEffect). */
export interface AlphaBiLevelEffectOptions {
  /** Alpha threshold (fixed percentage, required) */
  threshold: number;
}

/** Alpha inverse effect — inverts alpha, optionally with color (CT_AlphaInverseEffect). */
export interface AlphaInverseEffectOptions {
  /** Optional color to apply inverse to */
  color?: SolidFillOptions;
}

/** Alpha modulate fixed effect (CT_AlphaModulateFixedEffect). */
export interface AlphaModulateFixedEffectOptions {
  /** Amount percentage (default 100%) */
  amount?: number;
}

/** Alpha outset effect (CT_AlphaOutsetEffect). */
export interface AlphaOutsetEffectOptions {
  /** Outset radius */
  radius?: number;
}

/** Alpha replace effect — replaces alpha value (CT_AlphaReplaceEffect). */
export interface AlphaReplaceEffectOptions {
  /** New alpha value (fixed percentage, required) */
  alpha: number;
}

/** Blend effect — blends with nested container (CT_BlendEffect). */
export interface BlendEffectOptions {
  /** Blend mode (required) */
  blend: (typeof BlendMode)[keyof typeof BlendMode];
  /** Nested effect container */
  container: EffectDagOptions;
}

/** Color change effect — maps one color to another (CT_ColorChangeEffect). */
export interface ColorChangeEffectOptions {
  /** Source color to replace */
  from: SolidFillOptions;
  /** Target color */
  to: SolidFillOptions;
  /** Whether to use alpha channel (default true) */
  useA?: boolean;
}

/** Duotone effect — two-color tone mapping (CT_DuotoneEffect). */
export interface DuotoneEffectOptions {
  /** First color */
  color1: SolidFillOptions;
  /** Second color */
  color2: SolidFillOptions;
}

/** Fill effect — applies a fill as an effect (CT_FillEffect). */
export interface FillEffectOptions {
  /** Solid fill */
  solidFill?: SolidFillOptions;
  /** Gradient fill */
  gradientFill?: GradientFillOptions;
  /** Pattern fill */
  patternFill?: PatternFillOptions;
  /** Group fill (inherit from parent) */
  groupFill?: boolean;
  /** No fill */
  noFill?: boolean;
}

/** HSL effect — adjusts hue, saturation, luminance (CT_HSLEffect). */
export interface HSLEffectOptions {
  /** Hue angle in 60,000ths of a degree (default 0) */
  hue?: number;
  /** Saturation percentage (default 0%) */
  saturation?: number;
  /** Luminance percentage (default 0%) */
  luminance?: number;
}

/** Luminance effect — adjusts brightness and contrast (CT_LuminanceEffect). */
export interface LuminanceEffectOptions {
  /** Brightness percentage (default 0%) */
  bright?: number;
  /** Contrast percentage (default 0%) */
  contrast?: number;
}

/** Tint effect — applies tint with hue and amount (CT_TintEffect). */
export interface TintEffectOptions {
  /** Hue angle in 60,000ths of a degree (default 0) */
  hue?: number;
  /** Tint amount percentage (default 0%) */
  amount?: number;
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
  /** Horizontal skew angle (default 0) */
  skewX?: number;
  /** Vertical skew angle (default 0) */
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
  type?: (typeof EffectContainerType)[keyof typeof EffectContainerType];
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
    attributes: { thresh: { key: "thresh", value: options.threshold } },
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
  if (options?.amount === undefined) {
    return new BuilderElement({ name: "a:alphaModFix" });
  }
  return new BuilderElement<{ readonly amt: number }>({
    name: "a:alphaModFix",
    attributes: { amt: { key: "amt", value: options.amount } },
  });
};

const createAlphaOutsetEffect = (options?: AlphaOutsetEffectOptions): XmlComponent => {
  if (options?.radius === undefined) {
    return new BuilderElement({ name: "a:alphaOutset" });
  }
  return new BuilderElement<{ readonly rad: number }>({
    name: "a:alphaOutset",
    attributes: { rad: { key: "rad", value: options.radius } },
  });
};

const createAlphaReplaceEffect = (options: AlphaReplaceEffectOptions): XmlComponent =>
  new BuilderElement<{ readonly a: number }>({
    name: "a:alphaRepl",
    attributes: { a: { key: "a", value: options.alpha } },
  });

const createBiLevelEffect = (thresh: number): XmlComponent =>
  new BuilderElement<{ readonly thresh: number }>({
    name: "a:biLevel",
    attributes: { thresh: { key: "thresh", value: thresh } },
  });

const createBlendEffect = (options: BlendEffectOptions): XmlComponent =>
  new BuilderElement<{ readonly blend: string }>({
    name: "a:blend",
    attributes: { blend: { key: "blend", value: xsdBlendMode.to(options.blend) } },
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
    sat: options?.saturation,
    lum: options?.luminance,
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
    amt: options?.amount,
  });
  return new BuilderElement({
    name: "a:tint",
    attributes: attrs,
  });
};

const createRelativeOffsetEffect = (options?: RelativeOffsetEffectOptions): XmlComponent => {
  const attrs = buildOptionalAttributes({
    tx: options?.translateX,
    ty: options?.translateY,
  });
  return new BuilderElement({
    name: "a:relOff",
    attributes: attrs,
  });
};

const createTransformEffect = (options?: TransformEffectOptions): XmlComponent => {
  const attrs = buildOptionalAttributes({
    sx: options?.scaleX,
    sy: options?.scaleY,
    kx: options?.skewX,
    ky: options?.skewY,
    tx: options?.translateX,
    ty: options?.translateY,
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
  const hasAttributes = options.radius !== undefined || options.grow === false;

  if (!hasAttributes) {
    return new BuilderElement({ name: "a:blur" });
  }

  const attributePayload = {
    ...(options.radius !== undefined && {
      rad: { key: "rad", value: options.radius },
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
    type: options.type ? xsdEffectContainer.to(options.type) : undefined,
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
 *   glow: { radius: 50800, color: { value: "FF0000" } },
 *   outerShadow: { color: { value: "000000" }, blurRadius: 76200 },
 * });
 *
 * // DAG with alpha effects and nested container
 * createEffectDag({
 *   type: "tree",
 *   alphaBiLevel: { threshold: 50000 },
 *   containers: [{
 *     glow: { color: { value: "FF0000" } },
 *   }],
 * });
 * ```
 */
export const createEffectDag = (options: EffectDagOptions): XmlComponent =>
  createEffectContainer(options, "a:effectDag");
