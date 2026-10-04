/**
 * Word 2010 text effects (wml-2010.xsd, EG_RPrTextEffects + EG_RPrOpenType).
 *
 * Types mirror the w14 complex types — CT_Glow, CT_Shadow, CT_Reflection,
 * CT_TextOutlineEffect, CT_FillTextEffect, CT_Scene3D, CT_Props3D — whose
 * attribute sets are narrower than their DrawingML counterparts, so a local
 * color type (EG_ColorChoice: srgbClr/schemeClr only) is faithful rather
 * than a reuse of DML's wider fill options.
 *
 * @module
 */
import { convertToEmu, emitAngle, emitPercent, parsePercentAttr } from "@office-open/core";
import {
  xsdCompoundLine,
  xsdLineCap,
  xsdLightRigDirection,
  xsdMaterialType,
  xsdPenAlignment,
  xsdRectAlignment,
} from "@office-open/core";
import type {
  BevelOptions,
  CameraPreset,
  CompoundLine,
  LightRigOptions,
  PenAlignment,
  PresetDash,
  PresetMaterial,
  RectAlignment,
} from "@office-open/core";
import type { Element } from "@office-open/xml";
import { attr, attrBool, attrNum, attrsRaw, children, findChild } from "@office-open/xml";

// ─── Color (EG_ColorChoice) ──────────────────────────────────────────────────

/** w14 scheme slot (ST_SchemeColorVal) — full words; tokens dk1→dark1, phClr→placeholder. */
export type W14SchemeColorValue =
  | "background1"
  | "text1"
  | "background2"
  | "text2"
  | "dark1"
  | "light1"
  | "dark2"
  | "light2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hyperlink"
  | "followedHyperlink"
  | "placeholder";

const xsdW14SchemeColor = {
  to: (value: W14SchemeColorValue): string =>
    value === "background1"
      ? "bg1"
      : value === "text1"
        ? "tx1"
        : value === "background2"
          ? "bg2"
          : value === "text2"
            ? "tx2"
            : value === "dark1"
              ? "dk1"
              : value === "light1"
                ? "lt1"
                : value === "dark2"
                  ? "dk2"
                  : value === "light2"
                    ? "lt2"
                    : value === "hyperlink"
                      ? "hlink"
                      : value === "followedHyperlink"
                        ? "folHlink"
                        : value === "placeholder"
                          ? "phClr"
                          : value,
  from: (token: string): W14SchemeColorValue | undefined => {
    const byToken: Record<string, W14SchemeColorValue> = {
      bg1: "background1",
      tx1: "text1",
      bg2: "background2",
      tx2: "text2",
      dk1: "dark1",
      lt1: "light1",
      dk2: "dark2",
      lt2: "light2",
      hlink: "hyperlink",
      folHlink: "followedHyperlink",
      phClr: "placeholder",
    };
    return byToken[token] ?? (token as W14SchemeColorValue);
  },
} as const;

/** Color transforms in integer percent (50 = 50%). */
export interface W14ColorTransformOptions {
  tint?: number;
  shade?: number;
  alpha?: number;
  hueMod?: number;
  sat?: number;
  satOff?: number;
  satMod?: number;
  lum?: number;
  lumOff?: number;
  lumMod?: number;
}

export interface W14SrgbColorOptions extends W14ColorTransformOptions {
  /** 6-digit hex RGB, e.g. "4472C4". */
  val: string;
}

export interface W14SchemeColorOptions extends W14ColorTransformOptions {
  /** Scheme slot, e.g. "accent1" or "placeholder" (phClr). */
  val: W14SchemeColorValue;
}

export type W14ColorOptions = { srgb: W14SrgbColorOptions } | { scheme: W14SchemeColorOptions };

const W14_TRANSFORMS = [
  ["tint", emitPercent],
  ["shade", emitPercent],
  ["alpha", emitPercent],
  ["hueMod", emitPercent],
  ["sat", emitPercent],
  ["satOff", emitPercent],
  ["satMod", emitPercent],
  ["lum", emitPercent],
  ["lumOff", emitPercent],
  ["lumMod", emitPercent],
] as const;

function w14TransformsXml(transforms: W14ColorTransformOptions): string {
  let xml = "";
  for (const [name, emit] of W14_TRANSFORMS) {
    const percent = transforms[name];
    if (percent !== undefined) xml += `<w14:${name} w14:val="${emit(percent)}"/>`;
  }
  return xml;
}

function w14ColorXml(color: W14ColorOptions): string {
  if ("srgb" in color) {
    return `<w14:srgbClr w14:val="${color.srgb.val}">${w14TransformsXml(color.srgb)}</w14:srgbClr>`;
  }
  return `<w14:schemeClr w14:val="${xsdW14SchemeColor.to(color.scheme.val)}">${w14TransformsXml(color.scheme)}</w14:schemeClr>`;
}

function parseW14Color(el: Element | undefined): W14ColorOptions | undefined {
  if (!el) return undefined;
  const parseTransforms = (colorEl: Element): W14ColorTransformOptions => {
    const t: W14ColorTransformOptions = {};
    for (const [name] of W14_TRANSFORMS) {
      const child = findChild(colorEl, `w14:${name}`);
      if (child) {
        const percent = parsePercentAttr(attr(child, "w14:val"));
        if (percent !== undefined) t[name] = percent;
      }
    }
    return t;
  };
  const srgb = findChild(el, "w14:srgbClr");
  if (srgb) return { srgb: { val: attr(srgb, "w14:val") ?? "", ...parseTransforms(srgb) } };
  const scheme = findChild(el, "w14:schemeClr");
  if (scheme) {
    const token = xsdW14SchemeColor.from(attr(scheme, "w14:val") ?? "");
    if (token) return { scheme: { val: token, ...parseTransforms(scheme) } };
  }
  return undefined;
}

// ─── Fill (EG_FillProperties) ────────────────────────────────────────────────

/** Gradient stop position in integer percent (0 = 0%, 100 = 100%). */
export interface W14GradientStopOptions {
  position: number;
  color: W14ColorOptions;
}

export interface W14PathShadeOptions {
  /** Path shade geometry. */
  shape?: "circle" | "rect" | "shape";
  /** Gradient rectangle inset in integer percent (left/top/right/bottom). */
  fillToRect?: { left?: number; top?: number; right?: number; bottom?: number };
}

export interface W14GradientFillOptions {
  /** Two to ten stops, ordered by position. */
  stops: W14GradientStopOptions[];
  /** Linear shade — angle in degrees. */
  linear?: { angle?: number; scaled?: boolean };
  /** Path shade. */
  path?: W14PathShadeOptions;
}

/**
 * Text fill (CT_FillTextEffect's EG_FillProperties) — at most one branch set.
 */
export interface W14FillOptions {
  noFill?: boolean;
  solid?: W14ColorOptions;
  gradient?: W14GradientFillOptions;
}

function w14FillXml(fill: W14FillOptions): string {
  if (fill.noFill) return "<w14:noFill/>";
  if (fill.solid) {
    return `<w14:solidFill>${w14ColorXml(fill.solid)}</w14:solidFill>`;
  }
  if (fill.gradient) {
    const stops = fill.gradient.stops
      .map(
        (stop) =>
          `<w14:gs w14:pos="${emitPercent(stop.position)}">${w14ColorXml(stop.color)}</w14:gs>`,
      )
      .join("");
    let shade = "";
    if (fill.gradient.linear) {
      const attrs =
        (fill.gradient.linear.angle !== undefined
          ? ` w14:ang="${emitAngle(fill.gradient.linear.angle)}"`
          : "") +
        (fill.gradient.linear.scaled !== undefined
          ? ` w14:scaled="${fill.gradient.linear.scaled ? 1 : 0}"`
          : "");
      shade = `<w14:lin${attrs}/>`;
    } else if (fill.gradient.path) {
      const rect = fill.gradient.path.fillToRect;
      const rectXml = rect
        ? `<w14:fillToRect${attrsRaw({
            "w14:l": rect.left !== undefined ? emitPercent(rect.left) : undefined,
            "w14:t": rect.top !== undefined ? emitPercent(rect.top) : undefined,
            "w14:r": rect.right !== undefined ? emitPercent(rect.right) : undefined,
            "w14:b": rect.bottom !== undefined ? emitPercent(rect.bottom) : undefined,
          })}/>`
        : "";
      shade = `<w14:path${fill.gradient.path.shape !== undefined ? ` w14:path="${fill.gradient.path.shape}"` : ""}>${rectXml}</w14:path>`;
    }
    return `<w14:gradFill><w14:gsLst>${stops}</w14:gsLst>${shade}</w14:gradFill>`;
  }
  return "";
}

function parseW14Fill(el: Element | undefined): W14FillOptions | undefined {
  if (!el) return undefined;
  if (findChild(el, "w14:noFill")) return { noFill: true };
  const solid = findChild(el, "w14:solidFill");
  if (solid) {
    const color = parseW14Color(solid);
    return color ? { solid: color } : undefined;
  }
  const grad = findChild(el, "w14:gradFill");
  if (!grad) return undefined;
  const stops: W14GradientStopOptions[] = [];
  const gsLst = findChild(grad, "w14:gsLst");
  if (gsLst) {
    for (const gs of children(gsLst, "w14:gs")) {
      const rawPos = attr(gs, "w14:pos");
      const position = parsePercentAttr(rawPos);
      const color = parseW14Color(gs);
      if (position !== undefined && color) stops.push({ position, color });
    }
  }
  const fill: W14GradientFillOptions = { stops };
  const lin = findChild(grad, "w14:lin");
  if (lin) {
    const angle = attrNum(lin, "w14:ang");
    fill.linear = {
      angle: angle !== undefined ? angle / 60000 : undefined,
      scaled: attrBool(lin, "w14:scaled"),
    };
  }
  const path = findChild(grad, "w14:path");
  if (path) {
    const shape = attr(path, "w14:path");
    const rectEl = findChild(path, "w14:fillToRect");
    const rectPercent = (name: string): number | undefined => {
      if (!rectEl) return undefined;
      return parsePercentAttr(attr(rectEl, name));
    };
    fill.path = {
      shape: shape === "circle" || shape === "rect" || shape === "shape" ? shape : undefined,
      fillToRect: rectEl
        ? {
            left: rectPercent("w14:l"),
            top: rectPercent("w14:t"),
            right: rectPercent("w14:r"),
            bottom: rectPercent("w14:b"),
          }
        : undefined,
    };
  }
  return { gradient: fill };
}

// ─── Text outline (CT_TextOutlineEffect) ─────────────────────────────────────

export interface W14TextOutlineOptions {
  fill?: W14FillOptions;
  dash?: PresetDash;
  join?: "round" | "bevel" | { miterLimit: number };
  /** Line width in EMU (12700 = 1pt). */
  width?: number;
  cap?: "round" | "square" | "flat";
  compound?: CompoundLine;
  alignment?: PenAlignment;
}

function w14TextOutlineXml(opts: W14TextOutlineOptions): string {
  let inner = opts.fill ? w14FillXml(opts.fill) : "";
  if (opts.dash !== undefined) inner += `<w14:prstDash w14:val="${opts.dash}"/>`;
  if (opts.join === "round") inner += "<w14:round/>";
  else if (opts.join === "bevel") inner += "<w14:bevel/>";
  else if (typeof opts.join === "object")
    inner += `<w14:miter w14:lim="${emitPercent(opts.join.miterLimit)}"/>`;
  const attrs = attrsRaw({
    "w14:w": opts.width !== undefined ? convertToEmu(opts.width) : undefined,
    "w14:cap": opts.cap !== undefined ? xsdLineCap.to(opts.cap) : undefined,
    "w14:cmpd": opts.compound !== undefined ? xsdCompoundLine.to(opts.compound) : undefined,
    "w14:algn": opts.alignment !== undefined ? xsdPenAlignment.to(opts.alignment) : undefined,
  });
  return `<w14:textOutline${attrs}>${inner}</w14:textOutline>`;
}

function parseW14TextOutline(el: Element): W14TextOutlineOptions | undefined {
  const opts: W14TextOutlineOptions = {};
  opts.fill = parseW14Fill(el);
  const dash = findChild(el, "w14:prstDash");
  if (dash) opts.dash = attr(dash, "w14:val") as PresetDash;
  if (findChild(el, "w14:round")) opts.join = "round";
  else if (findChild(el, "w14:bevel")) opts.join = "bevel";
  else {
    const miter = findChild(el, "w14:miter");
    if (miter) {
      const limit = parsePercentAttr(attr(miter, "w14:lim"));
      if (limit !== undefined) opts.join = { miterLimit: limit };
    }
  }
  const width = attrNum(el, "w14:w");
  if (width !== undefined) opts.width = width;
  const cap = attr(el, "w14:cap");
  if (cap) opts.cap = xsdLineCap.from(cap) as typeof opts.cap;
  const compound = attr(el, "w14:cmpd");
  if (compound) opts.compound = xsdCompoundLine.from(compound) as typeof opts.compound;
  const alignment = attr(el, "w14:algn");
  if (alignment) opts.alignment = xsdPenAlignment.from(alignment) as typeof opts.alignment;
  return opts;
}

// ─── Effects ─────────────────────────────────────────────────────────────────

export interface W14GlowOptions {
  /** Glow radius in EMU (63500 = 5pt). */
  radius?: number;
  color: W14ColorOptions;
}

export interface W14ShadowOptions {
  blurRadius?: number;
  distance?: number;
  /** Shadow direction in degrees. */
  direction?: number;
  horizontalScale?: number;
  verticalScale?: number;
  /** Horizontal skew in degrees. */
  horizontalSkew?: number;
  /** Vertical skew in degrees. */
  verticalSkew?: number;
  alignment?: RectAlignment;
  color: W14ColorOptions;
}

export interface W14ReflectionOptions {
  blurRadius?: number;
  startAlpha?: number;
  startPosition?: number;
  endAlpha?: number;
  endPosition?: number;
  distance?: number;
  direction?: number;
  fadeDirection?: number;
  horizontalScale?: number;
  verticalScale?: number;
  horizontalSkew?: number;
  verticalSkew?: number;
  alignment?: RectAlignment;
}

export interface W14Scene3dOptions {
  camera: { preset: CameraPreset };
  lightRig: LightRigOptions;
}

export interface W14Props3dOptions {
  topBevel?: BevelOptions;
  bottomBevel?: BevelOptions;
  extrusionColor?: W14ColorOptions;
  contourColor?: W14ColorOptions;
  extrusionHeight?: number;
  contourWidth?: number;
  material?: PresetMaterial;
}

const emuAttr = (value: number | undefined): number | undefined =>
  value !== undefined ? convertToEmu(value) : undefined;

const angleAttr = (value: number | undefined): number | undefined =>
  value !== undefined ? emitAngle(value) : undefined;

const percentAttr = (value: number | undefined): number | undefined =>
  value !== undefined ? emitPercent(value) : undefined;

function w14ShadowXml(opts: W14ShadowOptions): string {
  const attrs = attrsRaw({
    "w14:blurRad": emuAttr(opts.blurRadius),
    "w14:dist": emuAttr(opts.distance),
    "w14:dir": angleAttr(opts.direction),
    "w14:sx": percentAttr(opts.horizontalScale),
    "w14:sy": percentAttr(opts.verticalScale),
    "w14:kx": angleAttr(opts.horizontalSkew),
    "w14:ky": angleAttr(opts.verticalSkew),
    "w14:algn": opts.alignment !== undefined ? xsdRectAlignment.to(opts.alignment) : undefined,
  });
  return `<w14:shadow${attrs}>${w14ColorXml(opts.color)}</w14:shadow>`;
}

function w14ReflectionXml(opts: W14ReflectionOptions): string {
  const attrs = attrsRaw({
    "w14:blurRad": emuAttr(opts.blurRadius),
    "w14:stA": percentAttr(opts.startAlpha),
    "w14:stPos": percentAttr(opts.startPosition),
    "w14:endA": percentAttr(opts.endAlpha),
    "w14:endPos": percentAttr(opts.endPosition),
    "w14:dist": emuAttr(opts.distance),
    "w14:dir": angleAttr(opts.direction),
    "w14:fadeDir": angleAttr(opts.fadeDirection),
    "w14:sx": percentAttr(opts.horizontalScale),
    "w14:sy": percentAttr(opts.verticalScale),
    "w14:kx": angleAttr(opts.horizontalSkew),
    "w14:ky": angleAttr(opts.verticalSkew),
    "w14:algn": opts.alignment !== undefined ? xsdRectAlignment.to(opts.alignment) : undefined,
  });
  return `<w14:reflection${attrs}/>`;
}

function w14Scene3dXml(opts: W14Scene3dOptions): string {
  const rot = opts.lightRig.rotation;
  const rotXml = rot
    ? `<w14:rot w14:lat="${emitAngle(rot.lat)}" w14:lon="${emitAngle(rot.lon)}" w14:rev="${emitAngle(rot.rev)}"/>`
    : "";
  return (
    "<w14:scene3d>" +
    `<w14:camera w14:prst="${opts.camera.preset}"/>` +
    `<w14:lightRig w14:rig="${opts.lightRig.rig}" w14:dir="${xsdLightRigDirection.to(opts.lightRig.direction)}">${rotXml}</w14:lightRig>` +
    "</w14:scene3d>"
  );
}

function w14Props3dXml(opts: W14Props3dOptions): string {
  const bevel = (name: "w14:bevelT" | "w14:bevelB", bevel?: BevelOptions): string =>
    bevel
      ? `<${name}${attrsRaw({
          "w14:w": emuAttr(bevel.w),
          "w14:h": emuAttr(bevel.h),
          "w14:prst": bevel.prst,
        })}/>`
      : "";
  const color = (name: "w14:extrusionClr" | "w14:contourClr", value?: W14ColorOptions): string =>
    value ? `<${name}>${w14ColorXml(value)}</${name}>` : "";
  const attrs = attrsRaw({
    "w14:extrusionH": emuAttr(opts.extrusionHeight),
    "w14:contourW": emuAttr(opts.contourWidth),
    "w14:prstMaterial": opts.material !== undefined ? xsdMaterialType.to(opts.material) : undefined,
  });
  return (
    `<w14:props3d${attrs}>` +
    bevel("w14:bevelT", opts.topBevel) +
    bevel("w14:bevelB", opts.bottomBevel) +
    color("w14:extrusionClr", opts.extrusionColor) +
    color("w14:contourClr", opts.contourColor) +
    `</w14:props3d>`
  );
}

// ─── OpenType features (EG_RPrOpenType) ──────────────────────────────────────

/** Digit form (CT_NumForm). */
export type W14NumberForm = "default" | "lining" | "oldStyle";

/** Digit spacing (CT_NumSpacing). */
export type W14NumberSpacing = "default" | "proportional" | "tabular";

/** Ligature set (CT_Ligatures). */
export type W14Ligatures =
  | "none"
  | "standard"
  | "contextual"
  | "historical"
  | "discretional"
  | "standardContextual"
  | "standardHistorical"
  | "contextualHistorical"
  | "standardDiscretional"
  | "contextualDiscretional"
  | "historicalDiscretional"
  | "standardContextualHistorical"
  | "standardContextualDiscretional"
  | "standardHistoricalDiscretional"
  | "contextualHistoricalDiscretional"
  | "all";

export interface W14StylisticSetOptions {
  /** OpenType stylistic set ID (1–20). */
  id: number;
  enabled?: boolean;
}

export interface W14StylisticSetsOptions {
  sets?: W14StylisticSetOptions[];
}

// ─── rPr integration ─────────────────────────────────────────────────────────

/** Typed w14 children of w:rPr, in EG_RPrTextEffects then EG_RPrOpenType order. */
export interface W14RunEffectsOptions {
  glow?: W14GlowOptions;
  shadowEffect?: W14ShadowOptions;
  reflection?: W14ReflectionOptions;
  textOutline?: W14TextOutlineOptions;
  textFill?: W14FillOptions;
  scene3d?: W14Scene3dOptions;
  props3d?: W14Props3dOptions;
  ligatures?: W14Ligatures;
  numForm?: W14NumberForm;
  numSpacing?: W14NumberSpacing;
  stylisticSets?: W14StylisticSetsOptions;
  cntxtAlts?: boolean;
}

/** Emit the w14 extension children for w:rPr, or "" when none are set. */
export function w14RunEffectsXml(opts: W14RunEffectsOptions): string {
  let xml = "";
  if (opts.glow)
    xml += `<w14:glow${opts.glow.radius !== undefined ? ` w14:rad="${convertToEmu(opts.glow.radius)}"` : ""}>${w14ColorXml(opts.glow.color)}</w14:glow>`;
  if (opts.shadowEffect) xml += w14ShadowXml(opts.shadowEffect);
  if (opts.reflection) xml += w14ReflectionXml(opts.reflection);
  if (opts.textOutline) xml += w14TextOutlineXml(opts.textOutline);
  if (opts.textFill) {
    const fill = w14FillXml(opts.textFill);
    if (fill) xml += `<w14:textFill>${fill}</w14:textFill>`;
  }
  if (opts.scene3d) xml += w14Scene3dXml(opts.scene3d);
  if (opts.props3d) xml += w14Props3dXml(opts.props3d);
  if (opts.ligatures !== undefined) xml += `<w14:ligatures w14:val="${opts.ligatures}"/>`;
  if (opts.numForm !== undefined) xml += `<w14:numForm w14:val="${opts.numForm}"/>`;
  if (opts.numSpacing !== undefined) xml += `<w14:numSpacing w14:val="${opts.numSpacing}"/>`;
  if (opts.stylisticSets?.sets?.length) {
    xml += `<w14:stylisticSets>${opts.stylisticSets.sets
      .map(
        (set) =>
          `<w14:styleSet w14:id="${set.id}"${set.enabled !== undefined ? ` w14:val="${set.enabled ? 1 : 0}"` : ""}/>`,
      )
      .join("")}</w14:stylisticSets>`;
  }
  if (opts.cntxtAlts !== undefined) xml += `<w14:cntxtAlts w14:val="${opts.cntxtAlts ? 1 : 0}"/>`;
  return xml;
}

/** Parse one w14:rPr child into a typed effect option, or undefined to keep raw. */
export function parseW14RunEffect(
  name: string,
  el: Element,
): [keyof W14RunEffectsOptions, unknown] | undefined {
  switch (name) {
    case "w14:glow": {
      const radius = attrNum(el, "w14:rad");
      const color = parseW14Color(el);
      if (!color) return undefined;
      return ["glow", { radius, color }];
    }
    case "w14:shadow": {
      const color = parseW14Color(el);
      if (!color) return undefined;
      const opts: W14ShadowOptions = { color };
      const blur = attrNum(el, "w14:blurRad");
      if (blur !== undefined) opts.blurRadius = blur;
      const dist = attrNum(el, "w14:dist");
      if (dist !== undefined) opts.distance = dist;
      const dir = attrNum(el, "w14:dir");
      if (dir !== undefined) opts.direction = dir / 60000;
      const sx = parsePercentAttr(attr(el, "w14:sx"));
      if (sx !== undefined) opts.horizontalScale = sx;
      const sy = parsePercentAttr(attr(el, "w14:sy"));
      if (sy !== undefined) opts.verticalScale = sy;
      const kx = attrNum(el, "w14:kx");
      if (kx !== undefined) opts.horizontalSkew = kx / 60000;
      const ky = attrNum(el, "w14:ky");
      if (ky !== undefined) opts.verticalSkew = ky / 60000;
      const algn = attr(el, "w14:algn");
      if (algn) opts.alignment = xsdRectAlignment.from(algn) as typeof opts.alignment;
      return ["shadowEffect", opts];
    }
    case "w14:reflection": {
      const opts: W14ReflectionOptions = {};
      const blur = attrNum(el, "w14:blurRad");
      if (blur !== undefined) opts.blurRadius = blur;
      const stA = parsePercentAttr(attr(el, "w14:stA"));
      if (stA !== undefined) opts.startAlpha = stA;
      const stPos = parsePercentAttr(attr(el, "w14:stPos"));
      if (stPos !== undefined) opts.startPosition = stPos;
      const endA = parsePercentAttr(attr(el, "w14:endA"));
      if (endA !== undefined) opts.endAlpha = endA;
      const endPos = parsePercentAttr(attr(el, "w14:endPos"));
      if (endPos !== undefined) opts.endPosition = endPos;
      const dist = attrNum(el, "w14:dist");
      if (dist !== undefined) opts.distance = dist;
      const dir = attrNum(el, "w14:dir");
      if (dir !== undefined) opts.direction = dir / 60000;
      const fadeDir = attrNum(el, "w14:fadeDir");
      if (fadeDir !== undefined) opts.fadeDirection = fadeDir / 60000;
      const sx = parsePercentAttr(attr(el, "w14:sx"));
      if (sx !== undefined) opts.horizontalScale = sx;
      const sy = parsePercentAttr(attr(el, "w14:sy"));
      if (sy !== undefined) opts.verticalScale = sy;
      const kx = attrNum(el, "w14:kx");
      if (kx !== undefined) opts.horizontalSkew = kx / 60000;
      const ky = attrNum(el, "w14:ky");
      if (ky !== undefined) opts.verticalSkew = ky / 60000;
      const algn = attr(el, "w14:algn");
      if (algn) opts.alignment = xsdRectAlignment.from(algn) as typeof opts.alignment;
      return ["reflection", opts];
    }
    case "w14:textOutline": {
      const opts = parseW14TextOutline(el);
      return opts ? ["textOutline", opts] : undefined;
    }
    case "w14:textFill": {
      const fill = parseW14Fill(el);
      return fill ? ["textFill", fill] : undefined;
    }
    case "w14:scene3d": {
      const camera = findChild(el, "w14:camera");
      const lightRig = findChild(el, "w14:lightRig");
      if (!camera || !lightRig) return undefined;
      const preset = attr(camera, "w14:prst");
      const rig = attr(lightRig, "w14:rig");
      const dir = attr(lightRig, "w14:dir");
      if (!preset || !rig || !dir) return undefined;
      const opts: W14Scene3dOptions = {
        camera: { preset: preset as CameraPreset },
        lightRig: {
          rig: rig as LightRigOptions["rig"],
          direction: xsdLightRigDirection.from(dir) as LightRigOptions["direction"],
        },
      };
      const rot = findChild(lightRig, "w14:rot");
      if (rot) {
        const lat = attrNum(rot, "w14:lat");
        const lon = attrNum(rot, "w14:lon");
        const rev = attrNum(rot, "w14:rev");
        if (lat !== undefined && lon !== undefined && rev !== undefined) {
          opts.lightRig.rotation = { lat: lat / 60000, lon: lon / 60000, rev: rev / 60000 };
        }
      }
      return ["scene3d", opts];
    }
    case "w14:props3d": {
      const opts: W14Props3dOptions = {};
      const bevel = (name: "w14:bevelT" | "w14:bevelB"): BevelOptions | undefined => {
        const child = findChild(el, name);
        if (!child) return undefined;
        const bevelOpts: BevelOptions = {};
        const w = attrNum(child, "w14:w");
        if (w !== undefined) bevelOpts.w = w;
        const h = attrNum(child, "w14:h");
        if (h !== undefined) bevelOpts.h = h;
        const prst = attr(child, "w14:prst");
        if (prst) bevelOpts.prst = prst as BevelOptions["prst"];
        return bevelOpts;
      };
      const topBevel = bevel("w14:bevelT");
      if (topBevel) opts.topBevel = topBevel;
      const bottomBevel = bevel("w14:bevelB");
      if (bottomBevel) opts.bottomBevel = bottomBevel;
      const extrusionColor = parseW14Color(findChild(el, "w14:extrusionClr"));
      if (extrusionColor) opts.extrusionColor = extrusionColor;
      const contourColor = parseW14Color(findChild(el, "w14:contourClr"));
      if (contourColor) opts.contourColor = contourColor;
      const extrusionH = attrNum(el, "w14:extrusionH");
      if (extrusionH !== undefined) opts.extrusionHeight = extrusionH;
      const contourW = attrNum(el, "w14:contourW");
      if (contourW !== undefined) opts.contourWidth = contourW;
      const material = attr(el, "w14:prstMaterial");
      if (material) opts.material = xsdMaterialType.from(material) as typeof opts.material;
      return ["props3d", opts];
    }
    case "w14:ligatures":
      return ["ligatures", attr(el, "w14:val") as W14Ligatures];
    case "w14:numForm":
      return ["numForm", attr(el, "w14:val") as W14NumberForm];
    case "w14:numSpacing":
      return ["numSpacing", attr(el, "w14:val") as W14NumberSpacing];
    case "w14:stylisticSets": {
      const sets: W14StylisticSetOptions[] = [];
      for (const set of children(el, "w14:styleSet")) {
        const id = attrNum(set, "w14:id");
        if (id !== undefined) sets.push({ id, enabled: attrBool(set, "w14:val") });
      }
      return ["stylisticSets", { sets }];
    }
    case "w14:cntxtAlts":
      return ["cntxtAlts", attrBool(el, "w14:val") ?? true];
    default:
      return undefined;
  }
}
