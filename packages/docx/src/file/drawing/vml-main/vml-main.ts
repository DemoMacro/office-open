/**
 * VML Main elements (v: namespace).
 *
 * Factory functions for VML shape primitives, formulas, handles,
 * image data, and text path elements.
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

// ── Shape Primitives ──

/**
 * Creates a v:rect element (Rectangle).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="rect" type="CT_Rect"/>
 * ```
 */
export const createVmlRect = (): XmlComponent => new BuilderElement({ name: "v:rect" });

/**
 * Creates a v:roundrect element (Rounded Rectangle).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="roundrect" type="CT_RoundRect"/>
 * ```
 */
export const createVmlRoundRect = (): XmlComponent => new BuilderElement({ name: "v:roundrect" });

/**
 * Creates a v:oval element (Oval / Ellipse).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="oval" type="CT_Oval"/>
 * ```
 */
export const createVmlOval = (): XmlComponent => new BuilderElement({ name: "v:oval" });

/**
 * Creates a v:line element (Line).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="line" type="CT_Line"/>
 * ```
 */
export const createVmlLine = (): XmlComponent => new BuilderElement({ name: "v:line" });

export interface VmlArcOptions {
  readonly startAngle?: number;
  readonly endAngle?: number;
}

/**
 * Creates a v:arc element (Arc).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="arc" type="CT_Arc"/>
 * ```
 */
export const createVmlArc = (options?: VmlArcOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: number }> = {};
  if (options?.startAngle !== undefined)
    attrs.startAngle = { key: "startAngle", value: options.startAngle };
  if (options?.endAngle !== undefined)
    attrs.endAngle = { key: "endAngle", value: options.endAngle };
  return new BuilderElement({
    name: "v:arc",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

export interface VmlCurveOptions {
  readonly from?: string;
  readonly control1?: string;
  readonly control2?: string;
  readonly to?: string;
}

/**
 * Creates a v:curve element (Bezier Curve).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="curve" type="CT_Curve"/>
 * ```
 */
export const createVmlCurve = (options?: VmlCurveOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.from !== undefined) attrs.from = { key: "from", value: options.from };
  if (options?.control1 !== undefined)
    attrs.control1 = { key: "control1", value: options.control1 };
  if (options?.control2 !== undefined)
    attrs.control2 = { key: "control2", value: options.control2 };
  if (options?.to !== undefined) attrs.to = { key: "to", value: options.to };
  return new BuilderElement({
    name: "v:curve",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

export interface VmlPolylineOptions {
  readonly points?: string;
}

/**
 * Creates a v:polyline element (Polyline).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="polyline" type="CT_PolyLine"/>
 * ```
 */
export const createVmlPolyline = (options?: VmlPolylineOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.points !== undefined) attrs.points = { key: "points", value: options.points };
  return new BuilderElement({
    name: "v:polyline",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

export interface VmlImageOptions {
  readonly cropLeft?: string;
  readonly cropTop?: string;
  readonly cropRight?: string;
  readonly cropBottom?: string;
  readonly gain?: string;
  readonly blackLevel?: string;
  readonly gamma?: string;
  readonly grayscale?: boolean;
  readonly bilevel?: boolean;
}

/**
 * Creates a v:image element (Image Shape).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="image" type="CT_Image"/>
 * ```
 */
export const createVmlImage = (options?: VmlImageOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.cropLeft !== undefined)
    attrs.cropLeft = { key: "cropleft", value: options.cropLeft };
  if (options?.cropTop !== undefined) attrs.cropTop = { key: "croptop", value: options.cropTop };
  if (options?.cropRight !== undefined)
    attrs.cropRight = { key: "cropright", value: options.cropRight };
  if (options?.cropBottom !== undefined)
    attrs.cropBottom = { key: "cropbottom", value: options.cropBottom };
  if (options?.gain !== undefined) attrs.gain = { key: "gain", value: options.gain };
  if (options?.blackLevel !== undefined)
    attrs.blackLevel = { key: "blacklevel", value: options.blackLevel };
  if (options?.gamma !== undefined) attrs.gamma = { key: "gamma", value: options.gamma };
  if (options?.grayscale !== undefined)
    attrs.grayscale = { key: "grayscale", value: options.grayscale ? "t" : "f" };
  if (options?.bilevel !== undefined)
    attrs.bilevel = { key: "bilevel", value: options.bilevel ? "t" : "f" };
  return new BuilderElement({
    name: "v:image",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

export interface VmlGroupOptions {
  readonly editAs?: "canvas" | "orgchart" | "radial" | "cycle" | "stacked" | "venn" | "bullseye";
  readonly wrapCoords?: string;
}

/**
 * Creates a v:group element (Shape Group).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="group" type="CT_Group"/>
 * ```
 */
export const createVmlGroup = (options?: VmlGroupOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.editAs !== undefined) attrs.editAs = { key: "editas", value: options.editAs };
  if (options?.wrapCoords !== undefined)
    attrs.wrapCoords = { key: "wrapcoords", value: options.wrapCoords };
  return new BuilderElement({
    name: "v:group",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Image Data ──

export interface VmlImageDataOptions {
  readonly cropLeft?: string;
  readonly cropTop?: string;
  readonly cropRight?: string;
  readonly cropBottom?: string;
  readonly gain?: string;
  readonly blackLevel?: string;
  readonly gamma?: string;
  readonly grayscale?: boolean;
  readonly bilevel?: boolean;
  readonly chromaKey?: string;
  readonly embossColor?: string;
  readonly recolorTarget?: string;
}

/**
 * Creates a v:imagedata element (Image Data).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="imagedata" type="CT_ImageData"/>
 * ```
 */
export const createVmlImagedata = (options?: VmlImageDataOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.cropLeft !== undefined)
    attrs.cropLeft = { key: "cropleft", value: options.cropLeft };
  if (options?.cropTop !== undefined) attrs.cropTop = { key: "croptop", value: options.cropTop };
  if (options?.cropRight !== undefined)
    attrs.cropRight = { key: "cropright", value: options.cropRight };
  if (options?.cropBottom !== undefined)
    attrs.cropBottom = { key: "cropbottom", value: options.cropBottom };
  if (options?.gain !== undefined) attrs.gain = { key: "gain", value: options.gain };
  if (options?.blackLevel !== undefined)
    attrs.blackLevel = { key: "blacklevel", value: options.blackLevel };
  if (options?.gamma !== undefined) attrs.gamma = { key: "gamma", value: options.gamma };
  if (options?.grayscale !== undefined)
    attrs.grayscale = { key: "grayscale", value: options.grayscale ? "t" : "f" };
  if (options?.bilevel !== undefined)
    attrs.bilevel = { key: "bilevel", value: options.bilevel ? "t" : "f" };
  if (options?.chromaKey !== undefined)
    attrs.chromaKey = { key: "chromakey", value: options.chromaKey };
  if (options?.embossColor !== undefined)
    attrs.embossColor = { key: "embosscolor", value: options.embossColor };
  if (options?.recolorTarget !== undefined)
    attrs.recolorTarget = { key: "recolortarget", value: options.recolorTarget };
  return new BuilderElement({
    name: "v:imagedata",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Formulas & Handles ──

/**
 * Creates a v:formulas element (Formulas Container).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="formulas" type="CT_Formulas"/>
 * ```
 */
export const createVmlFormulas = (): XmlComponent => new BuilderElement({ name: "v:formulas" });

export interface VmlFormulaOptions {
  readonly equation?: string;
}

/**
 * Creates a v:f element (Formula).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="f" type="CT_F"/>
 * ```
 */
export const createVmlF = (options?: VmlFormulaOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.equation !== undefined) attrs.equation = { key: "eqn", value: options.equation };
  return new BuilderElement({
    name: "v:f",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

/**
 * Creates a v:handles element (Handles Container).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="handles" type="CT_Handles"/>
 * ```
 */
export const createVmlHandles = (): XmlComponent => new BuilderElement({ name: "v:handles" });

export interface VmlHandleOptions {
  readonly position?: string;
  readonly polar?: string;
  readonly map?: string;
  readonly invertX?: boolean;
  readonly invertY?: boolean;
  readonly switchHandle?: boolean;
  readonly xRange?: string;
  readonly yRange?: string;
  readonly radiusRange?: string;
}

/**
 * Creates a v:h element (Handle).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="h" type="CT_H"/>
 * ```
 */
export const createVmlH = (options?: VmlHandleOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.position !== undefined)
    attrs.position = { key: "position", value: options.position };
  if (options?.polar !== undefined) attrs.polar = { key: "polar", value: options.polar };
  if (options?.map !== undefined) attrs.map = { key: "map", value: options.map };
  if (options?.invertX !== undefined)
    attrs.invertX = { key: "invx", value: options.invertX ? "t" : "f" };
  if (options?.invertY !== undefined)
    attrs.invertY = { key: "invy", value: options.invertY ? "t" : "f" };
  if (options?.switchHandle !== undefined)
    attrs.switchHandle = { key: "switch", value: options.switchHandle ? "t" : "f" };
  if (options?.xRange !== undefined) attrs.xRange = { key: "xrange", value: options.xRange };
  if (options?.yRange !== undefined) attrs.yRange = { key: "yrange", value: options.yRange };
  if (options?.radiusRange !== undefined)
    attrs.radiusRange = { key: "radiusrange", value: options.radiusRange };
  return new BuilderElement({
    name: "v:h",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Text Path ──

export interface VmlTextPathOptions {
  readonly on?: boolean;
  readonly fitShape?: boolean;
  readonly fitPath?: boolean;
  readonly trim?: boolean;
  readonly xScale?: boolean;
  readonly string?: string;
}

/**
 * Creates a v:textpath element (Text Path).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="textpath" type="CT_TextPath"/>
 * ```
 */
export const createVmlTextpath = (options?: VmlTextPathOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.on !== undefined) attrs.on = { key: "on", value: options.on ? "t" : "f" };
  if (options?.fitShape !== undefined)
    attrs.fitShape = { key: "fitshape", value: options.fitShape ? "t" : "f" };
  if (options?.fitPath !== undefined)
    attrs.fitPath = { key: "fitpath", value: options.fitPath ? "t" : "f" };
  if (options?.trim !== undefined) attrs.trim = { key: "trim", value: options.trim ? "t" : "f" };
  if (options?.xScale !== undefined)
    attrs.xScale = { key: "xscale", value: options.xScale ? "t" : "f" };
  if (options?.string !== undefined) attrs.string = { key: "string", value: options.string };
  return new BuilderElement({
    name: "v:textpath",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Fill ──

export interface VmlFillOptions {
  readonly type?: "solid" | "gradient" | "gradientRadial" | "tile" | "pattern" | "frame";
  readonly on?: boolean;
  readonly color?: string;
  readonly opacity?: string;
  readonly color2?: string;
  readonly src?: string;
  readonly aspect?: "ignore" | "atMost" | "atLeast";
  readonly colors?: string;
  readonly angle?: number;
  readonly alignShape?: boolean;
  readonly focus?: string;
  readonly focusPosition?: string;
  readonly focusSize?: string;
  readonly method?: "none" | "linear" | "sigma" | "any";
  readonly rotate?: boolean;
  readonly recolor?: boolean;
}

/**
 * Creates a v:fill element (Fill).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="fill" type="CT_Fill"/>
 * ```
 */
export const createVmlFill = (options?: VmlFillOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string | number }> = {};
  if (options?.type !== undefined) attrs.type = { key: "type", value: options.type };
  if (options?.on !== undefined) attrs.on = { key: "on", value: options.on ? "t" : "f" };
  if (options?.color !== undefined) attrs.color = { key: "color", value: options.color };
  if (options?.opacity !== undefined) attrs.opacity = { key: "opacity", value: options.opacity };
  if (options?.color2 !== undefined) attrs.color2 = { key: "color2", value: options.color2 };
  if (options?.src !== undefined) attrs.src = { key: "src", value: options.src };
  if (options?.aspect !== undefined) attrs.aspect = { key: "aspect", value: options.aspect };
  if (options?.colors !== undefined) attrs.colors = { key: "colors", value: options.colors };
  if (options?.angle !== undefined) attrs.angle = { key: "angle", value: options.angle };
  if (options?.alignShape !== undefined)
    attrs.alignShape = { key: "alignshape", value: options.alignShape ? "t" : "f" };
  if (options?.focus !== undefined) attrs.focus = { key: "focus", value: options.focus };
  if (options?.focusPosition !== undefined)
    attrs.focusPosition = { key: "focusposition", value: options.focusPosition };
  if (options?.focusSize !== undefined)
    attrs.focusSize = { key: "focussize", value: options.focusSize };
  if (options?.method !== undefined) attrs.method = { key: "method", value: options.method };
  if (options?.rotate !== undefined)
    attrs.rotate = { key: "rotate", value: options.rotate ? "t" : "f" };
  if (options?.recolor !== undefined)
    attrs.recolor = { key: "recolor", value: options.recolor ? "t" : "f" };
  return new BuilderElement({
    name: "v:fill",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Stroke ──

export interface VmlStrokeOptions {
  readonly on?: boolean;
  readonly weight?: string;
  readonly color?: string;
  readonly opacity?: string;
  readonly lineStyle?: "single" | "thinThin" | "thinThick" | "thickThin" | "thickBetweenThin";
  readonly miterLimit?: number;
  readonly joinStyle?: "round" | "bevel" | "miter";
  readonly endCap?: "flat" | "square" | "round";
  readonly dashStyle?: string;
  readonly fillType?: "solid" | "tile" | "pattern" | "frame";
  readonly src?: string;
  readonly imageAspect?: "ignore" | "atMost" | "atLeast";
  readonly imageSize?: string;
  readonly imageAlignShape?: boolean;
  readonly color2?: string;
  readonly startArrow?: "none" | "block" | "classic" | "oval" | "diamond" | "open";
  readonly startArrowWidth?: "narrow" | "medium" | "wide";
  readonly startArrowLength?: "short" | "medium" | "long";
  readonly endArrow?: "none" | "block" | "classic" | "oval" | "diamond" | "open";
  readonly endArrowWidth?: "narrow" | "medium" | "wide";
  readonly endArrowLength?: "short" | "medium" | "long";
  readonly insetPen?: boolean;
}

/**
 * Creates a v:stroke element (Stroke).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="stroke" type="CT_Stroke"/>
 * ```
 */
export const createVmlStroke = (options?: VmlStrokeOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string | number }> = {};
  if (options?.on !== undefined) attrs.on = { key: "on", value: options.on ? "t" : "f" };
  if (options?.weight !== undefined) attrs.weight = { key: "weight", value: options.weight };
  if (options?.color !== undefined) attrs.color = { key: "color", value: options.color };
  if (options?.opacity !== undefined) attrs.opacity = { key: "opacity", value: options.opacity };
  if (options?.lineStyle !== undefined)
    attrs.lineStyle = { key: "linestyle", value: options.lineStyle };
  if (options?.miterLimit !== undefined)
    attrs.miterLimit = { key: "miterlimit", value: options.miterLimit };
  if (options?.joinStyle !== undefined)
    attrs.joinStyle = { key: "joinstyle", value: options.joinStyle };
  if (options?.endCap !== undefined) attrs.endCap = { key: "endcap", value: options.endCap };
  if (options?.dashStyle !== undefined)
    attrs.dashStyle = { key: "dashstyle", value: options.dashStyle };
  if (options?.fillType !== undefined)
    attrs.fillType = { key: "filltype", value: options.fillType };
  if (options?.src !== undefined) attrs.src = { key: "src", value: options.src };
  if (options?.imageAspect !== undefined)
    attrs.imageAspect = { key: "imageaspect", value: options.imageAspect };
  if (options?.imageSize !== undefined)
    attrs.imageSize = { key: "imagesize", value: options.imageSize };
  if (options?.imageAlignShape !== undefined)
    attrs.imageAlignShape = {
      key: "imagealignshape",
      value: options.imageAlignShape ? "t" : "f",
    };
  if (options?.color2 !== undefined) attrs.color2 = { key: "color2", value: options.color2 };
  if (options?.startArrow !== undefined)
    attrs.startArrow = { key: "startarrow", value: options.startArrow };
  if (options?.startArrowWidth !== undefined)
    attrs.startArrowWidth = { key: "startarrowwidth", value: options.startArrowWidth };
  if (options?.startArrowLength !== undefined)
    attrs.startArrowLength = { key: "startarrowlength", value: options.startArrowLength };
  if (options?.endArrow !== undefined)
    attrs.endArrow = { key: "endarrow", value: options.endArrow };
  if (options?.endArrowWidth !== undefined)
    attrs.endArrowWidth = { key: "endarrowwidth", value: options.endArrowWidth };
  if (options?.endArrowLength !== undefined)
    attrs.endArrowLength = { key: "endarrowlength", value: options.endArrowLength };
  if (options?.insetPen !== undefined)
    attrs.insetPen = { key: "insetpen", value: options.insetPen ? "t" : "f" };
  return new BuilderElement({
    name: "v:stroke",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Shadow ──

export interface VmlShadowOptions {
  readonly on?: boolean;
  readonly type?: "single" | "double" | "emboss" | "perspective";
  readonly obscured?: boolean;
  readonly color?: string;
  readonly opacity?: string;
  readonly offset?: string;
  readonly color2?: string;
  readonly offset2?: string;
  readonly origin?: string;
  readonly matrix?: string;
}

/**
 * Creates a v:shadow element (Shadow).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="shadow" type="CT_Shadow"/>
 * ```
 */
export const createVmlShadow = (options?: VmlShadowOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.on !== undefined) attrs.on = { key: "on", value: options.on ? "t" : "f" };
  if (options?.type !== undefined) attrs.type = { key: "type", value: options.type };
  if (options?.obscured !== undefined)
    attrs.obscured = { key: "obscured", value: options.obscured ? "t" : "f" };
  if (options?.color !== undefined) attrs.color = { key: "color", value: options.color };
  if (options?.opacity !== undefined) attrs.opacity = { key: "opacity", value: options.opacity };
  if (options?.offset !== undefined) attrs.offset = { key: "offset", value: options.offset };
  if (options?.color2 !== undefined) attrs.color2 = { key: "color2", value: options.color2 };
  if (options?.offset2 !== undefined) attrs.offset2 = { key: "offset2", value: options.offset2 };
  if (options?.origin !== undefined) attrs.origin = { key: "origin", value: options.origin };
  if (options?.matrix !== undefined) attrs.matrix = { key: "matrix", value: options.matrix };
  return new BuilderElement({
    name: "v:shadow",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Path ──

export interface VmlPathOptions {
  readonly v?: string;
  readonly limo?: string;
  readonly textboxRect?: string;
  readonly fillOk?: boolean;
  readonly strokeOk?: boolean;
  readonly shadowOk?: boolean;
  readonly arrowOk?: boolean;
  readonly textPathOk?: boolean;
  readonly insetPenOk?: boolean;
}

/**
 * Creates a v:path element (Path).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="path" type="CT_Path"/>
 * ```
 */
export const createVmlPath = (options?: VmlPathOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.v !== undefined) attrs.v = { key: "v", value: options.v };
  if (options?.limo !== undefined) attrs.limo = { key: "limo", value: options.limo };
  if (options?.textboxRect !== undefined)
    attrs.textboxRect = { key: "textboxrect", value: options.textboxRect };
  if (options?.fillOk !== undefined)
    attrs.fillOk = { key: "fillok", value: options.fillOk ? "t" : "f" };
  if (options?.strokeOk !== undefined)
    attrs.strokeOk = { key: "strokeok", value: options.strokeOk ? "t" : "f" };
  if (options?.shadowOk !== undefined)
    attrs.shadowOk = { key: "shadowok", value: options.shadowOk ? "t" : "f" };
  if (options?.arrowOk !== undefined)
    attrs.arrowOk = { key: "arrowok", value: options.arrowOk ? "t" : "f" };
  if (options?.textPathOk !== undefined)
    attrs.textPathOk = { key: "textpathok", value: options.textPathOk ? "t" : "f" };
  if (options?.insetPenOk !== undefined)
    attrs.insetPenOk = { key: "insetpenok", value: options.insetPenOk ? "t" : "f" };
  return new BuilderElement({
    name: "v:path",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};
