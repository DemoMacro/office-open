/**
 * VML Office Drawing elements (o: namespace).
 *
 * Factory functions for Office-specific VML elements including shape layout,
 * defaults, locks, extrusion, callout, skew, diagram, and OLE helper elements.
 *
 * Reference: ISO/IEC 29500-4, vml-officeDrawing.xsd
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

// ── Shape Layout & Defaults ──

export interface OShapeLayoutOptions {
  readonly ext?: string;
}

/**
 * Creates an o:shapelayout element (Shape Layout).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="shapelayout" type="CT_ShapeLayout"/>
 * ```
 */
export const createOShapelayout = (options?: OShapeLayoutOptions): XmlComponent =>
  new BuilderElement({
    name: "o:shapelayout",
    attributes:
      options?.ext !== undefined ? { ext: { key: "v:ext", value: options.ext } } : undefined,
  });

export interface OIdMapOptions {
  readonly ext?: string;
  readonly data?: string;
}

/**
 * Creates an o:idmap element (ID Map).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="idmap" type="CT_IdMap"/>
 * ```
 */
export const createOIdmap = (options?: OIdMapOptions): XmlComponent =>
  new BuilderElement({
    name: "o:idmap",
    attributes: {
      ...(options?.ext !== undefined ? { ext: { key: "v:ext", value: options.ext } } : {}),
      ...(options?.data !== undefined ? { data: { key: "data", value: options.data } } : {}),
    },
  });

export interface OShapeDefaultsOptions {
  readonly ext?: string;
  readonly spidMax?: number;
  readonly style?: string;
  readonly fill?: string;
  readonly stroke?: string;
  readonly allowInCell?: boolean;
  readonly fillColor?: string;
  readonly strokeColor?: string;
}

/**
 * Creates an o:shapedefaults element (Shape Defaults).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="shapedefaults" type="CT_ShapeDefaults"/>
 * ```
 */
export const createOShapedefaults = (options?: OShapeDefaultsOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string | number }> = {};
  if (options?.ext !== undefined) attrs.ext = { key: "v:ext", value: options.ext };
  if (options?.spidMax !== undefined) attrs.spidMax = { key: "spidmax", value: options.spidMax };
  if (options?.style !== undefined) attrs.style = { key: "style", value: options.style };
  if (options?.fill !== undefined) attrs.fill = { key: "fill", value: options.fill };
  if (options?.stroke !== undefined) attrs.stroke = { key: "stroke", value: options.stroke };
  if (options?.allowInCell !== undefined)
    attrs.allowInCell = { key: "allowincell", value: options.allowInCell ? "t" : "f" };
  if (options?.fillColor !== undefined)
    attrs.fillColor = { key: "fillcolor", value: options.fillColor };
  if (options?.strokeColor !== undefined)
    attrs.strokeColor = { key: "strokecolor", value: options.strokeColor };
  return new BuilderElement({
    name: "o:shapedefaults",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Color ──

export interface OColorMruOptions {
  readonly ext?: string;
  readonly colors?: string;
}

/**
 * Creates an o:colormru element (Color Most Recently Used).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="colormru" type="CT_ColorMru"/>
 * ```
 */
export const createOColormru = (options?: OColorMruOptions): XmlComponent =>
  new BuilderElement({
    name: "o:colormru",
    attributes: {
      ...(options?.ext !== undefined ? { ext: { key: "v:ext", value: options.ext } } : {}),
      ...(options?.colors !== undefined
        ? { colors: { key: "colors", value: options.colors } }
        : {}),
    },
  });

/**
 * Creates an o:colormenu element (Color Menu).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="colormenu" type="CT_ColorMenu"/>
 * ```
 */
export interface OColorMenuOptions {
  readonly ext?: string;
  readonly strokeColor?: string;
  readonly fillColor?: string;
  readonly shadowColor?: string;
  readonly extrusionColor?: string;
}

export const createOColormenu = (options?: OColorMenuOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.ext !== undefined) attrs.ext = { key: "v:ext", value: options.ext };
  if (options?.strokeColor !== undefined)
    attrs.strokeColor = { key: "strokecolor", value: options.strokeColor };
  if (options?.fillColor !== undefined)
    attrs.fillColor = { key: "fillcolor", value: options.fillColor };
  if (options?.shadowColor !== undefined)
    attrs.shadowColor = { key: "shadowcolor", value: options.shadowColor };
  if (options?.extrusionColor !== undefined)
    attrs.extrusionColor = { key: "extrusioncolor", value: options.extrusionColor };
  return new BuilderElement({
    name: "o:colormenu",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Lock ──

export interface OLockOptions {
  readonly ext?: string;
  readonly position?: boolean;
  readonly selection?: boolean;
  readonly grouping?: boolean;
  readonly ungrouping?: boolean;
  readonly rotation?: boolean;
  readonly cropping?: boolean;
  readonly vertices?: boolean;
  readonly adjustHandles?: boolean;
  readonly text?: boolean;
  readonly aspectRatio?: boolean;
  readonly formatting?: boolean;
  readonly locking?: boolean;
  readonly shapeType?: boolean;
}

/**
 * Creates an o:lock element (Shape Lock).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="lock" type="CT_Lock"/>
 * ```
 */
export const createOLock = (options?: OLockOptions): XmlComponent => {
  const boolMap: readonly { readonly prop: keyof OLockOptions; readonly xmlKey: string }[] = [
    { prop: "position", xmlKey: "position" },
    { prop: "selection", xmlKey: "selection" },
    { prop: "grouping", xmlKey: "grouping" },
    { prop: "ungrouping", xmlKey: "ungrouping" },
    { prop: "rotation", xmlKey: "rotation" },
    { prop: "cropping", xmlKey: "cropping" },
    { prop: "vertices", xmlKey: "verticies" },
    { prop: "adjustHandles", xmlKey: "adjusthandles" },
    { prop: "text", xmlKey: "text" },
    { prop: "aspectRatio", xmlKey: "aspectratio" },
    { prop: "formatting", xmlKey: "formatting" },
    { prop: "locking", xmlKey: "locking" },
    { prop: "shapeType", xmlKey: "shapetype" },
  ];
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.ext !== undefined) attrs.ext = { key: "v:ext", value: options.ext };
  for (const { prop, xmlKey } of boolMap) {
    const val = options?.[prop];
    if (val !== undefined) attrs[prop as string] = { key: xmlKey, value: val ? "t" : "f" };
  }
  return new BuilderElement({
    name: "o:lock",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── 3D Effects ──

export interface OExtrusionOptions {
  readonly ext?: string;
  readonly on?: boolean;
  readonly type?: "parallel" | "perspective";
  readonly render?: "solid" | "wireFrame" | "boundingBox";
  readonly viewpointOrigin?: string;
  readonly viewpoint?: string;
  readonly plane?: "xy" | "xz" | "yz";
  readonly skewAngle?: number;
  readonly skewAmount?: number;
  readonly foreDepth?: number;
  readonly backDepth?: string;
  readonly orientation?: string;
  readonly orientationAngle?: number;
  readonly color?: string;
  readonly shininess?: number;
  readonly metal?: boolean;
  readonly edge?: string;
  readonly facet?: string;
  readonly lightFace?: boolean;
  readonly brightness?: number;
  readonly lightPosition?: string;
  readonly lightLevel?: string;
  readonly lightHarsh?: boolean;
  readonly lightPosition2?: string;
  readonly lightLevel2?: string;
  readonly lightHarsh2?: boolean;
  readonly autoRotationCenter?: boolean;
  readonly rotationAngle?: string;
  readonly rotationCenter?: string;
  readonly colorMode?: "auto" | "custom";
  readonly lockRotationCenter?: boolean;
  readonly specularity?: string;
  readonly diffusity?: string;
}

/**
 * Creates an o:extrusion element (Extrusion).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="extrusion" type="CT_Extrusion"/>
 * ```
 */
export const createOExtrusion = (options?: OExtrusionOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string | number }> = {};
  if (options?.ext !== undefined) attrs.ext = { key: "v:ext", value: options.ext };
  if (options?.on !== undefined) attrs.on = { key: "on", value: options.on ? "t" : "f" };
  if (options?.type !== undefined) attrs.type = { key: "type", value: options.type };
  if (options?.render !== undefined) attrs.render = { key: "render", value: options.render };
  if (options?.viewpointOrigin !== undefined)
    attrs.viewpointOrigin = { key: "viewpointorigin", value: options.viewpointOrigin };
  if (options?.viewpoint !== undefined)
    attrs.viewpoint = { key: "viewpoint", value: options.viewpoint };
  if (options?.plane !== undefined) attrs.plane = { key: "plane", value: options.plane };
  if (options?.skewAngle !== undefined)
    attrs.skewAngle = { key: "skewangle", value: options.skewAngle };
  if (options?.skewAmount !== undefined)
    attrs.skewAmount = { key: "skewamt", value: options.skewAmount };
  if (options?.foreDepth !== undefined)
    attrs.foreDepth = { key: "foredepth", value: options.foreDepth };
  if (options?.backDepth !== undefined)
    attrs.backDepth = { key: "backdepth", value: options.backDepth };
  if (options?.orientation !== undefined)
    attrs.orientation = { key: "orientation", value: options.orientation };
  if (options?.orientationAngle !== undefined)
    attrs.orientationAngle = { key: "orientationangle", value: options.orientationAngle };
  if (options?.color !== undefined) attrs.color = { key: "color", value: options.color };
  if (options?.shininess !== undefined)
    attrs.shininess = { key: "shininess", value: options.shininess };
  if (options?.metal !== undefined)
    attrs.metal = { key: "metal", value: options.metal ? "t" : "f" };
  if (options?.edge !== undefined) attrs.edge = { key: "edge", value: options.edge };
  if (options?.facet !== undefined) attrs.facet = { key: "facet", value: options.facet };
  if (options?.lightFace !== undefined)
    attrs.lightFace = { key: "lightface", value: options.lightFace ? "t" : "f" };
  if (options?.brightness !== undefined)
    attrs.brightness = { key: "brightness", value: options.brightness };
  if (options?.lightPosition !== undefined)
    attrs.lightPosition = { key: "lightposition", value: options.lightPosition };
  if (options?.lightLevel !== undefined)
    attrs.lightLevel = { key: "lightlevel", value: options.lightLevel };
  if (options?.lightHarsh !== undefined)
    attrs.lightHarsh = { key: "lightharsh", value: options.lightHarsh ? "t" : "f" };
  if (options?.lightPosition2 !== undefined)
    attrs.lightPosition2 = { key: "lightposition2", value: options.lightPosition2 };
  if (options?.lightLevel2 !== undefined)
    attrs.lightLevel2 = { key: "lightlevel2", value: options.lightLevel2 };
  if (options?.lightHarsh2 !== undefined)
    attrs.lightHarsh2 = { key: "lightharsh2", value: options.lightHarsh2 ? "t" : "f" };
  if (options?.autoRotationCenter !== undefined)
    attrs.autoRotationCenter = {
      key: "autorotationcenter",
      value: options.autoRotationCenter ? "t" : "f",
    };
  if (options?.rotationAngle !== undefined)
    attrs.rotationAngle = { key: "rotationangle", value: options.rotationAngle };
  if (options?.rotationCenter !== undefined)
    attrs.rotationCenter = { key: "rotationcenter", value: options.rotationCenter };
  if (options?.colorMode !== undefined)
    attrs.colorMode = { key: "colormode", value: options.colorMode };
  if (options?.lockRotationCenter !== undefined)
    attrs.lockRotationCenter = {
      key: "lockrotationcenter",
      value: options.lockRotationCenter ? "t" : "f",
    };
  if (options?.specularity !== undefined)
    attrs.specularity = { key: "specularity", value: options.specularity };
  if (options?.diffusity !== undefined)
    attrs.diffusity = { key: "diffusity", value: options.diffusity };
  return new BuilderElement({
    name: "o:extrusion",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Callout ──

export interface OCalloutOptions {
  readonly ext?: string;
  readonly on?: boolean;
  readonly type?: string;
  readonly gap?: string;
  readonly angle?: "any" | "30" | "45" | "60" | "90" | "auto";
  readonly dropAuto?: boolean;
  readonly drop?: "top" | "center" | "bottom";
  readonly distance?: string;
  readonly lengthSpecified?: boolean;
  readonly length?: string;
  readonly accentBar?: boolean;
  readonly border?: boolean;
  readonly button?: boolean;
  readonly textBorder?: boolean;
  readonly minusX?: boolean;
  readonly minusY?: boolean;
}

/**
 * Creates an o:callout element (Callout).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="callout" type="CT_Callout"/>
 * ```
 */
export const createOCallout = (options?: OCalloutOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.ext !== undefined) attrs.ext = { key: "v:ext", value: options.ext };
  if (options?.on !== undefined) attrs.on = { key: "on", value: options.on ? "t" : "f" };
  if (options?.type !== undefined) attrs.type = { key: "type", value: options.type };
  if (options?.gap !== undefined) attrs.gap = { key: "gap", value: options.gap };
  if (options?.angle !== undefined) attrs.angle = { key: "angle", value: options.angle };
  if (options?.dropAuto !== undefined)
    attrs.dropAuto = { key: "dropauto", value: options.dropAuto ? "t" : "f" };
  if (options?.drop !== undefined) attrs.drop = { key: "drop", value: options.drop };
  if (options?.distance !== undefined)
    attrs.distance = { key: "distance", value: options.distance };
  if (options?.lengthSpecified !== undefined)
    attrs.lengthSpecified = { key: "lengthspecified", value: options.lengthSpecified ? "t" : "f" };
  if (options?.length !== undefined) attrs.length = { key: "length", value: options.length };
  if (options?.accentBar !== undefined)
    attrs.accentBar = { key: "accentbar", value: options.accentBar ? "t" : "f" };
  if (options?.border !== undefined)
    attrs.border = { key: "border", value: options.border ? "t" : "f" };
  if (options?.button !== undefined)
    attrs.button = { key: "button", value: options.button ? "t" : "f" };
  if (options?.textBorder !== undefined)
    attrs.textBorder = { key: "textborder", value: options.textBorder ? "t" : "f" };
  if (options?.minusX !== undefined)
    attrs.minusX = { key: "minusx", value: options.minusX ? "t" : "f" };
  if (options?.minusY !== undefined)
    attrs.minusY = { key: "minusy", value: options.minusY ? "t" : "f" };
  return new BuilderElement({
    name: "o:callout",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Skew ──

export interface OSkewOptions {
  readonly ext?: string;
  readonly on?: boolean;
  readonly offset?: string;
  readonly origin?: string;
  readonly matrix?: string;
}

/**
 * Creates an o:skew element (Skew Transform).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="skew" type="CT_Skew"/>
 * ```
 */
export const createOSkew = (options?: OSkewOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.ext !== undefined) attrs.ext = { key: "v:ext", value: options.ext };
  if (options?.on !== undefined) attrs.on = { key: "on", value: options.on ? "t" : "f" };
  if (options?.offset !== undefined) attrs.offset = { key: "offset", value: options.offset };
  if (options?.origin !== undefined) attrs.origin = { key: "origin", value: options.origin };
  if (options?.matrix !== undefined) attrs.matrix = { key: "matrix", value: options.matrix };
  return new BuilderElement({
    name: "o:skew",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Diagram ──

export interface ODiagramOptions {
  readonly ext?: string;
  readonly style?: string;
  readonly autoFormat?: boolean;
  readonly autoLayout?: boolean;
  readonly diagramScaleX?: number;
  readonly diagramScaleY?: number;
  readonly diagramFontSize?: number;
  readonly diagramBaseTextScale?: number;
  readonly diagramStyle?: number;
  readonly diagramLayout?: string;
  readonly diagramLayoutMru?: string;
  readonly diagramNodeKind?: number;
  readonly reverse?: boolean;
  readonly constrainBounds?: string;
}

/**
 * Creates an o:diagram element (Diagram).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="diagram" type="CT_Diagram"/>
 * ```
 */
export const createODiagram = (options?: ODiagramOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string | number }> = {};
  if (options?.ext !== undefined) attrs.ext = { key: "v:ext", value: options.ext };
  if (options?.style !== undefined) attrs.style = { key: "style", value: options.style };
  if (options?.autoFormat !== undefined)
    attrs.autoFormat = { key: "autoformat", value: options.autoFormat ? "t" : "f" };
  if (options?.autoLayout !== undefined)
    attrs.autoLayout = { key: "autolayout", value: options.autoLayout ? "t" : "f" };
  if (options?.diagramScaleX !== undefined)
    attrs.diagramScaleX = { key: "dgmscalex", value: options.diagramScaleX };
  if (options?.diagramScaleY !== undefined)
    attrs.diagramScaleY = { key: "dgmscaley", value: options.diagramScaleY };
  if (options?.diagramFontSize !== undefined)
    attrs.diagramFontSize = { key: "dgmfontsize", value: options.diagramFontSize };
  if (options?.diagramBaseTextScale !== undefined)
    attrs.diagramBaseTextScale = { key: "dgmbasetextscale", value: options.diagramBaseTextScale };
  if (options?.diagramStyle !== undefined)
    attrs.diagramStyle = { key: "dgmstyle", value: options.diagramStyle };
  if (options?.diagramLayout !== undefined)
    attrs.diagramLayout = { key: "dgmlayout", value: options.diagramLayout };
  if (options?.diagramLayoutMru !== undefined)
    attrs.diagramLayoutMru = { key: "dgmlayoutmru", value: options.diagramLayoutMru };
  if (options?.diagramNodeKind !== undefined)
    attrs.diagramNodeKind = { key: "dgmnodekind", value: options.diagramNodeKind };
  if (options?.reverse !== undefined)
    attrs.reverse = { key: "reverse", value: options.reverse ? "t" : "f" };
  if (options?.constrainBounds !== undefined)
    attrs.constrainBounds = { key: "constrainbounds", value: options.constrainBounds };
  return new BuilderElement({
    name: "o:diagram",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Ink ──

/**
 * Creates an o:ink element (Ink).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="ink" type="CT_Ink"/>
 * ```
 */
export interface OInkOptions {
  readonly i?: string;
  readonly annotation?: boolean;
  readonly contentType?: string;
}

export const createOInk = (options?: OInkOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.i !== undefined) attrs.i = { key: "i", value: options.i };
  if (options?.annotation !== undefined)
    attrs.annotation = { key: "annotation", value: options.annotation ? "t" : "f" };
  if (options?.contentType !== undefined)
    attrs.contentType = { key: "contentType", value: options.contentType };
  return new BuilderElement({
    name: "o:ink",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Equation XML ──

/**
 * Creates an o:equationxml element (Equation XML).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="equationxml" type="CT_EquationXml"/>
 * ```
 */
export interface OEquationXmlOptions {
  readonly contentType?: string;
}

export const createOEquationxml = (options?: OEquationXmlOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.contentType !== undefined)
    attrs.contentType = { key: "contentType", value: options.contentType };
  return new BuilderElement({
    name: "o:equationxml",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Signature Line ──

export interface OSignatureLineOptions {
  readonly ext?: string;
  readonly isSignatureLine?: boolean;
  readonly id?: string;
  readonly providerGuid?: string;
  readonly signingInstructionsSet?: boolean;
  readonly allowComments?: boolean;
  readonly showSignDate?: boolean;
  readonly signatureLine?: boolean;
  readonly suggestedSigner?: string;
  readonly suggestedSigner2?: string;
  readonly suggestedSignerEmail?: string;
  readonly signingInstructions?: string;
  readonly additionalXml?: string;
  readonly provision?: string;
  readonly button?: boolean;
  readonly sigProvUrl?: string;
}

/**
 * Creates an o:signatureline element (Signature Line).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="signatureline" type="CT_SignatureLine"/>
 * ```
 */
export const createOSignatureline = (options?: OSignatureLineOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.ext !== undefined) attrs.ext = { key: "v:ext", value: options.ext };
  if (options?.isSignatureLine !== undefined)
    attrs.isSignatureLine = { key: "issignatureline", value: options.isSignatureLine ? "t" : "f" };
  if (options?.id !== undefined) attrs.id = { key: "id", value: options.id };
  if (options?.providerGuid !== undefined)
    attrs.providerGuid = { key: "providerguid", value: options.providerGuid };
  if (options?.signingInstructionsSet !== undefined)
    attrs.signingInstructionsSet = {
      key: "signinginstructionsset",
      value: options.signingInstructionsSet ? "t" : "f",
    };
  if (options?.allowComments !== undefined)
    attrs.allowComments = { key: "allowcomments", value: options.allowComments ? "t" : "f" };
  if (options?.showSignDate !== undefined)
    attrs.showSignDate = { key: "showsigndate", value: options.showSignDate ? "t" : "f" };
  if (options?.signatureLine !== undefined)
    attrs.signatureLine = { key: "signatureline", value: options.signatureLine ? "t" : "f" };
  if (options?.suggestedSigner !== undefined)
    attrs.suggestedSigner = { key: "suggestedsigner", value: options.suggestedSigner };
  if (options?.suggestedSigner2 !== undefined)
    attrs.suggestedSigner2 = { key: "suggestedsigner2", value: options.suggestedSigner2 };
  if (options?.suggestedSignerEmail !== undefined)
    attrs.suggestedSignerEmail = {
      key: "suggestedsigneremail",
      value: options.suggestedSignerEmail,
    };
  if (options?.signingInstructions !== undefined)
    attrs.signingInstructions = { key: "signinginstructions", value: options.signingInstructions };
  if (options?.additionalXml !== undefined)
    attrs.additionalXml = { key: "addlxml", value: options.additionalXml };
  if (options?.provision !== undefined)
    attrs.provision = { key: "provid", value: options.provision };
  if (options?.sigProvUrl !== undefined)
    attrs.sigProvUrl = { key: "sigprovurl", value: options.sigProvUrl };
  if (options?.button !== undefined)
    attrs.button = { key: "button", value: options.button ? "t" : "f" };
  return new BuilderElement({
    name: "o:signatureline",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Complex ──

/**
 * Creates an o:complex element (Complex).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="complex" type="CT_Complex"/>
 * ```
 */
export const createOComplex = (): XmlComponent => new BuilderElement({ name: "o:complex" });

// ── Fill ──

export interface OFillOptions {
  readonly ext?: string;
  readonly type?: string;
}

/**
 * Creates an o:fill element (Office Fill).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="fill" type="CT_Fill"/>
 * ```
 */
export const createOFill = (options?: OFillOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.ext !== undefined) attrs.ext = { key: "v:ext", value: options.ext };
  if (options?.type !== undefined) attrs.type = { key: "type", value: options.type };
  return new BuilderElement({
    name: "o:fill",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Clip Path ──

/**
 * Creates an o:clippath element (Clip Path).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="clippath" type="CT_ClipPath"/>
 * ```
 */
export const createOClippath = (): XmlComponent => new BuilderElement({ name: "o:clippath" });

// ── Stroke Children (left, top, right, bottom, column) ──

export interface OStrokeChildOptions {
  readonly ext?: string;
  readonly on?: boolean;
  readonly weight?: string;
  readonly color?: string;
  readonly dashStyle?: string;
  readonly miterLimit?: number;
  readonly lineStyle?: "single" | "thinThin" | "thinThick" | "thickThin" | "thickBetweenThin";
  readonly insetPen?: boolean;
  readonly endCap?: "flat" | "square" | "round";
  readonly joinStyle?: "round" | "bevel" | "miter";
  readonly color2?: string;
  readonly opacity?: string;
  readonly fillType?: "solid" | "tile" | "pattern" | "frame";
  readonly src?: string;
  readonly imageAspect?: "ignore" | "atMost" | "atLeast";
  readonly imageSize?: string;
  readonly imageAlignShape?: boolean;
  readonly startArrow?: "none" | "block" | "classic" | "oval" | "diamond" | "open";
  readonly startArrowWidth?: "narrow" | "medium" | "wide";
  readonly startArrowLength?: "short" | "medium" | "long";
  readonly endArrow?: "none" | "block" | "classic" | "oval" | "diamond" | "open";
  readonly endArrowWidth?: "narrow" | "medium" | "wide";
  readonly endArrowLength?: "short" | "medium" | "long";
  readonly forceDash?: boolean;
}

function buildStrokeChildAttrs(
  options?: OStrokeChildOptions,
): Record<string, { key: string; value: string | number }> | undefined {
  if (!options) return undefined;
  const attrs: Record<string, { key: string; value: string | number }> = {};
  if (options.ext !== undefined) attrs.ext = { key: "v:ext", value: options.ext };
  if (options.on !== undefined) attrs.on = { key: "on", value: options.on ? "t" : "f" };
  if (options.weight !== undefined) attrs.weight = { key: "weight", value: options.weight };
  if (options.color !== undefined) attrs.color = { key: "color", value: options.color };
  if (options.dashStyle !== undefined)
    attrs.dashStyle = { key: "dashstyle", value: options.dashStyle };
  if (options.miterLimit !== undefined)
    attrs.miterLimit = { key: "miterlimit", value: options.miterLimit };
  if (options.lineStyle !== undefined)
    attrs.lineStyle = { key: "linestyle", value: options.lineStyle };
  if (options.insetPen !== undefined)
    attrs.insetPen = { key: "insetpen", value: options.insetPen ? "t" : "f" };
  if (options.endCap !== undefined) attrs.endCap = { key: "endcap", value: options.endCap };
  if (options.joinStyle !== undefined)
    attrs.joinStyle = { key: "joinstyle", value: options.joinStyle };
  if (options.color2 !== undefined) attrs.color2 = { key: "color2", value: options.color2 };
  if (options.opacity !== undefined) attrs.opacity = { key: "opacity", value: options.opacity };
  if (options.fillType !== undefined) attrs.fillType = { key: "filltype", value: options.fillType };
  if (options.src !== undefined) attrs.src = { key: "src", value: options.src };
  if (options.imageAspect !== undefined)
    attrs.imageAspect = { key: "imageaspect", value: options.imageAspect };
  if (options.imageSize !== undefined)
    attrs.imageSize = { key: "imagesize", value: options.imageSize };
  if (options.imageAlignShape !== undefined)
    attrs.imageAlignShape = {
      key: "imagealignshape",
      value: options.imageAlignShape ? "t" : "f",
    };
  if (options.startArrow !== undefined)
    attrs.startArrow = { key: "startarrow", value: options.startArrow };
  if (options.startArrowWidth !== undefined)
    attrs.startArrowWidth = { key: "startarrowwidth", value: options.startArrowWidth };
  if (options.startArrowLength !== undefined)
    attrs.startArrowLength = { key: "startarrowlength", value: options.startArrowLength };
  if (options.endArrow !== undefined) attrs.endArrow = { key: "endarrow", value: options.endArrow };
  if (options.endArrowWidth !== undefined)
    attrs.endArrowWidth = { key: "endarrowwidth", value: options.endArrowWidth };
  if (options.endArrowLength !== undefined)
    attrs.endArrowLength = { key: "endarrowlength", value: options.endArrowLength };
  if (options.forceDash !== undefined)
    attrs.forceDash = { key: "forcedash", value: options.forceDash ? "t" : "f" };
  return Object.keys(attrs).length > 0 ? attrs : undefined;
}

/**
 * Creates an o:left element (Left Stroke).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="left" type="CT_StrokeChild"/>
 * ```
 */
export const createOLeft = (options?: OStrokeChildOptions): XmlComponent =>
  new BuilderElement({ name: "o:left", attributes: buildStrokeChildAttrs(options) });

/**
 * Creates an o:top element (Top Stroke).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="top" type="CT_StrokeChild"/>
 * ```
 */
export const createOTop = (options?: OStrokeChildOptions): XmlComponent =>
  new BuilderElement({ name: "o:top", attributes: buildStrokeChildAttrs(options) });

/**
 * Creates an o:right element (Right Stroke).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="right" type="CT_StrokeChild"/>
 * ```
 */
export const createORight = (options?: OStrokeChildOptions): XmlComponent =>
  new BuilderElement({ name: "o:right", attributes: buildStrokeChildAttrs(options) });

/**
 * Creates an o:bottom element (Bottom Stroke).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="bottom" type="CT_StrokeChild"/>
 * ```
 */
export const createOBottom = (options?: OStrokeChildOptions): XmlComponent =>
  new BuilderElement({ name: "o:bottom", attributes: buildStrokeChildAttrs(options) });

/**
 * Creates an o:column element (Column Stroke).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="column" type="CT_StrokeChild"/>
 * ```
 */
export const createOColumn = (options?: OStrokeChildOptions): XmlComponent =>
  new BuilderElement({ name: "o:column", attributes: buildStrokeChildAttrs(options) });

// ── Regroup Table ──

/**
 * Creates an o:regrouptable element (Regroup Table).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="regrouptable" type="CT_RegroupTable"/>
 * ```
 */
export const createORegrouptable = (): XmlComponent =>
  new BuilderElement({ name: "o:regrouptable" });

/**
 * Creates an o:entry element (Regroup Table Entry).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="entry" type="CT_Entry"/>
 * ```
 */
export interface OEntryOptions {
  readonly newType?: number;
  readonly oldType?: number;
}

export const createOEntry = (options?: OEntryOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: number }> = {};
  if (options?.newType !== undefined) attrs.newType = { key: "new", value: options.newType };
  if (options?.oldType !== undefined) attrs.oldType = { key: "old", value: options.oldType };
  return new BuilderElement({
    name: "o:entry",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Rules ──

/**
 * Creates an o:rules element (Rules).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="rules" type="CT_Rules"/>
 * ```
 */
export const createORules = (): XmlComponent => new BuilderElement({ name: "o:rules" });

/**
 * Creates an o:r element (Rule).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="r" type="CT_R"/>
 * ```
 */
export interface ORuleOptions {
  readonly id?: string;
  readonly type?: string;
  readonly how?: string;
  readonly idref?: string;
}

export const createOR = (options?: ORuleOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.id !== undefined) attrs.id = { key: "id", value: options.id };
  if (options?.type !== undefined) attrs.type = { key: "type", value: options.type };
  if (options?.how !== undefined) attrs.how = { key: "how", value: options.how };
  if (options?.idref !== undefined) attrs.idref = { key: "idref", value: options.idref };
  return new BuilderElement({
    name: "o:r",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

/**
 * Creates an o:proxy element (Rule Proxy).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="proxy" type="CT_Proxy"/>
 * ```
 */
export interface OProxyOptions {
  readonly start?: boolean;
  readonly end?: boolean;
  readonly idref?: string;
  readonly connectLoc?: number;
}

export const createOProxy = (options?: OProxyOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string | number }> = {};
  if (options?.start !== undefined)
    attrs.start = { key: "start", value: options.start ? "t" : "f" };
  if (options?.end !== undefined) attrs.end = { key: "end", value: options.end ? "t" : "f" };
  if (options?.idref !== undefined) attrs.idref = { key: "idref", value: options.idref };
  if (options?.connectLoc !== undefined)
    attrs.connectLoc = { key: "connectloc", value: options.connectLoc };
  return new BuilderElement({
    name: "o:proxy",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── Relation Table ──

/**
 * Creates an o:relationtable element (Relation Table).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="relationtable" type="CT_RelationTable"/>
 * ```
 */
export const createORelationtable = (): XmlComponent =>
  new BuilderElement({ name: "o:relationtable" });

/**
 * Creates an o:rel element (Relation).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="rel" type="CT_Rel"/>
 * ```
 */
export interface ORelOptions {
  readonly idSrc?: string;
  readonly idDest?: string;
  readonly idCntr?: string;
}

export const createORel = (options?: ORelOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.idSrc !== undefined) attrs.idSrc = { key: "idsrc", value: options.idSrc };
  if (options?.idDest !== undefined) attrs.idDest = { key: "iddest", value: options.idDest };
  if (options?.idCntr !== undefined) attrs.idCntr = { key: "idcntr", value: options.idCntr };
  return new BuilderElement({
    name: "o:rel",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
  });
};

// ── OLE Object Children ──

/**
 * Creates an o:LinkType element (OLE Link Type).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="LinkType" type="ST_OLELinkType"/>
 * ```
 */
export const createOLinkType = (): XmlComponent => new BuilderElement({ name: "o:LinkType" });

/**
 * Creates an o:LockedField element (OLE Locked Field).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="LockedField" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createOLockedField = (): XmlComponent => new BuilderElement({ name: "o:LockedField" });

/**
 * Creates an o:FieldCodes element (OLE Field Codes).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="FieldCodes" type="xsd:string"/>
 * ```
 */
export const createOFieldCodes = (): XmlComponent => new BuilderElement({ name: "o:FieldCodes" });

// ── Global o: Shape Attributes ──
// These attributes (from vml-officeDrawing.xsd global declarations)
// can appear on any VML shape element (v:shape, v:shapetype, v:group, etc.)
// as o:-prefixed XML attributes.

export interface OGlobalShapeAttributes {
  readonly spid?: string;
  readonly shapeTypeIndex?: number;
  readonly oneDimensional?: boolean;
  readonly regroupId?: number;
  readonly doubleClickNotify?: boolean;
  readonly userHidden?: boolean;
  readonly bullet?: boolean;
  readonly horizontalRule?: boolean;
  readonly horizontalRuleStandard?: boolean;
  readonly horizontalRuleNoShade?: boolean;
  readonly horizontalRulePercent?: number;
  readonly horizontalRuleAlign?: "left" | "center" | "right";
  readonly allowOverlap?: boolean;
  readonly userDrawn?: boolean;
  readonly borderTopColor?: string;
  readonly borderLeftColor?: string;
  readonly borderBottomColor?: string;
  readonly borderRightColor?: string;
  readonly connectorType?: "none" | "straight" | "elbow" | "curved";
  readonly blackWhiteMode?:
    | "color"
    | "auto"
    | "gray"
    | "lightGray"
    | "inverseGray"
    | "grayOutline"
    | "highContrast"
    | "black"
    | "white"
    | "hide"
    | "undrawn";
  readonly blackWhiteNormal?: string;
  readonly blackWhitePure?: string;
  readonly forceDash?: boolean;
  readonly oleIcon?: boolean;
  readonly ole?: boolean;
  readonly preferRelative?: boolean;
  readonly clipToWrap?: boolean;
  readonly clip?: boolean;
  readonly master?: string;
  readonly extrusionOk?: boolean;
  readonly altHref?: string;
  readonly oleId?: number;
  readonly detectMouseClick?: boolean;
  readonly targetScreenSize?:
    | "544x376"
    | "640x480"
    | "720x512"
    | "800x600"
    | "1024x768"
    | "1152x862"
    | "1280x1024";
  readonly insetMode?: "auto" | "custom";
  readonly connectType?: "none" | "rect" | "segments" | "custom";
  readonly connectLocs?: string;
  readonly connectAngles?: string;
  readonly tableProperties?: string;
  readonly tableLimits?: string;
  readonly opacity2?: string;
  readonly relId?: string;
}

/**
 * Builds o:-prefixed global shape attributes for VML shape elements.
 * Returns attribute entries suitable for merging into BuilderElement attributes.
 */
export const buildOGlobalShapeAttrs = (
  options?: OGlobalShapeAttributes,
): Record<string, { key: string; value: string | number }> => {
  if (!options) return {};
  const attrs: Record<string, { key: string; value: string | number }> = {};
  if (options.spid !== undefined) attrs.spid = { key: "spid", value: options.spid };
  if (options.shapeTypeIndex !== undefined)
    attrs.shapeTypeIndex = { key: "spt", value: options.shapeTypeIndex };
  if (options.oneDimensional !== undefined)
    attrs.oneDimensional = { key: "oned", value: options.oneDimensional ? "t" : "f" };
  if (options.regroupId !== undefined)
    attrs.regroupId = { key: "regroupid", value: options.regroupId };
  if (options.doubleClickNotify !== undefined)
    attrs.doubleClickNotify = {
      key: "doubleclicknotify",
      value: options.doubleClickNotify ? "t" : "f",
    };
  if (options.userHidden !== undefined)
    attrs.userHidden = { key: "userhidden", value: options.userHidden ? "t" : "f" };
  if (options.bullet !== undefined)
    attrs.bullet = { key: "bullet", value: options.bullet ? "t" : "f" };
  if (options.horizontalRule !== undefined)
    attrs.horizontalRule = { key: "hr", value: options.horizontalRule ? "t" : "f" };
  if (options.horizontalRuleStandard !== undefined)
    attrs.horizontalRuleStandard = {
      key: "hrstd",
      value: options.horizontalRuleStandard ? "t" : "f",
    };
  if (options.horizontalRuleNoShade !== undefined)
    attrs.horizontalRuleNoShade = {
      key: "hrnoshade",
      value: options.horizontalRuleNoShade ? "t" : "f",
    };
  if (options.horizontalRulePercent !== undefined)
    attrs.horizontalRulePercent = { key: "hrpct", value: options.horizontalRulePercent };
  if (options.horizontalRuleAlign !== undefined)
    attrs.horizontalRuleAlign = { key: "hralign", value: options.horizontalRuleAlign };
  if (options.allowOverlap !== undefined)
    attrs.allowOverlap = { key: "allowoverlap", value: options.allowOverlap ? "t" : "f" };
  if (options.userDrawn !== undefined)
    attrs.userDrawn = { key: "userdrawn", value: options.userDrawn ? "t" : "f" };
  if (options.borderTopColor !== undefined)
    attrs.borderTopColor = { key: "bordertopcolor", value: options.borderTopColor };
  if (options.borderLeftColor !== undefined)
    attrs.borderLeftColor = { key: "borderleftcolor", value: options.borderLeftColor };
  if (options.borderBottomColor !== undefined)
    attrs.borderBottomColor = { key: "borderbottomcolor", value: options.borderBottomColor };
  if (options.borderRightColor !== undefined)
    attrs.borderRightColor = { key: "borderrightcolor", value: options.borderRightColor };
  if (options.connectorType !== undefined)
    attrs.connectorType = { key: "connectortype", value: options.connectorType };
  if (options.blackWhiteMode !== undefined)
    attrs.blackWhiteMode = { key: "bwmode", value: options.blackWhiteMode };
  if (options.blackWhiteNormal !== undefined)
    attrs.blackWhiteNormal = { key: "bwnormal", value: options.blackWhiteNormal };
  if (options.blackWhitePure !== undefined)
    attrs.blackWhitePure = { key: "bwpure", value: options.blackWhitePure };
  if (options.forceDash !== undefined)
    attrs.forceDash = { key: "forcedash", value: options.forceDash ? "t" : "f" };
  if (options.oleIcon !== undefined)
    attrs.oleIcon = { key: "oleicon", value: options.oleIcon ? "t" : "f" };
  if (options.ole !== undefined) attrs.ole = { key: "ole", value: options.ole ? "t" : "f" };
  if (options.preferRelative !== undefined)
    attrs.preferRelative = { key: "preferrelative", value: options.preferRelative ? "t" : "f" };
  if (options.clipToWrap !== undefined)
    attrs.clipToWrap = { key: "cliptowrap", value: options.clipToWrap ? "t" : "f" };
  if (options.clip !== undefined) attrs.clip = { key: "clip", value: options.clip ? "t" : "f" };
  if (options.master !== undefined) attrs.master = { key: "master", value: options.master };
  if (options.extrusionOk !== undefined)
    attrs.extrusionOk = { key: "extrusionok", value: options.extrusionOk ? "t" : "f" };
  if (options.altHref !== undefined) attrs.altHref = { key: "althref", value: options.altHref };
  if (options.oleId !== undefined) attrs.oleId = { key: "oleid", value: options.oleId };
  if (options.detectMouseClick !== undefined)
    attrs.detectMouseClick = {
      key: "detectmouseclick",
      value: options.detectMouseClick ? "t" : "f",
    };
  if (options.targetScreenSize !== undefined)
    attrs.targetScreenSize = { key: "targetscreensize", value: options.targetScreenSize };
  if (options.insetMode !== undefined)
    attrs.insetMode = { key: "insetmode", value: options.insetMode };
  if (options.connectType !== undefined)
    attrs.connectType = { key: "connecttype", value: options.connectType };
  if (options.connectLocs !== undefined)
    attrs.connectLocs = { key: "connectlocs", value: options.connectLocs };
  if (options.connectAngles !== undefined)
    attrs.connectAngles = { key: "connectangles", value: options.connectAngles };
  if (options.tableProperties !== undefined)
    attrs.tableProperties = { key: "tableproperties", value: options.tableProperties };
  if (options.tableLimits !== undefined)
    attrs.tableLimits = { key: "tablelimits", value: options.tableLimits };
  if (options.opacity2 !== undefined) attrs.opacity2 = { key: "opacity2", value: options.opacity2 };
  if (options.relId !== undefined) attrs.relId = { key: "relid", value: options.relId };
  return attrs;
};
