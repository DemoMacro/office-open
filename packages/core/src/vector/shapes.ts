/**
 * VML shape elements — CT_Shape, CT_Shapetype, CT_Group, CT_Background and
 * the eight basic shapes (arc, curve, image, line, oval, polyline, rect,
 * roundrect).
 *
 * All shape-ish elements share the AG_AllCoreAttributes / AG_AllShapeAttributes
 * vocabulary (see `attributes.ts`) plus the EG_ShapeElements child set (see
 * `shape-elements/`); the per-element extras are small. Child order follows
 * the EG_ShapeElements declaration order — the XSD content model is an
 * unordered repeating choice, and that order matches what Office emits.
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml, stringifyElement } from "@office-open/xml";

import {
  stringifyVmlAttributes,
  parseVmlAttributes,
  VML_CORE_ATTRS,
  VML_SHAPE_ATTRS,
  VML_OFFICE_CORE_ATTRS,
  VML_OFFICE_SHAPE_ATTRS,
  type VmlAttrSpec,
  type VmlColor,
  type VmlCoreAttributes,
  type VmlOfficeCoreAttributes,
  type VmlOfficeShapeAttributes,
  type VmlShapeAttributes,
  type VmlTrueFalse,
  type VmlBlackWhiteMode,
  type VmlScreenSize,
} from "./attributes";
import {
  stringifyVmlClientData,
  parseVmlClientData,
  type VmlClientDataOptions,
} from "./shape-elements/client-data";
import { stringifyVmlFill, parseVmlFill, type VmlFillOptions } from "./shape-elements/fill";
import {
  stringifyVmlFormulas,
  parseVmlFormulas,
  type VmlFormulasOptions,
} from "./shape-elements/formulas";
import {
  stringifyVmlHandles,
  parseVmlHandles,
  type VmlHandlesOptions,
} from "./shape-elements/handles";
import {
  stringifyVmlImageData,
  parseVmlImageData,
  type VmlImageDataOptions,
} from "./shape-elements/imagedata";
import {
  stringifyVmlSkew,
  parseVmlSkew,
  stringifyVmlExtrusion,
  parseVmlExtrusion,
  stringifyVmlCallout,
  parseVmlCallout,
  stringifyVmlLock,
  parseVmlLock,
  stringifyVmlClipPath,
  parseVmlClipPath,
  stringifyVmlSignatureLine,
  parseVmlSignatureLine,
  stringifyVmlInk,
  parseVmlInk,
  stringifyVmlEquationXml,
  parseVmlEquationXml,
  stringifyVmlDiagram,
  parseVmlDiagram,
  stringifyVmlComplex,
  parseVmlComplex,
  type VmlSkewOptions,
  type VmlExtrusionOptions,
  type VmlCalloutOptions,
  type VmlLockOptions,
  type VmlClipPathOptions,
  type VmlSignatureLineOptions,
  type VmlInkOptions,
  type VmlEquationXmlOptions,
  type VmlDiagramOptions,
  type VmlComplexOptions,
} from "./shape-elements/office-elements";
import { stringifyVmlPath, parseVmlPath, type VmlPathOptions } from "./shape-elements/path";
import {
  stringifyVmlIsComment,
  parseVmlIsComment,
  stringifyVmlTextData,
  parseVmlTextData,
  type VmlIsCommentOptions,
  type VmlTextDataOptions,
} from "./shape-elements/presentation-elements";
import { stringifyVmlShadow, parseVmlShadow, type VmlShadowOptions } from "./shape-elements/shadow";
import { stringifyVmlStroke, parseVmlStroke, type VmlStrokeOptions } from "./shape-elements/stroke";
import {
  stringifyVmlTextbox,
  parseVmlTextbox,
  type VmlTextboxOptions,
} from "./shape-elements/textbox";
import {
  stringifyVmlTextPath,
  parseVmlTextPath,
  type VmlTextPathOptions,
} from "./shape-elements/textpath";
import {
  stringifyVmlWrap,
  parseVmlWrap,
  stringifyVmlAnchorLock,
  parseVmlAnchorLock,
  stringifyVmlBorder,
  parseVmlBorder,
  type VmlWrapOptions,
  type VmlAnchorLockOptions,
  type VmlBorderOptions,
} from "./shape-elements/word-elements";
import { stringifyVmlStyle, parseVmlStyle, parseVmlShapeStyle, type VmlShapeStyle } from "./style";

// ── Shape id allocation ──

/** Next spid in the `_x0000_s` segment (Word hands out shape ids from 1025). */
let nextSpid = 1024;

/**
 * Allocate the next VML shape id (`_x0000_s1025`, `_x0000_s1026`, …).
 * Ids only need to be unique within a document; the counter is process-global
 * so concurrent generations never collide.
 */
export function nextVmlShapeId(): string {
  return `_x0000_s${++nextSpid}`;
}

// ── Shared shape options base ──

/**
 * Fields shared by every shape-ish element: the v: + o: attribute vocabulary
 * (AG_AllCoreAttributes / AG_AllShapeAttributes) and the EG_ShapeElements
 * child set.
 */
export interface VmlBaseShapeFields
  extends VmlCoreAttributes, VmlOfficeCoreAttributes, VmlShapeAttributes, VmlOfficeShapeAttributes {
  /** AG_Path attribute — path data overriding the referenced shapetype's geometry. */
  path?: string;
  /** v:path child element — a separate XSD slot from the path attribute. */
  pathElement?: VmlPathOptions;
  formulas?: VmlFormulasOptions;
  handles?: VmlHandlesOptions;
  fill?: VmlFillOptions;
  stroke?: VmlStrokeOptions;
  shadow?: VmlShadowOptions;
  textbox?: VmlTextboxOptions;
  textpath?: VmlTextPathOptions;
  imagedata?: VmlImageDataOptions;
  // o: members of EG_ShapeElements
  skew?: VmlSkewOptions;
  extrusion?: VmlExtrusionOptions;
  callout?: VmlCalloutOptions;
  lock?: VmlLockOptions;
  clippath?: VmlClipPathOptions;
  signatureline?: VmlSignatureLineOptions;
  // w10: / x: / pvml: members of EG_ShapeElements
  wrap?: VmlWrapOptions;
  anchorlock?: VmlAnchorLockOptions;
  bordertop?: VmlBorderOptions;
  borderbottom?: VmlBorderOptions;
  borderleft?: VmlBorderOptions;
  borderright?: VmlBorderOptions;
  clientData?: VmlClientDataOptions;
  textdata?: VmlTextDataOptions;
  /**
   * Verbatim XML of unmodeled children (e.g. Word's re-prefixed
   * `wvml:bordertop` border elements), re-emitted after the modeled children.
   */
  rawChildrenXml?: string;
}

/** ST_EditAs. */
export type VmlEditAs =
  | "canvas"
  | "orgchart"
  | "radial"
  | "cycle"
  | "stacked"
  | "venn"
  | "bullseye";

// ── Options types ──

/** v:shape options (CT_Shape). */
export interface VmlShapeOptions extends VmlBaseShapeFields {
  /** Shapetype reference, e.g. "#_x0000_t202". */
  type?: string;
  /** AG_Adj — adjustment values, e.g. "1,500". */
  adj?: string;
  /** o:gfxdata — base64 shape geometry payload (UI cache). */
  gfxdata?: string;
  /** `@equationxml` — alternative-equation string attribute. */
  equationxml?: string;
  /** o:ink child — Tablet PC ink annotation. */
  ink?: VmlInkOptions;
  /** o:equationxml child — alternate math content as verbatim XML. */
  equationxmlElement?: VmlEquationXmlOptions;
  /** pvml:iscomment child — PowerPoint ink-comment marker. */
  iscomment?: VmlIsCommentOptions;
}

/** v:shapetype options (CT_Shapetype). */
export interface VmlShapetypeOptions extends VmlBaseShapeFields {
  adj?: string;
  /** o:master — the shapetype this one derives from. */
  master?: string;
  /** o:complex — the trailing extension slot. */
  complex?: VmlComplexOptions;
}

/** One shape element — single-key wrapper, like the format packages' child
 *  unions. Hosts v:group children and w:pict content (both are xsd:any over
 *  the shape vocabulary). */
export type VmlShapeChild =
  | { shape: VmlShapeOptions }
  | { shapetype: VmlShapetypeOptions }
  | { group: VmlGroupOptions }
  | { arc: VmlArcOptions }
  | { curve: VmlCurveOptions }
  | { image: VmlImageOptions }
  | { line: VmlLineOptions }
  | { oval: VmlOvalOptions }
  | { polyline: VmlPolylineOptions }
  | { rect: VmlRectOptions }
  | { roundrect: VmlRoundRectOptions };

/** v:group options (CT_Group) — full attr vocabulary + fill subset + nested shapes. */
export interface VmlGroupOptions extends VmlBaseShapeFields {
  filled?: VmlTrueFalse;
  fillcolor?: VmlColor;
  editas?: VmlEditAs;
  /** o:tableproperties / o:tablelimits — legacy diagram table metadata. */
  tableproperties?: string;
  tablelimits?: string;
  /** o:diagram child — legacy org-chart/diagram payload. */
  diagram?: VmlDiagramOptions;
  children?: VmlShapeChild[];
}

/** v:background options (CT_Background). */
export interface VmlBackgroundOptions {
  id?: string;
  filled?: VmlTrueFalse;
  fillcolor?: VmlColor;
  bwmode?: VmlBlackWhiteMode;
  bwpure?: VmlBlackWhiteMode;
  bwnormal?: VmlBlackWhiteMode;
  targetscreensize?: VmlScreenSize;
  fill?: VmlFillOptions;
}

/** v:arc options (CT_Arc). */
export interface VmlArcOptions extends VmlBaseShapeFields {
  /** Start angle in degrees. */
  startAngle?: number;
  /** End angle in degrees. */
  endAngle?: number;
}

/** v:curve options (CT_Curve). */
export interface VmlCurveOptions extends VmlBaseShapeFields {
  from?: string;
  control1?: string;
  control2?: string;
  to?: string;
}

/** v:image options (CT_Image) — AG_ImageAttributes rides along. */
export interface VmlImageOptions extends VmlBaseShapeFields {
  src?: string;
  cropleft?: string;
  croptop?: string;
  cropright?: string;
  cropbottom?: string;
  gain?: string;
  blacklevel?: string;
  gamma?: string;
  grayscale?: VmlTrueFalse;
  bilevel?: VmlTrueFalse;
}

/** v:line options (CT_Line). */
export interface VmlLineOptions extends VmlBaseShapeFields {
  from?: string;
  to?: string;
}

/** v:oval options (CT_Oval). */
export interface VmlOvalOptions extends VmlBaseShapeFields {}

/** v:polyline options (CT_PolyLine). */
export interface VmlPolylineOptions extends VmlBaseShapeFields {
  points?: string;
  /** o:ink child — ink annotation riding on the polyline. */
  ink?: VmlInkOptions;
}

/** v:rect options (CT_Rect). */
export interface VmlRectOptions extends VmlBaseShapeFields {}

/** v:roundrect options (CT_RoundRect). */
export interface VmlRoundRectOptions extends VmlBaseShapeFields {
  arcsize?: string;
}

// ── Shared (de)serialization ──

const GROUP_FILL_ATTRS: readonly VmlAttrSpec[] = [
  { field: "filled", attr: "filled", kind: "trueFalse" },
  { field: "fillcolor", attr: "fillcolor", kind: "string" },
];

const ARC_ATTRS: readonly VmlAttrSpec[] = [
  { field: "startAngle", attr: "startAngle", kind: "number" },
  { field: "endAngle", attr: "endAngle", kind: "number" },
];

const CURVE_ATTRS: readonly VmlAttrSpec[] = [
  { field: "from", attr: "from", kind: "string" },
  { field: "control1", attr: "control1", kind: "string" },
  { field: "control2", attr: "control2", kind: "string" },
  { field: "to", attr: "to", kind: "string" },
];

const IMAGE_ATTRS: readonly VmlAttrSpec[] = [
  { field: "src", attr: "src", kind: "string" },
  { field: "cropleft", attr: "cropleft", kind: "string" },
  { field: "croptop", attr: "croptop", kind: "string" },
  { field: "cropright", attr: "cropright", kind: "string" },
  { field: "cropbottom", attr: "cropbottom", kind: "string" },
  { field: "gain", attr: "gain", kind: "string" },
  { field: "blacklevel", attr: "blacklevel", kind: "string" },
  { field: "gamma", attr: "gamma", kind: "string" },
  { field: "grayscale", attr: "grayscale", kind: "trueFalse" },
  { field: "bilevel", attr: "bilevel", kind: "trueFalse" },
];

const LINE_ATTRS: readonly VmlAttrSpec[] = [
  { field: "from", attr: "from", kind: "string" },
  { field: "to", attr: "to", kind: "string" },
];

const POLYLINE_ATTRS: readonly VmlAttrSpec[] = [
  { field: "points", attr: "points", kind: "string" },
];

const ROUNDRECT_ATTRS: readonly VmlAttrSpec[] = [
  { field: "arcsize", attr: "arcsize", kind: "string" },
];

/** The full AG_AllCoreAttributes + AG_AllShapeAttributes table. */
const ALL_SHAPE_ATTRS: readonly VmlAttrSpec[] = [
  ...VML_CORE_ATTRS,
  ...VML_OFFICE_CORE_ATTRS,
  ...VML_SHAPE_ATTRS,
  ...VML_OFFICE_SHAPE_ATTRS,
];

/** The AG_Path attribute — always emitted after the shared vocabulary. */
const PATH_ATTR: readonly VmlAttrSpec[] = [{ field: "path", attr: "path", kind: "string" }];

/** Serialize the shared attribute vocabulary (style special-cased) plus extras. */
function stringifyShapeAttrs(
  opts: Record<string, unknown>,
  extraSpecs: readonly VmlAttrSpec[] = [],
): string {
  let attrStr = stringifyVmlAttributes(opts, [...ALL_SHAPE_ATTRS, ...extraSpecs, ...PATH_ATTR]);
  const style = opts.style as VmlShapeStyle | undefined;
  if (style !== undefined) {
    attrStr += ` style="${escapeXml(stringifyVmlStyle(style))}"`;
  }
  return attrStr;
}

/** Serialize the EG_ShapeElements children present on `opts`, in declaration order. */
function stringifyShapeElements(opts: VmlBaseShapeFields): string {
  const parts: string[] = [];
  if (opts.pathElement !== undefined) parts.push(stringifyVmlPath(opts.pathElement));
  if (opts.formulas !== undefined) parts.push(stringifyVmlFormulas(opts.formulas));
  if (opts.handles !== undefined) parts.push(stringifyVmlHandles(opts.handles));
  if (opts.fill !== undefined) parts.push(stringifyVmlFill(opts.fill));
  if (opts.stroke !== undefined) parts.push(stringifyVmlStroke(opts.stroke));
  if (opts.shadow !== undefined) parts.push(stringifyVmlShadow(opts.shadow));
  if (opts.textbox !== undefined) parts.push(stringifyVmlTextbox(opts.textbox));
  if (opts.textpath !== undefined) parts.push(stringifyVmlTextPath(opts.textpath));
  if (opts.imagedata !== undefined) parts.push(stringifyVmlImageData(opts.imagedata));
  if (opts.skew !== undefined) parts.push(stringifyVmlSkew(opts.skew));
  if (opts.extrusion !== undefined) parts.push(stringifyVmlExtrusion(opts.extrusion));
  if (opts.callout !== undefined) parts.push(stringifyVmlCallout(opts.callout));
  if (opts.lock !== undefined) parts.push(stringifyVmlLock(opts.lock));
  if (opts.clippath !== undefined) parts.push(stringifyVmlClipPath(opts.clippath));
  if (opts.signatureline !== undefined) {
    parts.push(stringifyVmlSignatureLine(opts.signatureline));
  }
  if (opts.wrap !== undefined) parts.push(stringifyVmlWrap(opts.wrap));
  if (opts.anchorlock !== undefined) parts.push(stringifyVmlAnchorLock(opts.anchorlock));
  if (opts.bordertop !== undefined) parts.push(stringifyVmlBorder("w10:bordertop", opts.bordertop));
  if (opts.borderbottom !== undefined) {
    parts.push(stringifyVmlBorder("w10:borderbottom", opts.borderbottom));
  }
  if (opts.borderleft !== undefined) {
    parts.push(stringifyVmlBorder("w10:borderleft", opts.borderleft));
  }
  if (opts.borderright !== undefined) {
    parts.push(stringifyVmlBorder("w10:borderright", opts.borderright));
  }
  if (opts.clientData !== undefined) parts.push(stringifyVmlClientData(opts.clientData));
  if (opts.textdata !== undefined) parts.push(stringifyVmlTextData(opts.textdata));
  if (opts.rawChildrenXml) parts.push(opts.rawChildrenXml);
  return parts.join("");
}

/** Parse the shared attribute vocabulary plus extras onto `out`. */
function parseShapeAttrs(
  el: XmlElement,
  out: Record<string, unknown>,
  extraSpecs: readonly VmlAttrSpec[] = [],
): void {
  parseVmlAttributes(el, [...ALL_SHAPE_ATTRS, ...extraSpecs, ...PATH_ATTR], out);
  if (el.attributes?.style !== undefined) {
    out.style = parseVmlShapeStyle(parseVmlStyle(String(el.attributes.style)));
  }
}

/** Child shape element names — collected by the group child loop, never by
 * the EG_ShapeElements switch (their slots are group members, not shared
 * shape-element children). */
const NESTED_SHAPE_NAMES = new Set([
  "v:shape",
  "v:shapetype",
  "v:group",
  "v:rect",
  "v:roundrect",
  "v:oval",
  "v:line",
  "v:image",
  "v:arc",
  "v:curve",
  "v:polyline",
]);

/** Parse the EG_ShapeElements children from `el` onto `out`. */
function parseShapeElements(el: XmlElement, out: Record<string, unknown>): void {
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    switch (child.name) {
      case "v:path":
        out.pathElement = parseVmlPath(child);
        break;
      case "v:formulas":
        out.formulas = parseVmlFormulas(child);
        break;
      case "v:handles":
        out.handles = parseVmlHandles(child);
        break;
      case "v:fill":
        out.fill = parseVmlFill(child);
        break;
      case "v:stroke":
        out.stroke = parseVmlStroke(child);
        break;
      case "v:shadow":
        out.shadow = parseVmlShadow(child);
        break;
      case "v:textbox":
        out.textbox = parseVmlTextbox(child);
        break;
      case "v:textpath":
        out.textpath = parseVmlTextPath(child);
        break;
      case "v:imagedata":
        out.imagedata = parseVmlImageData(child);
        break;
      case "o:skew":
        out.skew = parseVmlSkew(child);
        break;
      case "o:extrusion":
        out.extrusion = parseVmlExtrusion(child);
        break;
      case "o:callout":
        out.callout = parseVmlCallout(child);
        break;
      case "o:lock":
        out.lock = parseVmlLock(child);
        break;
      case "o:clippath":
        out.clippath = parseVmlClipPath(child);
        break;
      case "o:signatureline":
        out.signatureline = parseVmlSignatureLine(child);
        break;
      case "w10:wrap":
        out.wrap = parseVmlWrap(child);
        break;
      case "w10:anchorlock":
        out.anchorlock = parseVmlAnchorLock(child);
        break;
      case "w10:bordertop":
        out.bordertop = parseVmlBorder(child);
        break;
      case "w10:borderbottom":
        out.borderbottom = parseVmlBorder(child);
        break;
      case "w10:borderleft":
        out.borderleft = parseVmlBorder(child);
        break;
      case "w10:borderright":
        out.borderright = parseVmlBorder(child);
        break;
      case "x:ClientData":
        out.clientData = parseVmlClientData(child);
        break;
      case "pvml:textdata":
        out.textdata = parseVmlTextData(child);
        break;
      case "pvml:iscomment":
        out.iscomment = parseVmlIsComment(child);
        break;
      default:
        // Unmodeled child — keep it verbatim instead of dropping it. Word
        // exporters sometimes bind a standard namespace under a second prefix
        // (wvml: for office:word), which the literal-prefix cases above miss.
        // o:complex, o:ink and o:equationxml have dedicated collectors
        // outside this switch, and nested v: shapes are collected by the
        // group child loop — collecting any of them here too would emit
        // them twice.
        if (
          child.name === "o:complex" ||
          child.name === "o:ink" ||
          child.name === "o:equationxml" ||
          (child.name !== undefined && NESTED_SHAPE_NAMES.has(child.name))
        ) {
          break;
        }
        const previous = typeof out.rawChildrenXml === "string" ? out.rawChildrenXml : "";
        out.rawChildrenXml = previous + stringifyElement(child);
        break;
    }
  }
}

/** Generic shape element serializer: tag + shared/extra attrs + children. */
function stringifyShapeElement(
  tag: string,
  opts: Record<string, unknown>,
  extraSpecs: readonly VmlAttrSpec[] = [],
  childrenXml = "",
): string {
  const attrStr = stringifyShapeAttrs(opts, extraSpecs);
  return childrenXml !== "" ? `<${tag}${attrStr}>${childrenXml}</${tag}>` : `<${tag}${attrStr}/>`;
}

// ── v:shape / v:shapetype ──

const SHAPE_EXTRA_ATTRS: readonly VmlAttrSpec[] = [
  { field: "type", attr: "type", kind: "string" },
  { field: "adj", attr: "adj", kind: "string" },
  { field: "gfxdata", attr: "o:gfxdata", kind: "string" },
  { field: "equationxml", attr: "equationxml", kind: "string" },
];

/** Serialize v:shape. */
export function stringifyVmlShape(opts: VmlShapeOptions): string {
  let children = stringifyShapeElements(opts);
  if (opts.ink !== undefined) children += stringifyVmlInk(opts.ink);
  if (opts.equationxmlElement !== undefined) {
    children += stringifyVmlEquationXml(opts.equationxmlElement);
  }
  if (opts.iscomment !== undefined) children += stringifyVmlIsComment(opts.iscomment);
  return stringifyShapeElement(
    "v:shape",
    opts as unknown as Record<string, unknown>,
    SHAPE_EXTRA_ATTRS,
    children,
  );
}

/** Parse a v:shape element. */
export function parseVmlShape(el: XmlElement): VmlShapeOptions {
  const out: Record<string, unknown> = {};
  parseShapeAttrs(el, out, SHAPE_EXTRA_ATTRS);
  parseShapeElements(el, out);
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === "o:ink") out.ink = parseVmlInk(child);
    else if (child.name === "o:equationxml") out.equationxmlElement = parseVmlEquationXml(child);
  }
  return out as VmlShapeOptions;
}

const SHAPETYPE_EXTRA_ATTRS: readonly VmlAttrSpec[] = [
  { field: "adj", attr: "adj", kind: "string" },
  { field: "master", attr: "o:master", kind: "string" },
];

/** Serialize v:shapetype. */
export function stringifyVmlShapetype(opts: VmlShapetypeOptions): string {
  // CT_Shapetype is a sequence: the o:complex extension slot comes last.
  let children = stringifyShapeElements(opts);
  if (opts.complex !== undefined) children += stringifyVmlComplex(opts.complex);
  return stringifyShapeElement(
    "v:shapetype",
    opts as unknown as Record<string, unknown>,
    SHAPETYPE_EXTRA_ATTRS,
    children,
  );
}

/** Parse a v:shapetype element. */
export function parseVmlShapetype(el: XmlElement): VmlShapetypeOptions {
  const out: Record<string, unknown> = {};
  parseShapeAttrs(el, out, SHAPETYPE_EXTRA_ATTRS);
  parseShapeElements(el, out);
  for (const child of el.elements ?? []) {
    if (child.type === "element" && child.name === "o:complex") {
      out.complex = parseVmlComplex(child);
      break;
    }
  }
  return out as VmlShapetypeOptions;
}

// ── v:group ──

/** Serialize one VmlShapeChild. */
export function stringifyVmlShapeChild(child: VmlShapeChild): string {
  if ("shape" in child) return stringifyVmlShape(child.shape);
  if ("shapetype" in child) return stringifyVmlShapetype(child.shapetype);
  if ("group" in child) return stringifyVmlGroup(child.group);
  if ("arc" in child) return stringifyVmlArc(child.arc);
  if ("curve" in child) return stringifyVmlCurve(child.curve);
  if ("image" in child) return stringifyVmlImage(child.image);
  if ("line" in child) return stringifyVmlLine(child.line);
  if ("oval" in child) return stringifyVmlOval(child.oval);
  if ("polyline" in child) return stringifyVmlPolyline(child.polyline);
  if ("rect" in child) return stringifyVmlRect(child.rect);
  return stringifyVmlRoundRect(child.roundrect);
}

/** CT_Group attribute table — AG_AllCoreAttributes + AG_Fill + editas + table props. */
const GROUP_ATTRS: readonly VmlAttrSpec[] = [
  ...VML_CORE_ATTRS,
  ...VML_OFFICE_CORE_ATTRS,
  ...GROUP_FILL_ATTRS,
  { field: "editas", attr: "editas", kind: "string" },
  { field: "tableproperties", attr: "o:tableproperties", kind: "string" },
  { field: "tablelimits", attr: "o:tablelimits", kind: "string" },
  ...PATH_ATTR,
];

/** Parse one shape element into its VmlShapeChild wrapper, or undefined when
 *  the element is not part of the shape vocabulary. */
export function parseVmlShapeChild(el: XmlElement): VmlShapeChild | undefined {
  switch (el.name) {
    case "v:shape":
      return { shape: parseVmlShape(el) };
    case "v:shapetype":
      return { shapetype: parseVmlShapetype(el) };
    case "v:group":
      return { group: parseVmlGroup(el) };
    case "v:arc":
      return { arc: parseVmlArc(el) };
    case "v:curve":
      return { curve: parseVmlCurve(el) };
    case "v:image":
      return { image: parseVmlImage(el) };
    case "v:line":
      return { line: parseVmlLine(el) };
    case "v:oval":
      return { oval: parseVmlOval(el) };
    case "v:polyline":
      return { polyline: parseVmlPolyline(el) };
    case "v:rect":
      return { rect: parseVmlRect(el) };
    case "v:roundrect":
      return { roundrect: parseVmlRoundRect(el) };
    default:
      return undefined;
  }
}

/** Serialize v:group. */
export function stringifyVmlGroup(opts: VmlGroupOptions): string {
  let children =
    (opts.children ?? []).map(stringifyVmlShapeChild).join("") + stringifyShapeElements(opts);
  if (opts.diagram !== undefined) children += stringifyVmlDiagram(opts.diagram);
  const attrStr = stringifyVmlAttributes(opts as Record<string, unknown>, GROUP_ATTRS);
  const style = opts.style;
  const styleStr = style !== undefined ? ` style="${escapeXml(stringifyVmlStyle(style))}"` : "";
  return children !== ""
    ? `<v:group${attrStr}${styleStr}>${children}</v:group>`
    : `<v:group${attrStr}${styleStr}/>`;
}

/** Parse a v:group element. */
export function parseVmlGroup(el: XmlElement): VmlGroupOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, GROUP_ATTRS, out);
  if (el.attributes?.style !== undefined) {
    out.style = parseVmlShapeStyle(parseVmlStyle(String(el.attributes.style)));
  }
  parseShapeElements(el, out);

  const children: VmlShapeChild[] = [];
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    const shapeChild = parseVmlShapeChild(child);
    if (shapeChild !== undefined) {
      children.push(shapeChild);
    } else if (child.name === "o:diagram") {
      out.diagram = parseVmlDiagram(child);
    }
  }
  if (children.length > 0) out.children = children;
  return out as VmlGroupOptions;
}

// ── v:background ──

const BACKGROUND_ATTRS: readonly VmlAttrSpec[] = [
  { field: "id", attr: "id", kind: "string" },
  { field: "filled", attr: "filled", kind: "trueFalse" },
  { field: "fillcolor", attr: "fillcolor", kind: "string" },
  { field: "bwmode", attr: "o:bwmode", kind: "string" },
  { field: "bwpure", attr: "o:bwpure", kind: "string" },
  { field: "bwnormal", attr: "o:bwnormal", kind: "string" },
  { field: "targetscreensize", attr: "o:targetscreensize", kind: "string" },
];

/** Serialize v:background. */
export function stringifyVmlBackground(opts: VmlBackgroundOptions): string {
  // CT_Background carries AG_Id + AG_Fill + the o:bw* attributes only — it is
  // not an AG_AllCoreAttributes host, so it must not join the shared shape
  // attribute table (that would double-emit id).
  const attrStr = stringifyVmlAttributes(
    opts as unknown as Record<string, unknown>,
    BACKGROUND_ATTRS,
  );
  const fillXml = opts.fill !== undefined ? stringifyVmlFill(opts.fill) : "";
  return fillXml !== ""
    ? `<v:background${attrStr}>${fillXml}</v:background>`
    : `<v:background${attrStr}/>`;
}

/** Parse a v:background element. */
export function parseVmlBackground(el: XmlElement): VmlBackgroundOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, BACKGROUND_ATTRS, out);
  for (const child of el.elements ?? []) {
    if (child.type === "element" && child.name === "v:fill") {
      out.fill = parseVmlFill(child);
      break;
    }
  }
  return out as VmlBackgroundOptions;
}

// ── Eight basic shapes ──

/** Serialize v:arc. */
export function stringifyVmlArc(opts: VmlArcOptions): string {
  return stringifyShapeElement(
    "v:arc",
    opts as unknown as Record<string, unknown>,
    ARC_ATTRS,
    stringifyShapeElements(opts),
  );
}

/** Parse a v:arc element. */
export function parseVmlArc(el: XmlElement): VmlArcOptions {
  const out: Record<string, unknown> = {};
  parseShapeAttrs(el, out, ARC_ATTRS);
  parseShapeElements(el, out);
  return out as VmlArcOptions;
}

/** Serialize v:curve. */
export function stringifyVmlCurve(opts: VmlCurveOptions): string {
  return stringifyShapeElement(
    "v:curve",
    opts as unknown as Record<string, unknown>,
    CURVE_ATTRS,
    stringifyShapeElements(opts),
  );
}

/** Parse a v:curve element. */
export function parseVmlCurve(el: XmlElement): VmlCurveOptions {
  const out: Record<string, unknown> = {};
  parseShapeAttrs(el, out, CURVE_ATTRS);
  parseShapeElements(el, out);
  return out as VmlCurveOptions;
}

/** Serialize v:image. */
export function stringifyVmlImage(opts: VmlImageOptions): string {
  return stringifyShapeElement(
    "v:image",
    opts as unknown as Record<string, unknown>,
    IMAGE_ATTRS,
    stringifyShapeElements(opts),
  );
}

/** Parse a v:image element. */
export function parseVmlImage(el: XmlElement): VmlImageOptions {
  const out: Record<string, unknown> = {};
  parseShapeAttrs(el, out, IMAGE_ATTRS);
  parseShapeElements(el, out);
  return out as VmlImageOptions;
}

/** Serialize v:line. */
export function stringifyVmlLine(opts: VmlLineOptions): string {
  return stringifyShapeElement(
    "v:line",
    opts as unknown as Record<string, unknown>,
    LINE_ATTRS,
    stringifyShapeElements(opts),
  );
}

/** Parse a v:line element. */
export function parseVmlLine(el: XmlElement): VmlLineOptions {
  const out: Record<string, unknown> = {};
  parseShapeAttrs(el, out, LINE_ATTRS);
  parseShapeElements(el, out);
  return out as VmlLineOptions;
}

/** Serialize v:oval. */
export function stringifyVmlOval(opts: VmlOvalOptions): string {
  return stringifyShapeElement(
    "v:oval",
    opts as unknown as Record<string, unknown>,
    [],
    stringifyShapeElements(opts),
  );
}

/** Parse a v:oval element. */
export function parseVmlOval(el: XmlElement): VmlOvalOptions {
  const out: Record<string, unknown> = {};
  parseShapeAttrs(el, out);
  parseShapeElements(el, out);
  return out as VmlOvalOptions;
}

/** Serialize v:polyline. */
export function stringifyVmlPolyline(opts: VmlPolylineOptions): string {
  const children =
    stringifyShapeElements(opts) + (opts.ink !== undefined ? stringifyVmlInk(opts.ink) : "");
  return stringifyShapeElement(
    "v:polyline",
    opts as unknown as Record<string, unknown>,
    POLYLINE_ATTRS,
    children,
  );
}

/** Parse a v:polyline element. */
export function parseVmlPolyline(el: XmlElement): VmlPolylineOptions {
  const out: Record<string, unknown> = {};
  parseShapeAttrs(el, out, POLYLINE_ATTRS);
  parseShapeElements(el, out);
  for (const child of el.elements ?? []) {
    if (child.type === "element" && child.name === "o:ink") {
      out.ink = parseVmlInk(child);
      break;
    }
  }
  return out as VmlPolylineOptions;
}

/** Serialize v:rect. */
export function stringifyVmlRect(opts: VmlRectOptions): string {
  return stringifyShapeElement(
    "v:rect",
    opts as unknown as Record<string, unknown>,
    [],
    stringifyShapeElements(opts),
  );
}

/** Parse a v:rect element. */
export function parseVmlRect(el: XmlElement): VmlRectOptions {
  const out: Record<string, unknown> = {};
  parseShapeAttrs(el, out);
  parseShapeElements(el, out);
  return out as VmlRectOptions;
}

/** Serialize v:roundrect. */
export function stringifyVmlRoundRect(opts: VmlRoundRectOptions): string {
  return stringifyShapeElement(
    "v:roundrect",
    opts as unknown as Record<string, unknown>,
    ROUNDRECT_ATTRS,
    stringifyShapeElements(opts),
  );
}

/** Parse a v:roundrect element. */
export function parseVmlRoundRect(el: XmlElement): VmlRoundRectOptions {
  const out: Record<string, unknown> = {};
  parseShapeAttrs(el, out, ROUNDRECT_ATTRS);
  parseShapeElements(el, out);
  return out as VmlRoundRectOptions;
}
