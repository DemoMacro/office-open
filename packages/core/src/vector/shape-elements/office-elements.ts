/**
 * Office-drawing o: elements — the EG_ShapeElements members defined in
 * vml-officeDrawing.xsd (skew, extrusion, callout, lock, clippath,
 * signatureline, ink, equationxml, diagram) plus the standalone o:fill,
 * o:complex and o:OLEObject.
 *
 * Every element carries the v:ext attribute (AG_Ext, ST_Ext) exposing the
 * extension-slot behavior.
 *
 * Reference: ISO/IEC 29500-4, vml-officeDrawing.xsd.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml, stringifyElement } from "@office-open/xml";

import type { Guid } from "../../util/values";
import {
  stringifyVmlAttributes,
  parseVmlAttributes,
  stringifyVmlTrueFalse,
  parseVmlTrueFalse,
  stringifyVmlTrueFalseBlank,
  parseVmlTrueFalseBlank,
  type VmlAttrSpec,
  type VmlColor,
  type VmlTrueFalse,
  type VmlTrueFalseBlank,
} from "../attributes";

export { stringifyVmlTrueFalseBlank, parseVmlTrueFalseBlank };

// The scalar o: enums live in `attributes.ts` (they feed the shared attribute
// tables); re-exported here so the office-elements module stays the one-stop
// import for its consumers.
export type {
  VmlBlackWhiteMode,
  VmlConnectorType,
  VmlHorizontalRuleAlign,
  VmlInsetMode,
  VmlConnectType,
  VmlScreenSize,
  VmlDiagramLayout,
  VmlTrueFalseBlank,
} from "../attributes";

// ── Shared types ──

/** ST_Ext — the v:ext extension-slot attribute carried by every o: element. */
export type VmlExtensionMode = "view" | "edit" | "backwardCompatible";

/** The AG_Ext attribute present on every o: element. */
export interface VmlExtAttribute {
  ext?: VmlExtensionMode;
}

/** AG_Ext as a spec entry — prepend to every o: element's spec list. */
export const EXT_ATTR: readonly VmlAttrSpec[] = [{ field: "ext", attr: "v:ext", kind: "string" }];

// ── o:skew ──

/** o:skew options (CT_Skew). */
export interface VmlSkewOptions extends VmlExtAttribute {
  id?: string;
  on?: VmlTrueFalse;
  offset?: string;
  origin?: string;
  matrix?: string;
}

const SKEW_ATTRS: readonly VmlAttrSpec[] = [
  { field: "id", attr: "id", kind: "string" },
  { field: "on", attr: "on", kind: "trueFalse" },
  { field: "offset", attr: "offset", kind: "string" },
  { field: "origin", attr: "origin", kind: "string" },
  { field: "matrix", attr: "matrix", kind: "string" },
];

/** Serialize o:skew. */
export function stringifyVmlSkew(opts: VmlSkewOptions): string {
  return `<o:skew${stringifyVmlAttributes(opts as Record<string, unknown>, [...EXT_ATTR, ...SKEW_ATTRS])}/>`;
}

/** Parse an o:skew element. */
export function parseVmlSkew(el: XmlElement): VmlSkewOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, [...EXT_ATTR, ...SKEW_ATTRS], out);
  return out as VmlSkewOptions;
}

// ── o:extrusion ──

/** ST_ExtrusionType. */
export type VmlExtrusionType = "perspective" | "parallel";

/** ST_ExtrusionRender. */
export type VmlExtrusionRender = "solid" | "wireFrame" | "boundingCube";

/** ST_ExtrusionPlane. */
export type VmlExtrusionPlane = "XY" | "ZX" | "YZ";

/** ST_ColorMode. */
export type VmlColorMode = "auto" | "custom";

/** o:extrusion options (CT_Extrusion). */
export interface VmlExtrusionOptions extends VmlExtAttribute {
  on?: VmlTrueFalse;
  type?: VmlExtrusionType;
  render?: VmlExtrusionRender;
  viewpointorigin?: string;
  viewpoint?: string;
  plane?: VmlExtrusionPlane;
  skewangle?: number;
  skewamt?: string;
  foredepth?: string;
  backdepth?: string;
  orientation?: string;
  orientationangle?: number;
  lockrotationcenter?: VmlTrueFalse;
  autorotationcenter?: VmlTrueFalse;
  rotationcenter?: string;
  rotationangle?: string;
  colormode?: VmlColorMode;
  color?: VmlColor;
  shininess?: number;
  specularity?: string;
  diffusity?: string;
  metal?: VmlTrueFalse;
  edge?: string;
  facet?: string;
  lightface?: VmlTrueFalse;
  brightness?: string;
  lightposition?: string;
  lightlevel?: string;
  lightharsh?: VmlTrueFalse;
  lightposition2?: string;
  lightlevel2?: string;
  lightharsh2?: VmlTrueFalse;
}

const EXTRUSION_ATTRS: readonly VmlAttrSpec[] = [
  { field: "on", attr: "on", kind: "trueFalse" },
  { field: "type", attr: "type", kind: "string" },
  { field: "render", attr: "render", kind: "string" },
  { field: "viewpointorigin", attr: "viewpointorigin", kind: "string" },
  { field: "viewpoint", attr: "viewpoint", kind: "string" },
  { field: "plane", attr: "plane", kind: "string" },
  { field: "skewangle", attr: "skewangle", kind: "number" },
  { field: "skewamt", attr: "skewamt", kind: "string" },
  { field: "foredepth", attr: "foredepth", kind: "string" },
  { field: "backdepth", attr: "backdepth", kind: "string" },
  { field: "orientation", attr: "orientation", kind: "string" },
  { field: "orientationangle", attr: "orientationangle", kind: "number" },
  { field: "lockrotationcenter", attr: "lockrotationcenter", kind: "trueFalse" },
  { field: "autorotationcenter", attr: "autorotationcenter", kind: "trueFalse" },
  { field: "rotationcenter", attr: "rotationcenter", kind: "string" },
  { field: "rotationangle", attr: "rotationangle", kind: "string" },
  { field: "colormode", attr: "colormode", kind: "string" },
  { field: "color", attr: "color", kind: "string" },
  { field: "shininess", attr: "shininess", kind: "number" },
  { field: "specularity", attr: "specularity", kind: "string" },
  { field: "diffusity", attr: "diffusity", kind: "string" },
  { field: "metal", attr: "metal", kind: "trueFalse" },
  { field: "edge", attr: "edge", kind: "string" },
  { field: "facet", attr: "facet", kind: "string" },
  { field: "lightface", attr: "lightface", kind: "trueFalse" },
  { field: "brightness", attr: "brightness", kind: "string" },
  { field: "lightposition", attr: "lightposition", kind: "string" },
  { field: "lightlevel", attr: "lightlevel", kind: "string" },
  { field: "lightharsh", attr: "lightharsh", kind: "trueFalse" },
  { field: "lightposition2", attr: "lightposition2", kind: "string" },
  { field: "lightlevel2", attr: "lightlevel2", kind: "string" },
  { field: "lightharsh2", attr: "lightharsh2", kind: "trueFalse" },
];

/** Serialize o:extrusion. */
export function stringifyVmlExtrusion(opts: VmlExtrusionOptions): string {
  return `<o:extrusion${stringifyVmlAttributes(opts as Record<string, unknown>, [...EXT_ATTR, ...EXTRUSION_ATTRS])}/>`;
}

/** Parse an o:extrusion element. */
export function parseVmlExtrusion(el: XmlElement): VmlExtrusionOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, [...EXT_ATTR, ...EXTRUSION_ATTRS], out);
  return out as VmlExtrusionOptions;
}

// ── o:callout ──

/** ST_Angle — the callout leader-angle presets, kept in XML spelling. */
export type VmlCalloutAngle = "any" | "30" | "45" | "60" | "90" | "auto";

/** o:callout options (CT_Callout). */
export interface VmlCalloutOptions extends VmlExtAttribute {
  on?: VmlTrueFalse;
  type?: string;
  gap?: string;
  angle?: VmlCalloutAngle;
  dropauto?: VmlTrueFalse;
  drop?: string;
  distance?: string;
  lengthspecified?: VmlTrueFalse;
  length?: string;
  accentbar?: VmlTrueFalse;
  textborder?: VmlTrueFalse;
  minusx?: VmlTrueFalse;
  minusy?: VmlTrueFalse;
}

const CALLOUT_ATTRS: readonly VmlAttrSpec[] = [
  { field: "on", attr: "on", kind: "trueFalse" },
  { field: "type", attr: "type", kind: "string" },
  { field: "gap", attr: "gap", kind: "string" },
  { field: "angle", attr: "angle", kind: "string" },
  { field: "dropauto", attr: "dropauto", kind: "trueFalse" },
  { field: "drop", attr: "drop", kind: "string" },
  { field: "distance", attr: "distance", kind: "string" },
  { field: "lengthspecified", attr: "lengthspecified", kind: "trueFalse" },
  { field: "length", attr: "length", kind: "string" },
  { field: "accentbar", attr: "accentbar", kind: "trueFalse" },
  { field: "textborder", attr: "textborder", kind: "trueFalse" },
  { field: "minusx", attr: "minusx", kind: "trueFalse" },
  { field: "minusy", attr: "minusy", kind: "trueFalse" },
];

/** Serialize o:callout. */
export function stringifyVmlCallout(opts: VmlCalloutOptions): string {
  return `<o:callout${stringifyVmlAttributes(opts as Record<string, unknown>, [...EXT_ATTR, ...CALLOUT_ATTRS])}/>`;
}

/** Parse an o:callout element. */
export function parseVmlCallout(el: XmlElement): VmlCalloutOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, [...EXT_ATTR, ...CALLOUT_ATTRS], out);
  return out as VmlCalloutOptions;
}

// ── o:lock ──

/** o:lock options (CT_Lock) — a true flag locks the capability. */
export interface VmlLockOptions extends VmlExtAttribute {
  position?: VmlTrueFalse;
  selection?: VmlTrueFalse;
  grouping?: VmlTrueFalse;
  ungrouping?: VmlTrueFalse;
  rotation?: VmlTrueFalse;
  cropping?: VmlTrueFalse;
  verticies?: VmlTrueFalse;
  adjusthandles?: VmlTrueFalse;
  text?: VmlTrueFalse;
  aspectratio?: VmlTrueFalse;
  shapetype?: VmlTrueFalse;
}

const LOCK_ATTRS: readonly VmlAttrSpec[] = [
  { field: "position", attr: "position", kind: "trueFalse" },
  { field: "selection", attr: "selection", kind: "trueFalse" },
  { field: "grouping", attr: "grouping", kind: "trueFalse" },
  { field: "ungrouping", attr: "ungrouping", kind: "trueFalse" },
  { field: "rotation", attr: "rotation", kind: "trueFalse" },
  { field: "cropping", attr: "cropping", kind: "trueFalse" },
  { field: "verticies", attr: "verticies", kind: "trueFalse" },
  { field: "adjusthandles", attr: "adjusthandles", kind: "trueFalse" },
  { field: "text", attr: "text", kind: "trueFalse" },
  { field: "aspectratio", attr: "aspectratio", kind: "trueFalse" },
  { field: "shapetype", attr: "shapetype", kind: "trueFalse" },
];

/** Serialize o:lock. */
export function stringifyVmlLock(opts: VmlLockOptions): string {
  return `<o:lock${stringifyVmlAttributes(opts as Record<string, unknown>, [...EXT_ATTR, ...LOCK_ATTRS])}/>`;
}

/** Parse an o:lock element. */
export function parseVmlLock(el: XmlElement): VmlLockOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, [...EXT_ATTR, ...LOCK_ATTRS], out);
  return out as VmlLockOptions;
}

// ── o:clippath ──

/** o:clippath options (CT_ClipPath). The path attribute is qualified: `v:v`. */
export interface VmlClipPathOptions {
  /** Path command string, serialized as the `v:v` attribute. */
  v: string;
}

/** Serialize o:clippath. */
export function stringifyVmlClipPath(opts: VmlClipPathOptions): string {
  return `<o:clippath v:v="${escapeXml(opts.v)}"/>`;
}

/** Parse an o:clippath element. */
export function parseVmlClipPath(el: XmlElement): VmlClipPathOptions {
  return { v: String(el.attributes?.["v:v"] ?? "") };
}

// ── o:signatureline ──

/** o:signatureline options (CT_SignatureLine). */
export interface VmlSignatureLineOptions extends VmlExtAttribute {
  issignatureline?: VmlTrueFalse;
  /** s:ST_Guid — a GUID string. */
  id?: Guid;
  provid?: Guid;
  signinginstructionsset?: VmlTrueFalse;
  allowcomments?: VmlTrueFalse;
  showsigndate?: VmlTrueFalse;
  suggestedsigner?: string;
  suggestedsigner2?: string;
  suggestedsigneremail?: string;
  signinginstructions?: string;
  addlxml?: string;
  sigprovurl?: string;
}

const SIGNATURE_ATTRS: readonly VmlAttrSpec[] = [
  { field: "issignatureline", attr: "issignatureline", kind: "trueFalse" },
  { field: "id", attr: "id", kind: "string" },
  { field: "provid", attr: "provid", kind: "string" },
  { field: "signinginstructionsset", attr: "signinginstructionsset", kind: "trueFalse" },
  { field: "allowcomments", attr: "allowcomments", kind: "trueFalse" },
  { field: "showsigndate", attr: "showsigndate", kind: "trueFalse" },
  // form="qualified" in the XSD — serialized with the o: prefix.
  { field: "suggestedsigner", attr: "o:suggestedsigner", kind: "string" },
  { field: "suggestedsigner2", attr: "o:suggestedsigner2", kind: "string" },
  { field: "suggestedsigneremail", attr: "o:suggestedsigneremail", kind: "string" },
  { field: "signinginstructions", attr: "signinginstructions", kind: "string" },
  { field: "addlxml", attr: "addlxml", kind: "string" },
  { field: "sigprovurl", attr: "sigprovurl", kind: "string" },
];

/** Serialize o:signatureline. */
export function stringifyVmlSignatureLine(opts: VmlSignatureLineOptions): string {
  return `<o:signatureline${stringifyVmlAttributes(opts as Record<string, unknown>, [...EXT_ATTR, ...SIGNATURE_ATTRS])}/>`;
}

/** Parse an o:signatureline element. */
export function parseVmlSignatureLine(el: XmlElement): VmlSignatureLineOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, [...EXT_ATTR, ...SIGNATURE_ATTRS], out);
  return out as VmlSignatureLineOptions;
}

// ── o:ink ──

/** o:ink options (CT_Ink) — Tablet PC ink annotation data. */
export interface VmlInkOptions {
  /** Ink data — an ISF-encoded string. */
  i: string;
  annotation?: VmlTrueFalse;
  contentType?: string;
}

/** Serialize o:ink (hand-written — `i` is required, so no table-driven cast). */
export function stringifyVmlInk(opts: VmlInkOptions): string {
  const attrs = [`i="${escapeXml(opts.i)}"`];
  if (opts.annotation !== undefined) {
    attrs.push(`annotation="${stringifyVmlTrueFalse(opts.annotation)}"`);
  }
  if (opts.contentType !== undefined) attrs.push(`contentType="${escapeXml(opts.contentType)}"`);
  return `<o:ink ${attrs.join(" ")}/>`;
}

/** Parse an o:ink element. */
export function parseVmlInk(el: XmlElement): VmlInkOptions {
  const out: VmlInkOptions = { i: String(el.attributes?.i ?? "") };
  if (el.attributes?.annotation !== undefined) {
    out.annotation = parseVmlTrueFalse(String(el.attributes.annotation));
  }
  if (el.attributes?.contentType !== undefined) {
    out.contentType = String(el.attributes.contentType);
  }
  return out;
}

// ── o:equationxml ──

/**
 * o:equationxml options (CT_EquationXml) — the alternative math content,
 * carried as verbatim inner XML (the XSD allows any single element).
 */
export interface VmlEquationXmlOptions {
  contentType?: string;
  /** Inner XML of the alternate equation element, serialized verbatim. */
  content?: string;
}

/** Serialize o:equationxml. */
export function stringifyVmlEquationXml(opts: VmlEquationXmlOptions): string {
  const attrStr =
    opts.contentType !== undefined ? ` contentType="${escapeXml(opts.contentType)}"` : "";
  return `<o:equationxml${attrStr}>${opts.content ?? ""}</o:equationxml>`;
}

/** Parse an o:equationxml element. */
export function parseVmlEquationXml(el: XmlElement): VmlEquationXmlOptions {
  const out: VmlEquationXmlOptions = {};
  const contentType = el.attributes?.contentType;
  if (contentType !== undefined) out.contentType = String(contentType);
  if ((el.elements ?? []).length > 0) {
    const child = el.elements!.find((c) => c.type === "element");
    if (child) {
      out.content = stringifyElement(child);
    }
  }
  return out;
}

// ── o:diagram ──

/** CT_Relation — one relationtable row. */
export interface VmlDiagramRelationOptions extends VmlExtAttribute {
  idsrc?: string;
  iddest?: string;
  idcntr?: string;
}

const RELATION_ATTRS: readonly VmlAttrSpec[] = [
  { field: "idsrc", attr: "idsrc", kind: "string" },
  { field: "iddest", attr: "iddest", kind: "string" },
  { field: "idcntr", attr: "idcntr", kind: "string" },
];

/** o:diagram options (CT_Diagram) — legacy org-chart/diagram shape group. */
export interface VmlDiagramOptions extends VmlExtAttribute {
  dgmstyle?: number;
  autoformat?: VmlTrueFalse;
  reverse?: VmlTrueFalse;
  autolayout?: VmlTrueFalse;
  dgmscalex?: number;
  dgmscaley?: number;
  dgmfontsize?: number;
  constrainbounds?: string;
  dgmbasetextscale?: number;
  relations?: VmlDiagramRelationOptions[];
}

const DIAGRAM_ATTRS: readonly VmlAttrSpec[] = [
  { field: "dgmstyle", attr: "dgmstyle", kind: "number" },
  { field: "autoformat", attr: "autoformat", kind: "trueFalse" },
  { field: "reverse", attr: "reverse", kind: "trueFalse" },
  { field: "autolayout", attr: "autolayout", kind: "trueFalse" },
  { field: "dgmscalex", attr: "dgmscalex", kind: "number" },
  { field: "dgmscaley", attr: "dgmscaley", kind: "number" },
  { field: "dgmfontsize", attr: "dgmfontsize", kind: "number" },
  { field: "constrainbounds", attr: "constrainbounds", kind: "string" },
  { field: "dgmbasetextscale", attr: "dgmbasetextscale", kind: "number" },
];

/** Serialize o:diagram. */
export function stringifyVmlDiagram(opts: VmlDiagramOptions): string {
  const relations = (opts.relations ?? [])
    .map(
      (rel) =>
        `<o:rel${stringifyVmlAttributes(rel as Record<string, unknown>, [...EXT_ATTR, ...RELATION_ATTRS])}/>`,
    )
    .join("");
  const children = relations !== "" ? `<o:relationtable>${relations}</o:relationtable>` : "";
  return `<o:diagram${stringifyVmlAttributes(opts as Record<string, unknown>, [...EXT_ATTR, ...DIAGRAM_ATTRS])}>${children}</o:diagram>`;
}

/** Parse an o:diagram element. */
export function parseVmlDiagram(el: XmlElement): VmlDiagramOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, [...EXT_ATTR, ...DIAGRAM_ATTRS], out);
  const relationTable = (el.elements ?? []).find(
    (c) => c.type === "element" && c.name === "o:relationtable",
  );
  if (relationTable) {
    const relations: VmlDiagramRelationOptions[] = [];
    for (const child of relationTable.elements ?? []) {
      if (child.type === "element" && child.name === "o:rel") {
        const rel: Record<string, unknown> = {};
        parseVmlAttributes(child, [...EXT_ATTR, ...RELATION_ATTRS], rel);
        relations.push(rel as VmlDiagramRelationOptions);
      }
    }
    if (relations.length > 0) out.relations = relations;
  }
  return out as VmlDiagramOptions;
}

// ── o:complex ──

/** o:complex options (CT_Complex) — the extension slot at the end of CT_Shapetype. */
export interface VmlComplexOptions extends VmlExtAttribute {}

/** Serialize o:complex. */
export function stringifyVmlComplex(opts: VmlComplexOptions): string {
  return `<o:complex${stringifyVmlAttributes(opts as Record<string, unknown>, EXT_ATTR)}/>`;
}

/** Parse an o:complex element. */
export function parseVmlComplex(el: XmlElement): VmlComplexOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, EXT_ATTR, out);
  return out as VmlComplexOptions;
}

// ── o:fill ──

/** ST_FillType (o: namespace form — includes the gradientCenter/background variants). */
export type VmlOfficeFillType =
  | "gradientCenter"
  | "solid"
  | "pattern"
  | "tile"
  | "frame"
  | "gradientUnscaled"
  | "gradientRadial"
  | "gradient"
  | "background";

/** o:fill options (CT_Fill) — the o: extension child of v:fill. */
export interface VmlOfficeFillOptions extends VmlExtAttribute {
  type?: VmlOfficeFillType;
}

/** Serialize o:fill. */
export function stringifyVmlOfficeFill(opts: VmlOfficeFillOptions): string {
  return `<o:fill${stringifyVmlAttributes(opts as Record<string, unknown>, EXT_ATTR)}${opts.type !== undefined ? ` type="${opts.type}"` : ""}/>`;
}

/** Parse an o:fill element. */
export function parseVmlOfficeFill(el: XmlElement): VmlOfficeFillOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, EXT_ATTR, out);
  if (el.attributes?.type !== undefined) out.type = String(el.attributes.type);
  return out as VmlOfficeFillOptions;
}

// ── o:OLEObject ──

/** ST_OLEType. */
export type VmlOleObjectType = "Embed" | "Link";

/** ST_OLEDrawAspect. */
export type VmlOleDrawAspect = "Content" | "Icon";

/** ST_OLEUpdateMode. */
export type VmlOleUpdateMode = "Always" | "OnCall";

/**
 * o:OLEObject options (CT_OLEObject). Attribute names keep the PascalCase
 * XML spelling (ProgID, ShapeID, …) — they mirror the serialized attribute.
 */
export interface VmlOleObjectOptions {
  Type?: VmlOleObjectType;
  ProgID?: string;
  ShapeID?: string;
  DrawAspect?: VmlOleDrawAspect;
  ObjectID?: string;
  /** r:id — relationship id, bridged by the caller. */
  relationshipId?: string;
  UpdateMode?: VmlOleUpdateMode;
  /** LinkType child (link-source type string). */
  linkType?: string;
  /** LockedField child. */
  lockedField?: VmlTrueFalseBlank;
  /** FieldCodes child (the FIELD instruction when the object is a field result). */
  fieldCodes?: string;
}

const OLE_OBJECT_ATTRS: readonly VmlAttrSpec[] = [
  { field: "Type", attr: "Type", kind: "string" },
  { field: "ProgID", attr: "ProgID", kind: "string" },
  { field: "ShapeID", attr: "ShapeID", kind: "string" },
  { field: "DrawAspect", attr: "DrawAspect", kind: "string" },
  { field: "ObjectID", attr: "ObjectID", kind: "string" },
  { field: "relationshipId", attr: "r:id", kind: "string" },
  { field: "UpdateMode", attr: "UpdateMode", kind: "string" },
];

/** Serialize o:OLEObject. */
export function stringifyVmlOleObject(opts: VmlOleObjectOptions): string {
  const children: string[] = [];
  if (opts.linkType !== undefined) {
    children.push(`<o:LinkType>${escapeXml(opts.linkType)}</o:LinkType>`);
  }
  if (opts.lockedField !== undefined) {
    children.push(`<o:LockedField>${stringifyVmlTrueFalseBlank(opts.lockedField)}</o:LockedField>`);
  }
  if (opts.fieldCodes !== undefined) {
    children.push(`<o:FieldCodes>${escapeXml(opts.fieldCodes)}</o:FieldCodes>`);
  }
  const attrStr = stringifyVmlAttributes(opts as Record<string, unknown>, OLE_OBJECT_ATTRS);
  return children.length > 0
    ? `<o:OLEObject${attrStr}>${children.join("")}</o:OLEObject>`
    : `<o:OLEObject${attrStr}/>`;
}

/** Read the concatenated text content of an element (LinkType/FieldCodes children). */
function elementText(el: XmlElement): string {
  return (el.elements ?? []).map((child) => String(child.text ?? "")).join("");
}

/** Parse an o:OLEObject element. */
export function parseVmlOleObject(el: XmlElement): VmlOleObjectOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, OLE_OBJECT_ATTRS, out);
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === "o:LinkType") {
      out.linkType = elementText(child);
    } else if (child.name === "o:LockedField") {
      out.lockedField = parseVmlTrueFalseBlank(elementText(child));
    } else if (child.name === "o:FieldCodes") {
      out.fieldCodes = elementText(child);
    }
  }
  return out as VmlOleObjectOptions;
}
