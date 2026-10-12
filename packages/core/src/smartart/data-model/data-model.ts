/**
 * CT_DataModel XML stringifier — assembles complete SmartArt data model XML.
 *
 * @module
 */

import { attr, attrBool, attrNum, findChild, stringifyElement } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ReadContext, WriteContext } from "../../descriptor/context";
import type { BlackWhiteMode } from "../../drawing/black-white-mode";
import { presentationLayoutVariablesDesc } from "../../drawing/diagram/diagram-descriptors";
import {
  parseDiagramExtensions,
  stringifyDiagramExtensions,
  type DiagramExtensionOptions,
} from "../../drawing/diagram/diagram-props";
import type { EffectDagOptions } from "../../drawing/effects/effect-dag";
import { createEffectDag } from "../../drawing/effects/effect-dag";
import { effectListDesc } from "../../drawing/effects/effect-descriptors";
import type { EffectListOptions } from "../../drawing/effects/effect-list";
import { fillDesc, findFillChild } from "../../drawing/fill/fill-descriptors";
import type { FillOptions } from "../../drawing/fill/fill-options";
import type { OutlineOptions } from "../../drawing/outline/outline";
import { outlineDesc } from "../../drawing/outline/outline-descriptors";
import type { ShapePropertiesOptions } from "../../drawing/shape-properties-desc";
import type { TextBodyOptions } from "../../drawing/text/text-body";
import { parseShapeStyle } from "../../theme/style-matrix";
import { stringifyConnection } from "./connection";
import { parsePointDetails, stringifyPoint } from "./point";
import type { PointPropertySetOptions } from "./point";

/**
 * Build the complete dgm:dataModel XML from pre-serialized point and connection strings.
 */
export function stringifyDataModel(
  points: readonly string[],
  connections: readonly string[],
): string {
  return [
    '<dgm:dataModel xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">',
    `<dgm:ptLst>${points.join("")}</dgm:ptLst>`,
    `<dgm:cxnLst>${connections.join("")}</dgm:cxnLst>`,
    "<dgm:bg/>",
    "<dgm:whole/>",
    "</dgm:dataModel>",
  ].join("");
}

/** dgm:pt — one data-model point (dgm:pt, CT_Pt). */
export interface DataModelPointOptions {
  /** Point identity (dgm:pt `@modelId`). */
  modelId: string;
  /** Point kind (dgm:pt `@type`, ST_PtType). */
  type?: "node" | "asst" | "doc" | "pres" | "parTrans" | "sibTrans";
  /** Transition point owning connection (dgm:pt `@cxnId`, transition points only). */
  connectionId?: string;
  /** Plain-text projection of dgm:t. */
  text?: string;
  /** Complete DrawingML text body (dgm:t). */
  textBody?: TextBodyOptions;
  /** Container black-and-white mode (dgm:spPr `@bwMode`). */
  blackWhiteMode?: BlackWhiteMode;
  /** DrawingML shape properties (dgm:spPr). */
  shapeProperties?: ShapePropertiesOptions;
  propertySet?: PointPropertySetOptions;
  /** Point extensions (dgm:extLst). */
  extensions?: DiagramExtensionOptions[];
}

/** dgm:cxn — one data-model connection (CT_Cxn). */
export interface DataModelConnectionOptions {
  modelId: string;
  sourceId: string;
  destinationId: string;
  /** Connection kind (dgm:cxn `@type`, ST_CxnType). */
  type?: "parOf" | "presOf" | "presParOf" | "unknownRelationship";
  sourceOrder?: number;
  destinationOrder?: number;
  parentTransitionId?: string;
  siblingTransitionId?: string;
  presentationId?: string;
  /** Connection extensions (dgm:extLst). */
  extensions?: DiagramExtensionOptions[];
}

/** dgm:bg — CT_BackgroundFormatting. */
export interface DataModelBackgroundOptions {
  fill?: FillOptions | null;
  effects?: EffectListOptions;
  effectDag?: EffectDagOptions;
}

/** dgm:whole — CT_WholeE2oFormatting. */
export interface DataModelWholeOptions {
  outline?: OutlineOptions;
  effects?: EffectListOptions;
  effectDag?: EffectDagOptions;
}

/** dgm:dataModel — structured point/connection model (CT_DataModel). */
export interface DataModelOptions {
  points?: DataModelPointOptions[];
  connections?: DataModelConnectionOptions[];
  background?: DataModelBackgroundOptions;
  whole?: DataModelWholeOptions;
  extensions?: DiagramExtensionOptions[];
}

/**
 * Serialize a structured data model (dgm:dataModel). Inside a layoutDef the
 * part-level namespaces are already declared, so this emits the bare element.
 */
export function stringifyDataModelOptions(
  model: DataModelOptions,
  ctx: WriteContext = DIRECT_WRITE_CONTEXT,
): string {
  const points = (model.points ?? []).map((p) => stringifyDataModelPoint(p, ctx));
  const connections = (model.connections ?? []).map((c) => stringifyDataModelConnection(c));
  const background = model.background ? stringifyDataModelBackground(model.background, ctx) : "";
  const whole = model.whole ? stringifyDataModelWhole(model.whole, ctx) : "";
  const extensions = stringifyDiagramExtensions(model.extensions);
  return (
    "<dgm:dataModel>" +
    `<dgm:ptLst>${points.join("")}</dgm:ptLst>` +
    (model.connections ? `<dgm:cxnLst>${connections.join("")}</dgm:cxnLst>` : "") +
    background +
    whole +
    extensions +
    "</dgm:dataModel>"
  );
}

function stringifyDataModelPoint(point: DataModelPointOptions, ctx: WriteContext): string {
  return stringifyPoint(
    point.modelId,
    point.text ?? "",
    point.type ?? "",
    point.propertySet,
    point.connectionId,
    {
      blackWhiteMode: point.blackWhiteMode,
      shapeProperties: point.shapeProperties,
      textBody: point.textBody,
      extensions: point.extensions,
    },
    ctx,
  );
}

function stringifyDataModelConnection(connection: DataModelConnectionOptions): string {
  const extensions = stringifyDiagramExtensions(connection.extensions);
  const body = stringifyConnection({
    modelId: connection.modelId,
    srcId: connection.sourceId,
    destId: connection.destinationId,
    type: connection.type,
    srcOrd: connection.sourceOrder,
    destOrd: connection.destinationOrder,
    parTransId: connection.parentTransitionId,
    sibTransId: connection.siblingTransitionId,
    presId: connection.presentationId,
  });
  return extensions ? body.replace(/\/>$/, `>${extensions}</dgm:cxn>`) : body;
}

function stringifyEffectChoice(
  options: {
    effects?: EffectListOptions;
    effectDag?: EffectDagOptions;
  },
  ctx: WriteContext,
): string {
  if (options.effectDag) return createEffectDag(options.effectDag);
  if (options.effects) return effectListDesc.stringify(options.effects, ctx) ?? "";
  return "";
}

function stringifyDataModelBackground(
  options: DataModelBackgroundOptions,
  ctx: WriteContext,
): string {
  const fill = options.fill == null ? "" : (fillDesc.stringify(options.fill, ctx) ?? "");
  const effects = stringifyEffectChoice(options, ctx);
  return `<dgm:bg>${fill}${effects}</dgm:bg>`;
}

function stringifyDataModelWhole(options: DataModelWholeOptions, ctx: WriteContext): string {
  const outline = options.outline ? (outlineDesc.stringify(options.outline, ctx) ?? "") : "";
  const effects = stringifyEffectChoice(options, ctx);
  return `<dgm:whole>${outline}${effects}</dgm:whole>`;
}

const DIRECT_WRITE_CONTEXT: WriteContext = {
  addRelationship: () => "",
  addMedia: () => "",
  addHyperlink: () => {},
} as never;

function parseEffectChoice(
  el: Element,
  ctx: ReadContext,
): {
  effects?: EffectListOptions;
  effectDag?: EffectDagOptions;
} {
  const effectDag = findChild(el, "a:effectDag");
  if (effectDag) return { effectDag: { raw: stringifyElement(effectDag) } };
  const effectList = findChild(el, "a:effectLst");
  return effectList ? { effects: effectListDesc.parse(effectList, ctx) } : {};
}

const DIRECT_READ_CONTEXT: ReadContext = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
};

/** Parse dgm:dataModel into the structured model. */
export function parseDataModelOptions(
  el: Element | undefined,
  ctx: ReadContext = DIRECT_READ_CONTEXT,
): DataModelOptions | undefined {
  if (!el) return undefined;
  const result: DataModelOptions = {};

  const ptLst = findChild(el, "dgm:ptLst");
  if (ptLst) {
    const points: DataModelPointOptions[] = [];
    for (const pt of ptLst.elements ?? []) {
      if (pt.name !== "dgm:pt") continue;
      const point: DataModelPointOptions = { modelId: attr(pt, "modelId") ?? "" };
      const type = attr(pt, "type");
      if (type) point.type = type as DataModelPointOptions["type"];
      const cxnId = attr(pt, "cxnId");
      if (cxnId) point.connectionId = cxnId;
      Object.assign(point, parsePointDetails(pt, ctx));
      const t = findChild(pt, "dgm:t");
      if (t && !point.textBody) point.text = extractPointText(t);
      const prSet = findChild(pt, "dgm:prSet");
      if (prSet) point.propertySet = parsePointPropertySet(prSet, ctx);
      const pointExtensions = findChild(pt, "dgm:extLst");
      if (pointExtensions) point.extensions = parseDiagramExtensions(pointExtensions);
      points.push(point);
    }
    result.points = points;
  }

  const cxnLst = findChild(el, "dgm:cxnLst");
  if (cxnLst) {
    const connections: DataModelConnectionOptions[] = [];
    for (const cxn of cxnLst.elements ?? []) {
      if (cxn.name !== "dgm:cxn") continue;
      const connection: DataModelConnectionOptions = {
        modelId: attr(cxn, "modelId") ?? "",
        sourceId: attr(cxn, "srcId") ?? "",
        destinationId: attr(cxn, "destId") ?? "",
      };
      const type = attr(cxn, "type");
      if (type) connection.type = type as DataModelConnectionOptions["type"];
      const srcOrd = attrNum(cxn, "srcOrd");
      if (srcOrd !== undefined) connection.sourceOrder = srcOrd;
      const destOrd = attrNum(cxn, "destOrd");
      if (destOrd !== undefined) connection.destinationOrder = destOrd;
      const parTransId = attr(cxn, "parTransId");
      if (parTransId) connection.parentTransitionId = parTransId;
      const sibTransId = attr(cxn, "sibTransId");
      if (sibTransId) connection.siblingTransitionId = sibTransId;
      const presId = attr(cxn, "presId");
      if (presId) connection.presentationId = presId;
      const connectionExtensions = findChild(cxn, "dgm:extLst");
      if (connectionExtensions)
        connection.extensions = parseDiagramExtensions(connectionExtensions);
      connections.push(connection);
    }
    result.connections = connections;
  }

  const bg = findChild(el, "dgm:bg");
  if (bg) {
    const background: DataModelBackgroundOptions = {};
    if (findFillChild(bg)) background.fill = fillDesc.parse(bg, ctx);
    Object.assign(background, parseEffectChoice(bg, DIRECT_READ_CONTEXT));
    result.background = background;
  }

  const wholeEl = findChild(el, "dgm:whole");
  if (wholeEl) {
    const whole: DataModelWholeOptions = {};
    const outline = findChild(wholeEl, "a:ln");
    if (outline) whole.outline = outlineDesc.parse(outline, ctx);
    Object.assign(whole, parseEffectChoice(wholeEl, DIRECT_READ_CONTEXT));
    result.whole = whole;
  }

  const extLst = findChild(el, "dgm:extLst");
  if (extLst) result.extensions = parseDiagramExtensions(extLst);

  return Object.keys(result).length > 0 ? result : undefined;
}

/** dgm:t holds an a:p paragraph tree; flatten its text runs. */
function extractPointText(t: Element): string {
  let out = "";
  const walk = (node: Element) => {
    for (const child of node.elements ?? []) {
      if (child.type === "text") out += String(child.text ?? "");
      else walk(child);
    }
  };
  walk(t);
  return out;
}

/** Attribute names of dgm:prSet keyed by PointPropertySetOptions field. */
const PRSET_ATTRS: readonly (readonly [
  string,
  string,
  "string" | "number" | "boolean" | "angle" | "percent",
])[] = [
  ["presentationAssociationId", "presAssocID", "string"],
  ["presentationName", "presName", "string"],
  ["presentationStyleLabel", "presStyleLbl", "string"],
  ["presentationStyleIndex", "presStyleIdx", "number"],
  ["presentationStyleCount", "presStyleCnt", "number"],
  ["placeholderText", "phldrT", "string"],
  ["placeholder", "phldr", "boolean"],
  ["customAngle", "custAng", "angle"],
  ["customFlipVertical", "custFlipVert", "boolean"],
  ["customFlipHorizontal", "custFlipHor", "boolean"],
  ["customSizeX", "custSzX", "number"],
  ["customSizeY", "custSzY", "number"],
  ["customScaleX", "custScaleX", "percent"],
  ["customScaleY", "custScaleY", "percent"],
  ["customText", "custT", "boolean"],
  ["customLinearFactorX", "custLinFactX", "percent"],
  ["customLinearFactorY", "custLinFactY", "percent"],
  ["customLinearFactorNeighborX", "custLinFactNeighborX", "percent"],
  ["customLinearFactorNeighborY", "custLinFactNeighborY", "percent"],
  ["customRadialScaleRadius", "custRadScaleRad", "percent"],
  ["customRadialScaleIncrement", "custRadScaleInc", "percent"],
  ["coherent3DOffset", "coherent3DOff", "boolean"],
  ["layoutTypeId", "loTypeId", "string"],
  ["layoutCategoryId", "loCatId", "string"],
  ["quickStyleTypeId", "qsTypeId", "string"],
  ["quickStyleCategoryId", "qsCatId", "string"],
  ["colorStyleTypeId", "csTypeId", "string"],
  ["colorStyleCategoryId", "csCatId", "string"],
];

function parsePointPropertySet(prSet: Element, ctx: ReadContext): PointPropertySetOptions {
  const result: Record<string, unknown> = {};
  for (const [field, xmlName, kind] of PRSET_ATTRS) {
    if (kind === "number" || kind === "angle" || kind === "percent") {
      const v = attrNum(prSet, xmlName);
      if (v !== undefined) {
        result[field] = kind === "angle" ? v / 60000 : kind === "percent" ? v / 1000 : v;
      }
    } else if (kind === "boolean") {
      const v = attrBool(prSet, xmlName);
      if (v !== undefined) result[field] = v;
    } else {
      const v = attr(prSet, xmlName);
      if (v !== undefined) result[field] = v;
    }
  }
  const variables = findChild(prSet, "dgm:presLayoutVars");
  if (variables) {
    result.presentationLayoutVariables = presentationLayoutVariablesDesc.parse(variables, ctx);
  }
  const style = parseShapeStyle(findChild(prSet, "dgm:style"), ctx);
  if (style) result.style = style;
  return result as PointPropertySetOptions;
}
