/**
 * CT_DataModel XML stringifier — assembles complete SmartArt data model XML.
 *
 * @module
 */

import { attr, attrBool, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { stringifyConnection } from "./connection";
import { stringifyPoint } from "./point";
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
export interface ModelPointOptions {
  /** Point identity (dgm:pt `@modelId`). */
  modelId: string;
  /** Point kind (dgm:pt `@type`, ST_PtType). */
  type?: "node" | "asst" | "doc" | "pres" | "parTrans" | "sibTrans";
  /** Transition point owning connection (dgm:pt `@cxnId`, transition points only). */
  connectionId?: string;
  /** Point text (dgm:t). */
  text?: string;
  propertySet?: PointPropertySetOptions;
}

/** dgm:cxn — one data-model connection (CT_Cxn). */
export interface ModelConnectionOptions {
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
}

/** dgm:dataModel — structured point/connection model (CT_DataModel). */
export interface DataModelOptions {
  points?: ModelPointOptions[];
  connections?: ModelConnectionOptions[];
}

/**
 * Serialize a structured data model (dgm:dataModel). Inside a layoutDef the
 * part-level namespaces are already declared, so this emits the bare element.
 */
export function stringifyDataModelOptions(model: DataModelOptions): string {
  const points = (model.points ?? []).map((p) =>
    stringifyPoint(p.modelId, p.text ?? "", p.type ?? "node", p.propertySet, p.connectionId),
  );
  const connections = (model.connections ?? []).map((c) =>
    stringifyConnection({
      modelId: c.modelId,
      srcId: c.sourceId,
      destId: c.destinationId,
      type: c.type,
      srcOrd: c.sourceOrder,
      destOrd: c.destinationOrder,
      parTransId: c.parentTransitionId,
      sibTransId: c.siblingTransitionId,
      presId: c.presentationId,
    }),
  );
  return (
    "<dgm:dataModel>" +
    `<dgm:ptLst>${points.join("")}</dgm:ptLst>` +
    `<dgm:cxnLst>${connections.join("")}</dgm:cxnLst>` +
    "<dgm:bg/>" +
    "<dgm:whole/>" +
    "</dgm:dataModel>"
  );
}

/** Parse dgm:dataModel into the structured model. */
export function parseDataModelOptions(el: Element | undefined): DataModelOptions | undefined {
  if (!el) return undefined;
  const result: DataModelOptions = {};

  const ptLst = findChild(el, "dgm:ptLst");
  if (ptLst) {
    const points: ModelPointOptions[] = [];
    for (const pt of ptLst.elements ?? []) {
      if (pt.name !== "dgm:pt") continue;
      const point: ModelPointOptions = { modelId: attr(pt, "modelId") ?? "" };
      const type = attr(pt, "type");
      if (type) point.type = type as ModelPointOptions["type"];
      const cxnId = attr(pt, "cxnId");
      if (cxnId) point.connectionId = cxnId;
      const t = findChild(pt, "dgm:t");
      if (t) point.text = extractPointText(t);
      const prSet = findChild(pt, "dgm:prSet");
      if (prSet) point.propertySet = parsePointPropertySet(prSet);
      points.push(point);
    }
    result.points = points;
  }

  const cxnLst = findChild(el, "dgm:cxnLst");
  if (cxnLst) {
    const connections: ModelConnectionOptions[] = [];
    for (const cxn of cxnLst.elements ?? []) {
      if (cxn.name !== "dgm:cxn") continue;
      const connection: ModelConnectionOptions = {
        modelId: attr(cxn, "modelId") ?? "",
        sourceId: attr(cxn, "srcId") ?? "",
        destinationId: attr(cxn, "destId") ?? "",
      };
      const type = attr(cxn, "type");
      if (type) connection.type = type as ModelConnectionOptions["type"];
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
      connections.push(connection);
    }
    result.connections = connections;
  }

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
const PRSET_ATTRS: readonly (readonly [string, string, "string" | "number" | "boolean"])[] = [
  ["presentationAssociationId", "presAssocID", "string"],
  ["presentationName", "presName", "string"],
  ["presentationStyleLabel", "presStyleLbl", "string"],
  ["presentationStyleIndex", "presStyleIdx", "number"],
  ["presentationStyleCount", "presStyleCnt", "number"],
  ["placeholderText", "phldrT", "string"],
  ["placeholder", "phldr", "boolean"],
  ["customAngle", "custAng", "number"],
  ["customFlipVertical", "custFlipVert", "boolean"],
  ["customFlipHorizontal", "custFlipHor", "boolean"],
  ["customSizeX", "custSzX", "number"],
  ["customSizeY", "custSzY", "number"],
  ["customScaleX", "custScaleX", "number"],
  ["customScaleY", "custScaleY", "number"],
  ["customText", "custT", "boolean"],
  ["customLinearFactorX", "custLinFactX", "number"],
  ["customLinearFactorY", "custLinFactY", "number"],
  ["customLinearFactorNeighborX", "custLinFactNeighborX", "number"],
  ["customLinearFactorNeighborY", "custLinFactNeighborY", "number"],
  ["customRadialScaleRadius", "custRadScaleRad", "number"],
  ["customRadialScaleIncrement", "custRadScaleInc", "number"],
  ["coherent3DOffset", "coherent3DOff", "boolean"],
  ["hideGeometry", "hideGeom", "boolean"],
  ["hideLastTransition", "hideLastTrans", "boolean"],
  ["lockTextEntry", "lkTxEntry", "boolean"],
  ["moveWith", "moveWith", "string"],
  ["useDefault", "useDef", "boolean"],
  ["zOrderOffset", "zOrderOff", "number"],
  ["layoutTypeId", "loTypeId", "string"],
  ["layoutCategoryId", "loCatId", "string"],
  ["quickStyleTypeId", "qsTypeId", "string"],
  ["quickStyleCategoryId", "qsCatId", "string"],
  ["colorStyleTypeId", "csTypeId", "string"],
  ["colorStyleCategoryId", "csCatId", "string"],
];

function parsePointPropertySet(prSet: Element): PointPropertySetOptions {
  const result: Record<string, unknown> = {};
  for (const [field, xmlName, kind] of PRSET_ATTRS) {
    if (kind === "number") {
      const v = attrNum(prSet, xmlName);
      if (v !== undefined) result[field] = v;
    } else if (kind === "boolean") {
      const v = attrBool(prSet, xmlName);
      if (v !== undefined) result[field] = v;
    } else {
      const v = attr(prSet, xmlName);
      if (v !== undefined && v !== "") result[field] = v;
    }
  }
  return result as PointPropertySetOptions;
}
