import { attr, escapeXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ReadContext, WriteContext } from "../../descriptor/context";
import type { BlackWhiteMode } from "../../drawing/black-white-mode";
import { presentationLayoutVariablesDesc } from "../../drawing/diagram/diagram-descriptors";
import {
  parseDiagramExtensions,
  stringifyDiagramExtensions,
  type DiagramExtensionOptions,
} from "../../drawing/diagram/diagram-props";
import type { PresentationLayoutVariablesOptions } from "../../drawing/diagram/layout-vars";
import { shapePropertiesDesc } from "../../drawing/shape-properties-desc";
import type { ShapePropertiesOptions } from "../../drawing/shape-properties-desc";
import { textBodyDesc } from "../../drawing/text/text-body";
import type { TextBodyOptions } from "../../drawing/text/text-body";
import { stringifyShapeStyle } from "../../theme/style-matrix";
import type { DefaultShapeStyleOptions } from "../../theme/theme-options";

export interface PointPropertySetOptions {
  /** Presentation layout variables (dgm:presLayoutVars). */
  presentationLayoutVariables?: PresentationLayoutVariablesOptions;
  /** Shape style reference (dgm:style). */
  style?: DefaultShapeStyleOptions;
  presentationAssociationId?: string;
  presentationName?: string;
  presentationStyleLabel?: string;
  presentationStyleIndex?: number;
  presentationStyleCount?: number;
  placeholderText?: string;
  placeholder?: boolean;
  customAngle?: number;
  customFlipVertical?: boolean;
  customFlipHorizontal?: boolean;
  customSizeX?: number;
  customSizeY?: number;
  customScaleX?: number;
  customScaleY?: number;
  customText?: boolean;
  customLinearFactorX?: number;
  customLinearFactorY?: number;
  customLinearFactorNeighborX?: number;
  customLinearFactorNeighborY?: number;
  customRadialScaleRadius?: number;
  customRadialScaleIncrement?: number;
  coherent3DOffset?: boolean;
  layoutTypeId?: string;
  layoutCategoryId?: string;
  quickStyleTypeId?: string;
  quickStyleCategoryId?: string;
  colorStyleTypeId?: string;
  colorStyleCategoryId?: string;
}

/**
 * dgm:pt — SmartArt data model point (node) XML stringifier.
 */
export function stringifyPoint(
  modelId: string | number,
  text: string,
  type: string = "node",
  propertySet?: PointPropertySetOptions,
  connectionId?: string,
  details?: {
    blackWhiteMode?: BlackWhiteMode;
    shapeProperties?: ShapePropertiesOptions;
    textBody?: TextBodyOptions;
    extensions?: DiagramExtensionOptions[];
  },
  ctx: WriteContext = NOOP_WRITE_CONTEXT,
): string {
  const attrs = [`modelId="${escapeXml(String(modelId))}"`];
  if (connectionId) attrs.push(`cxnId="${escapeXml(connectionId)}"`);
  if (type) attrs.push(`type="${escapeXml(type)}"`);
  const parts: string[] = [];
  parts.push(`<dgm:pt ${attrs.join(" ")}>`);
  if (propertySet) {
    parts.push(stringifyPropertySet(propertySet, ctx));
  }
  if (details?.shapeProperties) {
    const spPr = shapePropertiesDesc.stringify(details.shapeProperties, ctx) ?? "";
    const blackWhiteMode = details.blackWhiteMode ? ` bwMode="${details.blackWhiteMode}"` : "";
    parts.push(`<dgm:spPr${blackWhiteMode}>${spPr}</dgm:spPr>`);
  }
  if (details?.textBody) {
    const textBody = textBodyDesc.stringify(details.textBody, NOOP_WRITE_CONTEXT) ?? "";
    parts.push(`<dgm:t>${textBody}</dgm:t>`);
  } else if (text || !details) {
    parts.push(stringifyPointText(text));
  }
  parts.push(stringifyDiagramExtensions(details?.extensions));
  parts.push("</dgm:pt>");
  return parts.join("");
}

export function parsePointDetails(
  point: Element,
  ctx: ReadContext,
): {
  blackWhiteMode?: BlackWhiteMode;
  shapeProperties?: ShapePropertiesOptions;
  textBody?: TextBodyOptions;
  extensions?: DiagramExtensionOptions[];
} {
  const result: {
    blackWhiteMode?: BlackWhiteMode;
    shapeProperties?: ShapePropertiesOptions;
    textBody?: TextBodyOptions;
    extensions?: DiagramExtensionOptions[];
  } = {};
  const spPr = point.elements?.find(
    (child) => child.type === "element" && child.name === "dgm:spPr",
  );
  if (spPr) {
    const blackWhiteMode = attr(spPr, "bwMode");
    if (blackWhiteMode) result.blackWhiteMode = blackWhiteMode as BlackWhiteMode;
    result.shapeProperties = shapePropertiesDesc.parse(spPr, ctx);
  }
  const t = point.elements?.find((child) => child.type === "element" && child.name === "dgm:t");
  if (t) result.textBody = textBodyDesc.parse(t, ctx);
  const extLst = point.elements?.find(
    (child) => child.type === "element" && child.name === "dgm:extLst",
  );
  if (extLst) result.extensions = parseDiagramExtensions(extLst);
  return result;
}

const NOOP_WRITE_CONTEXT: WriteContext = {
  addRelationship: () => "",
  addMedia: () => "",
  addHyperlink: () => {},
};

/**
 * dgm:pt — transition point (parTrans or sibTrans) XML stringifier.
 */
export function stringifyTransPoint(modelId: string, type: string, cxnId: string): string {
  return `<dgm:pt modelId="${escapeXml(modelId)}" type="${escapeXml(type)}" cxnId="${escapeXml(cxnId)}"><dgm:spPr/></dgm:pt>`;
}

function stringifyPropertySet(opts: PointPropertySetOptions, ctx: WriteContext): string {
  const attr: string[] = [];
  if (opts.presentationAssociationId !== undefined)
    attr.push(`presAssocID="${escapeXml(opts.presentationAssociationId)}"`);
  if (opts.presentationName !== undefined)
    attr.push(`presName="${escapeXml(opts.presentationName)}"`);
  if (opts.presentationStyleLabel !== undefined)
    attr.push(`presStyleLbl="${escapeXml(opts.presentationStyleLabel)}"`);
  if (opts.presentationStyleIndex !== undefined)
    attr.push(`presStyleIdx="${opts.presentationStyleIndex}"`);
  if (opts.presentationStyleCount !== undefined)
    attr.push(`presStyleCnt="${opts.presentationStyleCount}"`);
  if (opts.placeholderText !== undefined) attr.push(`phldrT="${escapeXml(opts.placeholderText)}"`);
  if (opts.placeholder !== undefined) attr.push(`phldr="${opts.placeholder ? 1 : 0}"`);
  if (opts.customAngle !== undefined)
    attr.push(`custAng="${Math.round(opts.customAngle * 60000)}"`);
  if (opts.customFlipVertical !== undefined)
    attr.push(`custFlipVert="${opts.customFlipVertical ? 1 : 0}"`);
  if (opts.customFlipHorizontal !== undefined)
    attr.push(`custFlipHor="${opts.customFlipHorizontal ? 1 : 0}"`);
  if (opts.customSizeX !== undefined) attr.push(`custSzX="${opts.customSizeX}"`);
  if (opts.customSizeY !== undefined) attr.push(`custSzY="${opts.customSizeY}"`);
  if (opts.customScaleX !== undefined)
    attr.push(`custScaleX="${Math.round(opts.customScaleX * 1000)}"`);
  if (opts.customScaleY !== undefined)
    attr.push(`custScaleY="${Math.round(opts.customScaleY * 1000)}"`);
  if (opts.customText !== undefined) attr.push(`custT="${opts.customText ? 1 : 0}"`);
  if (opts.customLinearFactorX !== undefined)
    attr.push(`custLinFactX="${Math.round(opts.customLinearFactorX * 1000)}"`);
  if (opts.customLinearFactorY !== undefined)
    attr.push(`custLinFactY="${Math.round(opts.customLinearFactorY * 1000)}"`);
  if (opts.customLinearFactorNeighborX !== undefined)
    attr.push(`custLinFactNeighborX="${Math.round(opts.customLinearFactorNeighborX * 1000)}"`);
  if (opts.customLinearFactorNeighborY !== undefined)
    attr.push(`custLinFactNeighborY="${Math.round(opts.customLinearFactorNeighborY * 1000)}"`);
  if (opts.customRadialScaleRadius !== undefined)
    attr.push(`custRadScaleRad="${Math.round(opts.customRadialScaleRadius * 1000)}"`);
  if (opts.customRadialScaleIncrement !== undefined)
    attr.push(`custRadScaleInc="${Math.round(opts.customRadialScaleIncrement * 1000)}"`);
  if (opts.coherent3DOffset !== undefined)
    attr.push(`coherent3DOff="${opts.coherent3DOffset ? 1 : 0}"`);
  if (opts.layoutTypeId !== undefined) attr.push(`loTypeId="${escapeXml(opts.layoutTypeId)}"`);
  if (opts.layoutCategoryId !== undefined)
    attr.push(`loCatId="${escapeXml(opts.layoutCategoryId)}"`);
  if (opts.quickStyleTypeId !== undefined)
    attr.push(`qsTypeId="${escapeXml(opts.quickStyleTypeId)}"`);
  if (opts.quickStyleCategoryId !== undefined)
    attr.push(`qsCatId="${escapeXml(opts.quickStyleCategoryId)}"`);
  if (opts.colorStyleTypeId !== undefined)
    attr.push(`csTypeId="${escapeXml(opts.colorStyleTypeId)}"`);
  if (opts.colorStyleCategoryId !== undefined)
    attr.push(`csCatId="${escapeXml(opts.colorStyleCategoryId)}"`);

  const children =
    (opts.presentationLayoutVariables
      ? (presentationLayoutVariablesDesc.stringify(opts.presentationLayoutVariables, ctx) ?? "")
      : "") + (opts.style ? stringifyShapeStyle(opts.style, ctx, "dgm:style") : "");
  if (attr.length === 0 && !children) return "<dgm:prSet/>";
  return `<dgm:prSet ${attr.join(" ")}>${children}</dgm:prSet>`;
}

function stringifyPointText(text: string): string {
  const body = text ? `<a:r><a:t>${escapeXml(text)}</a:t></a:r>` : "";
  return `<dgm:t><a:bodyPr/><a:lstStyle/><a:p>${body}</a:p></dgm:t>`;
}
