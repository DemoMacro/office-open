import { escapeXml } from "@office-open/xml";

export interface PointPropertySetOptions {
  readonly presentationAssociationId?: string;
  readonly presentationName?: string;
  readonly presentationStyleLabel?: string;
  readonly presentationStyleIndex?: number;
  readonly presentationStyleCount?: number;
  readonly placeholderText?: string;
  readonly placeholder?: boolean;
  readonly customAngle?: number;
  readonly customFlipVertical?: boolean;
  readonly customFlipHorizontal?: boolean;
  readonly customSizeX?: number;
  readonly customSizeY?: number;
  readonly customScaleX?: number;
  readonly customScaleY?: number;
  readonly customText?: boolean;
  readonly customLinearFactorX?: number;
  readonly customLinearFactorY?: number;
  readonly customLinearFactorNeighborX?: number;
  readonly customLinearFactorNeighborY?: number;
  readonly customRadialScaleRadius?: number;
  readonly customRadialScaleIncrement?: number;
  readonly coherent3DOffset?: boolean;
  readonly hideGeometry?: boolean;
  readonly hideLastTransition?: boolean;
  readonly lockTextEntry?: boolean;
  readonly moveWith?: string;
  readonly useDefault?: boolean;
  readonly zOrderOffset?: number;
  readonly layoutTypeId?: string;
  readonly layoutCategoryId?: string;
  readonly quickStyleTypeId?: string;
  readonly quickStyleCategoryId?: string;
  readonly colorStyleTypeId?: string;
  readonly colorStyleCategoryId?: string;
}

/**
 * dgm:pt — SmartArt data model point (node) XML stringifier.
 */
export function stringifyPoint(
  modelId: string | number,
  text: string,
  type: string = "node",
  propertySet?: PointPropertySetOptions,
): string {
  const parts: string[] = [];
  parts.push(`<dgm:pt modelId="${escapeXml(String(modelId))}" type="${escapeXml(type)}">`);
  if (propertySet) {
    parts.push(stringifyPropertySet(propertySet));
  }
  parts.push(stringifyPointText(text));
  parts.push("</dgm:pt>");
  return parts.join("");
}

/**
 * dgm:pt — transition point (parTrans or sibTrans) XML stringifier.
 */
export function stringifyTransPoint(modelId: string, type: string, cxnId: string): string {
  return `<dgm:pt modelId="${escapeXml(modelId)}" type="${escapeXml(type)}" cxnId="${escapeXml(cxnId)}"><dgm:spPr/></dgm:pt>`;
}

function stringifyPropertySet(opts: PointPropertySetOptions): string {
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
  if (opts.customAngle !== undefined) attr.push(`custAng="${opts.customAngle}"`);
  if (opts.customFlipVertical !== undefined)
    attr.push(`custFlipVert="${opts.customFlipVertical ? 1 : 0}"`);
  if (opts.customFlipHorizontal !== undefined)
    attr.push(`custFlipHor="${opts.customFlipHorizontal ? 1 : 0}"`);
  if (opts.customSizeX !== undefined) attr.push(`custSzX="${opts.customSizeX}"`);
  if (opts.customSizeY !== undefined) attr.push(`custSzY="${opts.customSizeY}"`);
  if (opts.customScaleX !== undefined) attr.push(`custScaleX="${opts.customScaleX}"`);
  if (opts.customScaleY !== undefined) attr.push(`custScaleY="${opts.customScaleY}"`);
  if (opts.customText !== undefined) attr.push(`custT="${opts.customText ? 1 : 0}"`);
  if (opts.customLinearFactorX !== undefined)
    attr.push(`custLinFactX="${opts.customLinearFactorX}"`);
  if (opts.customLinearFactorY !== undefined)
    attr.push(`custLinFactY="${opts.customLinearFactorY}"`);
  if (opts.customLinearFactorNeighborX !== undefined)
    attr.push(`custLinFactNeighborX="${opts.customLinearFactorNeighborX}"`);
  if (opts.customLinearFactorNeighborY !== undefined)
    attr.push(`custLinFactNeighborY="${opts.customLinearFactorNeighborY}"`);
  if (opts.customRadialScaleRadius !== undefined)
    attr.push(`custRadScaleRad="${opts.customRadialScaleRadius}"`);
  if (opts.customRadialScaleIncrement !== undefined)
    attr.push(`custRadScaleInc="${opts.customRadialScaleIncrement}"`);
  if (opts.coherent3DOffset !== undefined)
    attr.push(`coherent3DOff="${opts.coherent3DOffset ? 1 : 0}"`);
  if (opts.hideGeometry !== undefined) attr.push(`hideGeom="${opts.hideGeometry ? 1 : 0}"`);
  if (opts.hideLastTransition !== undefined)
    attr.push(`hideLastTrans="${opts.hideLastTransition ? 1 : 0}"`);
  if (opts.lockTextEntry !== undefined) attr.push(`lkTxEntry="${opts.lockTextEntry ? 1 : 0}"`);
  if (opts.moveWith !== undefined) attr.push(`moveWith="${escapeXml(opts.moveWith)}"`);
  if (opts.useDefault !== undefined) attr.push(`useDef="${opts.useDefault ? 1 : 0}"`);
  if (opts.zOrderOffset !== undefined) attr.push(`zOrderOff="${opts.zOrderOffset}"`);
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

  if (attr.length === 0) return "";
  return `<dgm:prSet ${attr.join(" ")}/>`;
}

function stringifyPointText(text: string): string {
  const body = text ? `<a:r><a:t>${escapeXml(text)}</a:t></a:r>` : "";
  return `<dgm:t><a:bodyPr/><a:lstStyle/><a:p>${body}</a:p></dgm:t>`;
}
