import { XmlComponent, chartAttr } from "../../xml-components";

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
 * dgm:pt — SmartArt data model point (node).
 */
export class Point extends XmlComponent {
  public constructor(
    modelId: string,
    text: string,
    type: string = "node",
    propertySet?: PointPropertySetOptions,
  ) {
    super("dgm:pt");
    this.root.push(chartAttr({ modelId, type }));
    if (propertySet) {
      this.root.push(new PropertySet(propertySet));
    }
    this.root.push(new PointText(text));
  }
}

/**
 * dgm:prSet — Property set for a SmartArt data model point.
 */
class PropertySet extends XmlComponent {
  public constructor(opts: PointPropertySetOptions) {
    super("dgm:prSet");
    const attr: Record<string, string | number> = {};
    if (opts.presentationAssociationId !== undefined)
      attr.presAssocID = opts.presentationAssociationId;
    if (opts.presentationName !== undefined) attr.presName = opts.presentationName;
    if (opts.presentationStyleLabel !== undefined) attr.presStyleLbl = opts.presentationStyleLabel;
    if (opts.presentationStyleIndex !== undefined) attr.presStyleIdx = opts.presentationStyleIndex;
    if (opts.presentationStyleCount !== undefined) attr.presStyleCnt = opts.presentationStyleCount;
    if (opts.placeholderText !== undefined) attr.phldrT = opts.placeholderText;
    if (opts.placeholder !== undefined) attr.phldr = opts.placeholder ? "1" : "0";
    if (opts.customAngle !== undefined) attr.custAng = opts.customAngle;
    if (opts.customFlipVertical !== undefined)
      attr.custFlipVert = opts.customFlipVertical ? "1" : "0";
    if (opts.customFlipHorizontal !== undefined)
      attr.custFlipHor = opts.customFlipHorizontal ? "1" : "0";
    if (opts.customSizeX !== undefined) attr.custSzX = opts.customSizeX;
    if (opts.customSizeY !== undefined) attr.custSzY = opts.customSizeY;
    if (opts.customScaleX !== undefined) attr.custScaleX = opts.customScaleX;
    if (opts.customScaleY !== undefined) attr.custScaleY = opts.customScaleY;
    if (opts.customText !== undefined) attr.custT = opts.customText ? "1" : "0";
    if (opts.customLinearFactorX !== undefined) attr.custLinFactX = opts.customLinearFactorX;
    if (opts.customLinearFactorY !== undefined) attr.custLinFactY = opts.customLinearFactorY;
    if (opts.customLinearFactorNeighborX !== undefined)
      attr.custLinFactNeighborX = opts.customLinearFactorNeighborX;
    if (opts.customLinearFactorNeighborY !== undefined)
      attr.custLinFactNeighborY = opts.customLinearFactorNeighborY;
    if (opts.customRadialScaleRadius !== undefined)
      attr.custRadScaleRad = opts.customRadialScaleRadius;
    if (opts.customRadialScaleIncrement !== undefined)
      attr.custRadScaleInc = opts.customRadialScaleIncrement;
    if (opts.coherent3DOffset !== undefined) attr.coherent3DOff = opts.coherent3DOffset ? "1" : "0";
    if (opts.hideGeometry !== undefined) attr.hideGeom = opts.hideGeometry ? "1" : "0";
    if (opts.hideLastTransition !== undefined)
      attr.hideLastTrans = opts.hideLastTransition ? "1" : "0";
    if (opts.lockTextEntry !== undefined) attr.lkTxEntry = opts.lockTextEntry ? "1" : "0";
    if (opts.moveWith !== undefined) attr.moveWith = opts.moveWith;
    if (opts.useDefault !== undefined) attr.useDef = opts.useDefault ? "1" : "0";
    if (opts.zOrderOffset !== undefined) attr.zOrderOff = opts.zOrderOffset;
    if (opts.layoutTypeId !== undefined) attr.loTypeId = opts.layoutTypeId;
    if (opts.layoutCategoryId !== undefined) attr.loCatId = opts.layoutCategoryId;
    if (opts.quickStyleTypeId !== undefined) attr.qsTypeId = opts.quickStyleTypeId;
    if (opts.quickStyleCategoryId !== undefined) attr.qsCatId = opts.quickStyleCategoryId;
    if (opts.colorStyleTypeId !== undefined) attr.csTypeId = opts.colorStyleTypeId;
    if (opts.colorStyleCategoryId !== undefined) attr.csCatId = opts.colorStyleCategoryId;
    if (Object.keys(attr).length > 0) this.root.push({ _attr: attr });
  }
}

/**
 * Transition point (parTrans or sibTrans) — no text body, references a connection.
 */
export class TransPoint extends XmlComponent {
  public constructor(modelId: string, type: string, cxnId: string) {
    super("dgm:pt");
    this.root.push(chartAttr({ modelId, type, cxnId }));
    this.root.push(new EmptyElement("dgm:spPr"));
  }
}

class EmptyElement extends XmlComponent {
  public constructor(tag: string) {
    super(tag);
  }
}

/**
 * dgm:t — text body within a point.
 */
class PointText extends XmlComponent {
  public constructor(text: string) {
    super("dgm:t");

    this.root.push(new EmptyElement("a:bodyPr"));
    this.root.push(new EmptyElement("a:lstStyle"));

    const p = new EmptyElement("a:p");
    if (text) {
      const t = new (class extends XmlComponent {
        public constructor() {
          super("a:t");
        }
      })();
      t["root"].push(text);

      const r = new (class extends XmlComponent {
        public constructor() {
          super("a:r");
        }
      })();
      r["root"].push(t);

      p["root"].push(r);
    }
    this.root.push(p);
  }
}
