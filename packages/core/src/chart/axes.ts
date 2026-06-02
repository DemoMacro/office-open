import { BuilderElement, XmlComponent, chartAttr, wrapEl } from "../xml-components";

/**
 * c:scaling — axis scaling configuration (minOccurs=1 in EG_AxShared).
 */
class Scaling extends XmlComponent {
  public constructor() {
    super("c:scaling");
    this.root.push(wrapEl("c:orientation", chartAttr({ val: "minMax" })));
  }
}

/**
 * c:catAx — category axis.
 *
 * XSD sequence: EG_AxShared (axId, scaling, delete, axPos, ..., crossAx, crosses)
 * then: auto, lblAlgn, lblOffset, tickLblSkip, tickMarkSkip, noMultiLvlLbl
 */
export class CatAx extends XmlComponent {
  public constructor(axId: number, crossAx: number) {
    super("c:catAx");
    // EG_AxShared
    this.root.push(wrapEl("c:axId", chartAttr({ val: axId })));
    this.root.push(new Scaling());
    this.root.push(wrapEl("c:delete", chartAttr({ val: 0 })));
    this.root.push(wrapEl("c:axPos", chartAttr({ val: "b" })));
    this.root.push(wrapEl("c:crossAx", chartAttr({ val: crossAx })));
    this.root.push(wrapEl("c:crosses", chartAttr({ val: "autoZero" })));
    // CatAx-specific (after EG_AxShared)
    this.root.push(wrapEl("c:auto", chartAttr({ val: 1 })));
    this.root.push(wrapEl("c:lblOffset", chartAttr({ val: 100 })));
    this.root.push(wrapEl("c:noMultiLvlLbl", chartAttr({ val: 0 })));
  }
}

/**
 * c:valAx — value axis.
 *
 * XSD sequence: EG_AxShared (axId, scaling, delete, axPos, ..., spPr, ..., crossAx, crosses)
 * then: crossBetween, majorUnit, minorUnit, dispUnits
 */
export class ValAx extends XmlComponent {
  public constructor(axId: number, crossAx: number) {
    super("c:valAx");
    // EG_AxShared
    this.root.push(wrapEl("c:axId", chartAttr({ val: axId })));
    this.root.push(new Scaling());
    this.root.push(wrapEl("c:delete", chartAttr({ val: 0 })));
    this.root.push(wrapEl("c:axPos", chartAttr({ val: "l" })));
    this.root.push(wrapEl("c:numFmt", chartAttr({ formatCode: "General", sourceLinked: 1 })));
    this.root.push(new BuilderElement({ name: "c:spPr" }));
    this.root.push(wrapEl("c:crossAx", chartAttr({ val: crossAx })));
    this.root.push(wrapEl("c:crosses", chartAttr({ val: "autoZero" })));
  }
}

export class SerAx extends XmlComponent {
  public constructor(axId: number, crossAx: number) {
    super("c:serAx");
    this.root.push(wrapEl("c:axId", chartAttr({ val: axId })));
    this.root.push(new Scaling());
    this.root.push(wrapEl("c:delete", chartAttr({ val: 0 })));
    this.root.push(wrapEl("c:axPos", chartAttr({ val: "b" })));
    this.root.push(wrapEl("c:numFmt", chartAttr({ formatCode: "General", sourceLinked: 1 })));
    this.root.push(new BuilderElement({ name: "c:spPr" }));
    this.root.push(wrapEl("c:crossAx", chartAttr({ val: crossAx })));
    this.root.push(wrapEl("c:crosses", chartAttr({ val: "autoZero" })));
    this.root.push(wrapEl("c:tickLblSkip", chartAttr({ val: 1 })));
  }
}
