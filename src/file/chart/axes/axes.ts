/**
 * Chart axes — c:catAx and c:valAx.
 *
 * @module
 */
import { EmptyElement, XmlComponent } from "@file/xml-components";
import { chartAttr, wrapEl } from "@file/xml-components";

/**
 * c:catAx — category axis.
 */
export class CatAx extends XmlComponent {
    public constructor(axId: number, crossAx: number) {
        super("c:catAx");
        this.root.push(wrapEl("c:axId", chartAttr({ val: axId })));
        this.root.push(wrapEl("c:crossAx", chartAttr({ val: crossAx })));
        this.root.push(new EmptyElement("c:scaling"));
        this.root.push(new EmptyElement("c:delete"));
        this.root.push(wrapEl("c:axPos", chartAttr({ val: "b" })));
        this.root.push(wrapEl("c:auto", chartAttr({ val: true })));
        this.root.push(wrapEl("c:lblOffset", chartAttr({ val: "100" })));
        this.root.push(wrapEl("c:noMultiLvlLbl", chartAttr({ val: false })));
    }
}

/**
 * c:valAx — value axis.
 */
export class ValAx extends XmlComponent {
    public constructor(axId: number, crossAx: number) {
        super("c:valAx");
        this.root.push(wrapEl("c:axId", chartAttr({ val: axId })));
        this.root.push(wrapEl("c:crossAx", chartAttr({ val: crossAx })));
        this.root.push(new EmptyElement("c:scaling"));
        this.root.push(new EmptyElement("c:delete"));
        this.root.push(wrapEl("c:axPos", chartAttr({ val: "l" })));
        this.root.push(wrapEl("c:numFmt", chartAttr({ formatCode: "General", sourceLinked: true })));
        this.root.push(new EmptyElement("c:majorTickMark"));
        this.root.push(new EmptyElement("c:minorTickMark"));
        this.root.push(new EmptyElement("c:tickLblPos"));
        this.root.push(new EmptyElement("c:spPr"));
    }
}
