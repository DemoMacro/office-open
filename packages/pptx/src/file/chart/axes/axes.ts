import { BuilderElement, XmlComponent } from "@file/xml-components";
import { chartAttr, wrapEl } from "@file/xml-components";

export class CatAx extends XmlComponent {
    public constructor(axId: number, crossAx: number) {
        super("c:catAx");
        this.root.push(wrapEl("c:axId", chartAttr({ val: axId })));
        this.root.push(new Scaling());
        this.root.push(wrapEl("c:delete", chartAttr({ val: 0 })));
        this.root.push(wrapEl("c:axPos", chartAttr({ val: "b" })));
        this.root.push(wrapEl("c:auto", chartAttr({ val: 1 })));
        this.root.push(wrapEl("c:lblOffset", chartAttr({ val: 100 })));
        this.root.push(wrapEl("c:noMultiLvlLbl", chartAttr({ val: 0 })));
        this.root.push(wrapEl("c:crossAx", chartAttr({ val: crossAx })));
        this.root.push(wrapEl("c:crosses", chartAttr({ val: "autoZero" })));
    }
}

class Scaling extends XmlComponent {
    public constructor() {
        super("c:scaling");
        this.root.push(wrapEl("c:orientation", chartAttr({ val: "minMax" })));
    }
}

export class ValAx extends XmlComponent {
    public constructor(axId: number, crossAx: number) {
        super("c:valAx");
        this.root.push(wrapEl("c:axId", chartAttr({ val: axId })));
        this.root.push(new Scaling());
        this.root.push(wrapEl("c:delete", chartAttr({ val: 0 })));
        this.root.push(wrapEl("c:axPos", chartAttr({ val: "l" })));
        this.root.push(
            wrapEl("c:numFmt", chartAttr({ formatCode: "General", sourceLinked: 1 })),
        );
        this.root.push(wrapEl("c:crossAx", chartAttr({ val: crossAx })));
        this.root.push(wrapEl("c:crosses", chartAttr({ val: "autoZero" })));
        this.root.push(new BuilderElement({ name: "c:spPr" }));
    }
}
