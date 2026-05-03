import { XmlComponent, chartAttr } from "../../xml-components";

import { Connection } from "./connection";

/**
 * CT_DataModel — the complete data model for a SmartArt diagram.
 */
export class DataModel extends XmlComponent {
    public constructor(points: readonly XmlComponent[], connections: readonly Connection[]) {
        super("dgm:dataModel");

        this.root.push(
            chartAttr({
                "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                "xmlns:dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            }),
        );

        const ptLst = new (class extends XmlComponent {
            public constructor() {
                super("dgm:ptLst");
            }
        })();
        for (const pt of points) {
            ptLst["root"].push(pt);
        }
        this.root.push(ptLst);

        const cxnLst = new (class extends XmlComponent {
            public constructor() {
                super("dgm:cxnLst");
            }
        })();
        for (const cxn of connections) {
            cxnLst["root"].push(cxn);
        }
        this.root.push(cxnLst);

        this.root.push(new EmptyElement("dgm:bg"));
        this.root.push(new EmptyElement("dgm:whole"));
    }
}

class EmptyElement extends XmlComponent {
    public constructor(tag: string) {
        super(tag);
    }
}
