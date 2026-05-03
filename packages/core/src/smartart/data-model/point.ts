import { XmlComponent, chartAttr } from "../../xml-components";

/**
 * dgm:pt — SmartArt data model point (node).
 */
export class Point extends XmlComponent {
    public constructor(modelId: string, text: string, type: string = "node") {
        super("dgm:pt");
        this.root.push(chartAttr({ modelId, type }));
        this.root.push(new PointText(text));
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
