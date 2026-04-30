/**
 * dgm:pt — SmartArt data model point (node).
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";
import { chartAttr } from "@file/xml-components";

/**
 * CT_Pt — a single point in the data model.
 */
export class Point extends XmlComponent {
    public constructor(modelId: number, text: string, type: string = "node") {
        super("dgm:pt");
        this.root.push(chartAttr({ modelId, type }));
        this.root.push(new PointText(text));
    }
}

/**
 * dgm:t — text body within a point.
 *
 * Per XSD, dgm:t has type CT_TextBody, so bodyPr/lstStyle/p
 * are direct children (no a:txBody wrapper).
 */
class PointText extends XmlComponent {
    public constructor(text: string) {
        super("dgm:t");

        // a:bodyPr, a:lstStyle, a:p are direct children of dgm:t
        this.root.push(
            new (class extends XmlComponent {
                public constructor() {
                    super("a:bodyPr");
                }
            })(),
        );
        this.root.push(
            new (class extends XmlComponent {
                public constructor() {
                    super("a:lstStyle");
                }
            })(),
        );

        const p = new (class extends XmlComponent {
            public constructor() {
                super("a:p");
            }
        })();

        if (text) {
            const r = new (class extends XmlComponent {
                public constructor() {
                    super("a:r");
                }
            })();
            const t = new (class extends XmlComponent {
                public constructor() {
                    super("a:t");
                }
            })();
            t["root"].push(text);
            r["root"].push(t);
            p["root"].push(r);
        }

        this.root.push(p);
    }
}
