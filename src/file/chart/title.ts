/**
 * Chart title (c:title).
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";
import { chartAttr } from "@file/xml-components";

/**
 * c:title — chart title overlay.
 */
export class ChartTitle extends XmlComponent {
    public constructor(title: string) {
        super("c:title");
        this.root.push(new TitleTx(title));
        this.root.push(new TitleOverlay());
    }
}

class TitleTx extends XmlComponent {
    public constructor(title: string) {
        super("c:tx");

        const rich = new (class extends XmlComponent {
            public constructor() {
                super("c:rich");
            }
        })();

        rich["root"].push(
            new (class extends XmlComponent {
                public constructor() {
                    super("a:bodyPr");
                }
            })(),
        );
        rich["root"].push(
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

        const r = new (class extends XmlComponent {
            public constructor() {
                super("a:r");
            }
        })();
        r["root"].push(
            new (class extends XmlComponent {
                public constructor() {
                    super("a:t");
                }
            })(title),
        );

        p["root"].push(r);
        rich["root"].push(p);
        this.root.push(rich);
    }
}

class TitleOverlay extends XmlComponent {
    public constructor() {
        super("c:overlay");
        this.root.push(chartAttr({ val: 0 }));
    }
}
