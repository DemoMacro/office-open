/**
 * Converts a tree-shaped API into a flat data model (points + connections).
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";
import { chartAttr } from "@file/xml-components";

import { COLOR_CATEGORIES, LAYOUT_CATEGORIES, STYLE_CATEGORIES } from "./built-in-definitions";
import { Connection } from "./data-model/connection";
import { DataModel } from "./data-model/data-model";
import { Point, TransPoint } from "./data-model/point";

export interface ITreeNode {
    readonly text: string;
    readonly children?: readonly ITreeNode[];
}

function createDocPoint(layout: string, style: string, color: string): XmlComponent {
    const pt = new (class extends XmlComponent {
        public constructor() {
            super("dgm:pt");
        }
    })();
    pt["root"].push(chartAttr({ modelId: 0, type: "doc" }));

    const prSet = new (class extends XmlComponent {
        public constructor() {
            super("dgm:prSet");
        }
    })();
    prSet["root"].push(
        chartAttr({
            loTypeId: `urn:microsoft.com/office/officeart/2005/8/layout/${layout}`,
            loCatId: LAYOUT_CATEGORIES[layout] ?? "list",
            qsTypeId: `urn:microsoft.com/office/officeart/2005/8/quickstyle/${style}`,
            qsCatId: STYLE_CATEGORIES[style] ?? "simple",
            csTypeId: `urn:microsoft.com/office/officeart/2005/8/colors/${color}`,
            csCatId: COLOR_CATEGORIES[color] ?? "accent1",
            phldr: "0",
        }),
    );
    pt["root"].push(prSet);
    pt["root"].push(new EmptyElement("dgm:spPr"));
    pt["root"].push(createEmptyTextBody());

    return pt;
}

function createEmptyTextBody(): XmlComponent {
    const t = new (class extends XmlComponent {
        public constructor() {
            super("dgm:t");
        }
    })();
    t["root"].push(new EmptyElement("a:bodyPr"));
    t["root"].push(new EmptyElement("a:lstStyle"));
    t["root"].push(new EmptyElement("a:p"));
    return t;
}

class EmptyElement extends XmlComponent {
    public constructor(tag: string) {
        super(tag);
    }
}

function uuid(): string {
    return `{${crypto.randomUUID().toUpperCase()}}`;
}

/**
 * Creates a DataModel from tree nodes with layout/style/color settings.
 */
export const createDataModel = (
    nodes: readonly ITreeNode[],
    layout: string = "default",
    style: string = "simple1",
    color: string = "accent1_2",
): DataModel => {
    const points: XmlComponent[] = [];
    const connections: Connection[] = [];

    points.push(createDocPoint(layout, style, color));

    for (let i = 0; i < nodes.length; i++) {
        const walk = (node: ITreeNode, parentUuid: string, srcOrd: number): void => {
            const nodeUuid = uuid();
            const parTransUuid = uuid();
            const sibTransUuid = uuid();
            const cxnUuid = uuid();

            points.push(new TransPoint(parTransUuid, "parTrans", cxnUuid));
            points.push(new TransPoint(sibTransUuid, "sibTrans", cxnUuid));
            points.push(new Point(nodeUuid, node.text));

            connections.push(
                new Connection(
                    parTransUuid,
                    parentUuid,
                    nodeUuid,
                    undefined,
                    srcOrd,
                    0,
                    parTransUuid,
                    sibTransUuid,
                ),
            );

            if (node.children) {
                for (let j = 0; j < node.children.length; j++) {
                    walk(node.children[j], nodeUuid, j);
                }
            }
        };
        walk(nodes[i], "0", i);
    }

    return new DataModel(points, connections);
};
