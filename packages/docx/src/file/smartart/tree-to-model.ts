/**
 * Converts a tree-shaped API into a flat data model (points + connections).
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";
import { chartAttr } from "@file/xml-components";

import { Connection } from "./data-model/connection";
import { DataModel } from "./data-model/data-model";
import { Point } from "./data-model/point";

export interface ITreeNode {
    readonly text: string;
    readonly children?: readonly ITreeNode[];
}

interface IFlatResult {
    readonly points: XmlComponent[];
    readonly connections: Connection[];
}

/**
 * Layout/style/color uniqueId URIs for dgm:prSet on the doc root.
 */
const DEFAULT_LAYOUT_ID = "urn:microsoft.com/office/officeart/2005/8/layout/default";
const DEFAULT_STYLE_ID = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1";
const DEFAULT_COLOR_ID = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2";

/**
 * Creates the doc root point (type="doc") with layout/style/color prSet.
 *
 * Per XSD CT_Pt sequence: prSet, spPr, t, extLst
 * CT_ElemPropSet is attributes-only (no child elements).
 */
function createDocPoint(): XmlComponent {
    const pt = new (class extends XmlComponent {
        public constructor() {
            super("dgm:pt");
        }
    })();
    pt["root"].push(chartAttr({ modelId: 0, type: "doc" }));

    // dgm:prSet — self-closing, attributes only
    const prSet = new (class extends XmlComponent {
        public constructor() {
            super("dgm:prSet");
        }
    })();
    prSet["root"].push(
        chartAttr({
            loTypeId: DEFAULT_LAYOUT_ID,
            loCatId: "list",
            qsTypeId: DEFAULT_STYLE_ID,
            qsCatId: "simple",
            csTypeId: DEFAULT_COLOR_ID,
            csCatId: "accent1",
            phldr: "0",
        }),
    );
    pt["root"].push(prSet);

    // dgm:spPr — empty
    pt["root"].push(new EmptyElement("dgm:spPr"));

    // dgm:t — empty text body (placeholder)
    pt["root"].push(createEmptyTextBody());

    return pt;
}

/**
 * Creates a minimal dgm:t with empty text body.
 * Per XSD, dgm:t is CT_TextBody — bodyPr/lstStyle/p are direct children.
 */
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

/**
 * Helper for empty self-closing XML elements.
 */
class EmptyElement extends XmlComponent {
    public constructor(tag: string) {
        super(tag);
    }
}

/**
 * Converts a tree of nodes into flat points and connections.
 *
 * The first point is a "doc" root (type="doc") with layout/style/color prSet.
 * Subsequent points are regular "node" types connected via parOf.
 */
export const treeToModel = (nodes: readonly ITreeNode[]): IFlatResult => {
    const points: XmlComponent[] = [];
    const connections: Connection[] = [];
    let nextModelId = 1; // 0 reserved for doc root
    let nextCxnId = 100;

    // Create root doc point
    points.push(createDocPoint());

    // Walk tree and connect each top-level node to the doc root
    for (let i = 0; i < nodes.length; i++) {
        const walk = (node: ITreeNode, parentId: number, srcOrd: number): void => {
            const modelId = nextModelId++;
            points.push(new Point(modelId, node.text));

            connections.push(new Connection(nextCxnId++, parentId, modelId, "parOf", srcOrd));

            if (node.children) {
                for (let j = 0; j < node.children.length; j++) {
                    walk(node.children[j], modelId, j);
                }
            }
        };
        walk(nodes[i], 0, i);
    }

    return { connections, points };
};

/**
 * Creates a DataModel from tree nodes.
 */
export const createDataModel = (nodes: readonly ITreeNode[]): DataModel => {
    const { connections, points } = treeToModel(nodes);
    return new DataModel(points, connections);
};
