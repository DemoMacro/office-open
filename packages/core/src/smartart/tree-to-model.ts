import { escapeXml } from "@office-open/xml";

import { COLOR_CATEGORIES, LAYOUT_CATEGORIES, STYLE_CATEGORIES } from "./categories";
import { stringifyConnection } from "./data-model/connection";
import { stringifyDataModel } from "./data-model/data-model";
import { stringifyPoint, stringifyTransPoint } from "./data-model/point";

export interface TreeNode {
  readonly text: string;
  readonly children?: readonly TreeNode[];
}

function stringifyDocPoint(layout: string, style: string, color: string): string {
  const loTypeId = `urn:microsoft.com/office/officeart/2005/8/layout/${layout}`;
  const loCatId = LAYOUT_CATEGORIES[layout] ?? "list";
  const qsTypeId = `urn:microsoft.com/office/officeart/2005/8/quickstyle/${style}`;
  const qsCatId = STYLE_CATEGORIES[style] ?? "simple";
  const csTypeId = `urn:microsoft.com/office/officeart/2005/8/colors/${color}`;
  const csCatId = COLOR_CATEGORIES[color] ?? "accent1";

  return [
    '<dgm:pt modelId="0" type="doc">',
    `<dgm:prSet loTypeId="${escapeXml(loTypeId)}" loCatId="${escapeXml(loCatId)}" qsTypeId="${escapeXml(qsTypeId)}" qsCatId="${escapeXml(qsCatId)}" csTypeId="${escapeXml(csTypeId)}" csCatId="${escapeXml(csCatId)}" phldr="0"/>`,
    "<dgm:spPr/>",
    "<dgm:t><a:bodyPr/><a:lstStyle/><a:p/></dgm:t>",
    "</dgm:pt>",
  ].join("");
}

function uuid(): string {
  return `{${crypto.randomUUID().toUpperCase()}}`;
}

/**
 * Creates SmartArt data model XML from tree nodes with layout/style/color settings.
 * Returns the XML string body (no XML declaration).
 */
export const createDataModel = (
  nodes: readonly TreeNode[],
  layout: string = "default",
  style: string = "simple1",
  color: string = "accent1_2",
): string => {
  const pointStrs: string[] = [];
  const connectionStrs: string[] = [];

  pointStrs.push(stringifyDocPoint(layout, style, color));

  for (let i = 0; i < nodes.length; i++) {
    const walk = (node: TreeNode, parentUuid: string, srcOrd: number): void => {
      const nodeUuid = uuid();
      const parTransUuid = uuid();
      const sibTransUuid = uuid();
      const cxnUuid = uuid();

      pointStrs.push(stringifyTransPoint(parTransUuid, "parTrans", cxnUuid));
      pointStrs.push(stringifyTransPoint(sibTransUuid, "sibTrans", cxnUuid));
      pointStrs.push(stringifyPoint(nodeUuid, node.text));

      connectionStrs.push(
        stringifyConnection({
          modelId: parTransUuid,
          srcId: parentUuid,
          destId: nodeUuid,
          srcOrd,
          destOrd: 0,
          parTransId: parTransUuid,
          sibTransId: sibTransUuid,
        }),
      );

      if (node.children) {
        for (let j = 0; j < node.children.length; j++) {
          walk(node.children[j], nodeUuid, j);
        }
      }
    };
    walk(nodes[i], "0", i);
  }

  return stringifyDataModel(pointStrs, connectionStrs);
};
