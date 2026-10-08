/**
 * PPTX-specific OOXML constants.
 *
 * Text/outline/shadow/color-mapping defaults live in core DrawingML; this
 * module keeps only PPTX-specific shape-tree constants.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";

// ── Shape tree defaults (p:spTree) ──

/** Empty shape tree header: nvGrpSpPr + grpSpPr with zero-offset transform */
export function stringifySpTreeHeader(
  id = 1,
  name = "",
  blackWhiteMode?: string,
  transformless?: boolean,
): string {
  return (
    `<p:nvGrpSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr${blackWhiteMode ? ` bwMode="${blackWhiteMode}"` : ""}>` +
    (transformless
      ? "</p:grpSpPr>"
      : '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>')
  );
}

export const SP_TREE_HEADER = stringifySpTreeHeader();
