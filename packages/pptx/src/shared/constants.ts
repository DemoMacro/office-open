/**
 * PPTX-specific OOXML constants.
 *
 * Text/outline/shadow/color-mapping defaults live in core DrawingML; this
 * module keeps only PPTX-specific shape-tree constants.
 *
 * @module
 */

// ── Shape tree defaults (p:spTree) ──

/** Empty shape tree header: nvGrpSpPr + grpSpPr with zero-offset transform */
export const SP_TREE_HEADER =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
