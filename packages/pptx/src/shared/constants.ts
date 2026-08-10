/**
 * PPTX-specific OOXML constants.
 *
 * Text/outline/shadow defaults live in core DrawingML; this module keeps only
 * PPTX-specific shape-tree and color-map constants.
 *
 * @module
 */

// ── Shape tree defaults (p:spTree) ──

/** Empty shape tree header: nvGrpSpPr + grpSpPr with zero-offset transform */
export const SP_TREE_HEADER =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

// ── Default color map (p:clrMap) ──

/** Default theme color mapping attribute string for <p:clrMap> */
export const DEFAULT_COLOR_MAP =
  'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';
