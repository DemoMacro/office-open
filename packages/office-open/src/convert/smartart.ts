/**
 * Cross-format SmartArt conversion (docx ↔ pptx).
 *
 * SmartArt (diagrams) is isomorphic between docx and pptx: both store the data
 * as a core TreeNode tree (docx wraps it in data.nodes; pptx takes nodes
 * directly), reference the same built-in layout/style/color by ID, and anchor
 * via an absolute EMU bounding box (pptx top-level x/y/w/h ↔ docx
 * MediaTransformation). xlsx has no diagram part, so it does not participate.
 *
 * Position maps through the shared position helpers; docx floating positioning
 * is dropped on the pptx leg (no equivalent) and pptx produces an inline
 * transformation (like a picture) on the docx leg. The cNvPr fields
 * (name/description/title/hidden) pass straight through via
 * pickNonVisualDrawingProperties so alt text survives a cross-format copy,
 * mirroring the picture converter.
 *
 * @module
 */
import { pickNonVisualDrawingProperties } from "@office-open/core";
import type { NonVisualDrawingPropertiesOptions } from "@office-open/core";
import type { TreeNode } from "@office-open/core/smartart";
import type { SmartArtOptions as DocxSmartArt } from "@office-open/docx";
import type { SmartArtOptions as PptxSmartArt } from "@office-open/pptx";

import { boxFromDocx, boxFromPptx, boxToPptx } from "./position";

// SmartArtNode (docx) and TreeNode (core) are structurally identical trees.
// Map recursively since docx children are mutable and core's are readonly.
const toTreeNodes = (nodes: DocxSmartArt["data"]["nodes"]): TreeNode[] =>
  nodes.map((n) => ({
    text: n.text,
    ...(n.children ? { children: toTreeNodes(n.children) } : {}),
  }));

const toDocxNodes = (nodes: TreeNode[]): DocxSmartArt["data"]["nodes"] =>
  nodes.map((n) => ({
    text: n.text,
    ...(n.children ? { children: toDocxNodes([...n.children]) } : {}),
  }));

/**
 * Build the docx altText (wp:docPr) from the shared cNvPr. Only emitted when at
 * least one cNvPr field is authored; name defaults to "SmartArt" since docx
 * requires it. Structurally compatible with docx's DocPropertiesOptions.
 */
const altTextFromCnvPr = (
  cNvPr: NonVisualDrawingPropertiesOptions,
): { altText?: NonVisualDrawingPropertiesOptions & { name: string } } => {
  const picked = pickNonVisualDrawingProperties(cNvPr);
  if (
    picked.name === undefined &&
    picked.description === undefined &&
    picked.title === undefined &&
    picked.hidden === undefined
  ) {
    return {};
  }
  return { altText: { name: picked.name ?? "SmartArt", ...picked } };
};

// ── → docx ──

/** Convert a pptx SmartArt to a docx inline diagram. */
export function toDocxSmartArt(source: PptxSmartArt): DocxSmartArt {
  const box = boxFromPptx(source.x, source.y, source.width, source.height);
  return {
    data: { nodes: toDocxNodes(source.nodes) },
    transformation: {
      width: box.width,
      height: box.height,
      ...(source.x !== undefined || source.y !== undefined
        ? { offset: { left: box.x, top: box.y } }
        : {}),
    },
    ...altTextFromCnvPr(source),
    ...(source.layout ? { layout: source.layout } : {}),
    ...(source.style ? { style: source.style } : {}),
    ...(source.color ? { color: source.color } : {}),
  };
}

// ── → pptx ──

/** Convert a docx SmartArt to a pptx diagram (floating positioning is dropped). */
export function toPptxSmartArt(source: DocxSmartArt): PptxSmartArt {
  const box = boxFromDocx(source.transformation);
  const pos = boxToPptx(box);
  return {
    nodes: toTreeNodes(source.data.nodes),
    x: pos.x,
    y: pos.y,
    width: pos.width,
    height: pos.height,
    ...pickNonVisualDrawingProperties(source.altText),
    ...(source.layout ? { layout: source.layout } : {}),
    ...(source.style ? { style: source.style } : {}),
    ...(source.color ? { color: source.color } : {}),
  };
}
