/**
 * Textbox module for WordprocessingML documents.
 *
 * This module provides support for text boxes using VML shapes. Textboxes allow for
 * creating floating text containers with custom positioning and styling.
 *
 * The canonical type is defined inline in SectionChild (section-child.ts).
 *
 * @module
 */
import type { ParagraphOptions } from "@parts/paragraph";
import type { SectionChild } from "@shared/section";

import type { VmlShapeStyle } from "./shape/shape";

/**
 * Options for creating a Textbox.
 *
 * Extends paragraph options while replacing the style property with VML shape styling.
 */
export type TextboxOptions = Omit<ParagraphOptions, "style" | "children"> & {
  /** VML shape style properties for the textbox (positioning, sizing, wrapping, etc.) */
  style?: VmlShapeStyle;
  /** Array of block-level content elements (paragraphs, tables, etc.) */
  children?: SectionChild[];
};
