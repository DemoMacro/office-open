/**
 * Glossary document component — stores building block definitions.
 *
 * Generates word/glossary/document.xml containing Quick Parts entries
 * that appear in Word's Insert > Quick Parts gallery.
 *
 * @module
 */

import type { FileChild } from "@file/file-child";

/** Gallery type for building blocks (ST_DocPartGallery) */
export const DocPartGallery = {
  PLACEHOLDER: "placeholder",
  COVER_PAGE: "coverPage",
  BUILT_IN: "builtIn",
  CUSTOM1: "custom1",
  CUSTOM2: "custom2",
  CUSTOM3: "custom3",
  CUSTOM4: "custom4",
  CUSTOM5: "custom5",
  AUTO_TEXT: "autoTxt",
  TEXT_BOX: "txtBox",
  PAGE_NUMBERS_TOP: "pgNumTop",
  PAGE_NUMBERS_BOTTOM: "pgNumBottom",
  PAGE_NUMBERS_MARGIN: "pgNumMargin",
  TABLES: "tbls",
  HEADERS: "hdrs",
  FOOTERS: "ftrs",
  WATERMARKS: "watermarks",
} as const;

export type DocPartGallery = (typeof DocPartGallery)[keyof typeof DocPartGallery];

/** Building block type (ST_DocPartType) */
export const DocPartType = {
  NONE: "none",
  NORMAL: "normal",
  AUTO_EXPAND: "autoExp",
  FORM_FIELD: "formField",
  BARCODE: "barCode",
} as const;

export type DocPartType = (typeof DocPartType)[keyof typeof DocPartType];

/** Building block behavior (ST_DocPartBehavior) */
export const DocPartBehavior = {
  CONTENT: "content",
  PARAGRAPH: "p",
  PAGE: "pg",
  SECTION: "sect",
} as const;

export type DocPartBehavior = (typeof DocPartBehavior)[keyof typeof DocPartBehavior];

/** A single building block (CT_DocPart) */
export interface DocPartOptions {
  /** Building block name (required) */
  readonly name: string;
  /** Gallery category (required) */
  readonly gallery: DocPartGallery;
  /** Category name within the gallery */
  readonly category?: string;
  /** Building block types */
  readonly types?: readonly DocPartType[];
  /** Insertion behaviors */
  readonly behaviors?: readonly DocPartBehavior[];
  /** Description */
  readonly description?: string;
  /** GUID for this building block */
  readonly guid?: string;
  /** Whether the name is decorated (built-in) */
  readonly decorated?: boolean;
  /** Body content — paragraphs, tables, etc. */
  readonly children: readonly FileChild[];
}

/** Glossary document options */
export interface GlossaryDocumentOptions {
  /** Building blocks */
  readonly parts: readonly DocPartOptions[];
}
