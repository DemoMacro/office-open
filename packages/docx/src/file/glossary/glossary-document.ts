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
  DEFAULT: "default",
  DOC_PARTS: "docParts",
  COVER_PAGE: "coverPg",
  EQUATIONS: "eq",
  FOOTERS: "ftrs",
  HEADERS: "hdrs",
  PAGE_NUMBERS: "pgNum",
  TABLES: "tbls",
  WATERMARKS: "watermarks",
  AUTO_TEXT: "autoTxt",
  TEXT_BOX: "txtBox",
  PAGE_NUMBERS_TOP: "pgNumT",
  PAGE_NUMBERS_BOTTOM: "pgNumB",
  PAGE_NUMBERS_MARGIN: "pgNumMargins",
  TABLE_OF_CONTENTS: "tblOfContents",
  BIBLIOGRAPHY: "bib",
  CUSTOM_QUICK_PARTS: "custQuickParts",
  CUSTOM_COVER_PAGE: "custCoverPg",
  CUSTOM_EQUATIONS: "custEq",
  CUSTOM_FOOTERS: "custFtrs",
  CUSTOM_HEADERS: "custHdrs",
  CUSTOM_PAGE_NUMBERS: "custPgNum",
  CUSTOM_TABLES: "custTbls",
  CUSTOM_WATERMARKS: "custWatermarks",
  CUSTOM_AUTO_TEXT: "custAutoTxt",
  CUSTOM_TEXT_BOX: "custTxtBox",
  CUSTOM_PAGE_NUMBERS_TOP: "custPgNumT",
  CUSTOM_PAGE_NUMBERS_BOTTOM: "custPgNumB",
  CUSTOM_PAGE_NUMBERS_MARGIN: "custPgNumMargins",
  CUSTOM_TABLE_OF_CONTENTS: "custTblOfContents",
  CUSTOM_BIBLIOGRAPHY: "custBib",
  CUSTOM1: "custom1",
  CUSTOM2: "custom2",
  CUSTOM3: "custom3",
  CUSTOM4: "custom4",
  CUSTOM5: "custom5",
} as const;

export type DocPartGallery = (typeof DocPartGallery)[keyof typeof DocPartGallery];

/** Building block type (ST_DocPartType) */
export const DocPartType = {
  NONE: "none",
  NORMAL: "normal",
  AUTO_EXPAND: "autoExp",
  TOOLBAR: "toolbar",
  SPELLER: "speller",
  FORM_FIELD: "formFld",
  BUILDING_BLOCK_PLACEHOLDER: "bbPlcHdr",
} as const;

export type DocPartType = (typeof DocPartType)[keyof typeof DocPartType];

/** Building block behavior (ST_DocPartBehavior) */
export const DocPartBehavior = {
  CONTENT: "content",
  PARAGRAPH: "p",
  PAGE: "pg",
} as const;

export type DocPartBehavior = (typeof DocPartBehavior)[keyof typeof DocPartBehavior];

/** A single building block (CT_DocPart) */
export interface DocPartOptions {
  /** Building block name (required) */
  name: string;
  /** Gallery category (required) */
  gallery: DocPartGallery;
  /** Category name within the gallery */
  category?: string;
  /** Building block types */
  types?: DocPartType[];
  /** Whether all building block types are included (w:all attribute) */
  allTypes?: boolean;
  /** Insertion behaviors */
  behaviors?: DocPartBehavior[];
  /** Description */
  description?: string;
  /** GUID for this building block */
  guid?: string;
  /** Whether the name is decorated (built-in) */
  decorated?: boolean;
  /** Body content — paragraphs, tables, etc. */
  children: FileChild[];
}

/** Glossary document options */
export interface GlossaryDocumentOptions {
  /** Building blocks */
  parts: DocPartOptions[];
}
