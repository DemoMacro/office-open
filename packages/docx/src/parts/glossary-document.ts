/**
 * Glossary document component — stores building block definitions.
 *
 * Generates word/glossary/document.xml containing Quick Parts entries
 * that appear in Word's Insert > Quick Parts gallery.
 *
 * @module
 */

import type { SectionChild } from "@shared/section";

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
  children: SectionChild[];
}

/** Glossary document options */
export interface GlossaryDocumentOptions {
  /** Building blocks */
  parts: DocPartOptions[];
}

// ── Descriptor ──

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, findChild } from "@office-open/xml";

import { parseParagraph } from "../body";
import type { BodyContext, DocxReadContext } from "../context";

const GLOSSARY_NS =
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
  'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

function glossaryEscapeAttr(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function docPartPrXml(part: GlossaryDocumentOptions["parts"][number]): string {
  const prParts: string[] = [];
  prParts.push(
    `<w:name w:val="${glossaryEscapeAttr(part.name)}"${part.decorated ? ' w:decorated="1"' : ""}/>`,
  );
  if (part.category || part.gallery) {
    const catParts: string[] = [];
    if (part.category) {
      catParts.push(`<w:name w:val="${glossaryEscapeAttr(part.category)}"/>`);
    }
    catParts.push(`<w:gallery w:val="${part.gallery}"/>`);
    prParts.push(`<w:category>${catParts.join("")}</w:category>`);
  }
  if (part.types && part.types.length > 0) {
    const typeXml = part.types.map((t) => `<w:type w:val="${t}"/>`).join("");
    const allAttr = part.allTypes ? ' w:all="1"' : "";
    prParts.push(`<w:types${allAttr}>${typeXml}</w:types>`);
  }
  if (part.behaviors && part.behaviors.length > 0) {
    const behaviorXml = part.behaviors.map((b) => `<w:behavior w:val="${b}"/>`).join("");
    prParts.push(`<w:behaviors>${behaviorXml}</w:behaviors>`);
  }
  if (part.description) {
    prParts.push(`<w:description w:val="${glossaryEscapeAttr(part.description)}"/>`);
  }
  if (part.guid) {
    prParts.push(`<w:guid w:val="${glossaryEscapeAttr(part.guid)}"/>`);
  }
  return `<w:docPartPr>${prParts.join("")}</w:docPartPr>`;
}

export const glossaryDesc: CustomDescriptor<GlossaryDocumentOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const partsXml = opts.parts
      .map((part) => {
        const bodyContent = ((part.children ?? []) as SectionChild[])
          .map((child) => ctx.stringifyChild(child, ctx))
          .join("");

        return `<w:docPart>${docPartPrXml(part)}<w:docPartBody>${bodyContent}</w:docPartBody></w:docPart>`;
      })
      .join("");

    return `<w:glossaryDocument ${GLOSSARY_NS}><w:docParts>${partsXml}</w:docParts></w:glossaryDocument>`;
  },

  parse(el, ctx) {
    const dctx = ctx as DocxReadContext;
    const parts: Record<string, unknown>[] = [];

    const docPartsEl = findChild(el, "w:docParts");
    if (!docPartsEl) return { parts } as unknown as GlossaryDocumentOptions;

    for (const docPart of docPartsEl.elements ?? []) {
      if (docPart.name !== "w:docPart") continue;
      const part: Record<string, unknown> = {};

      // Parse w:docPartPr
      const pr = findChild(docPart, "w:docPartPr");
      if (pr) {
        // name
        const name = findChild(pr, "w:name");
        if (name) {
          part.name = attr(name, "w:val") ?? "";
          const decorated = attr(name, "w:decorated");
          if (decorated === "1") part.decorated = true;
        }

        // category
        const category = findChild(pr, "w:category");
        if (category) {
          const catName = findChild(category, "w:name");
          if (catName) part.category = attr(catName, "w:val");
          const gallery = findChild(category, "w:gallery");
          if (gallery) part.gallery = attr(gallery, "w:val");
        }

        // types
        const types = findChild(pr, "w:types");
        if (types) {
          const typeList: string[] = [];
          for (const t of types.elements ?? []) {
            if (t.name === "w:type") {
              const val = attr(t, "w:val");
              if (val) typeList.push(val);
            }
          }
          if (typeList.length > 0) part.types = typeList;
          const allAttr = attr(types, "w:all");
          if (allAttr === "1") part.allTypes = true;
        }

        // behaviors
        const behaviors = findChild(pr, "w:behaviors");
        if (behaviors) {
          const behaviorList: string[] = [];
          for (const b of behaviors.elements ?? []) {
            if (b.name === "w:behavior") {
              const val = attr(b, "w:val");
              if (val) behaviorList.push(val);
            }
          }
          if (behaviorList.length > 0) part.behaviors = behaviorList;
        }

        // description
        const desc = findChild(pr, "w:description");
        if (desc) {
          const val = attr(desc, "w:val");
          if (val) part.description = val;
        }

        // guid
        const guid = findChild(pr, "w:guid");
        if (guid) {
          const val = attr(guid, "w:val");
          if (val) part.guid = val;
        }
      }

      // Parse w:docPartBody children
      const body = findChild(docPart, "w:docPartBody");
      if (body) {
        const childList: unknown[] = [];
        for (const sub of body.elements ?? []) {
          if (sub.name === "w:p") {
            childList.push({ paragraph: parseParagraph(sub, dctx) });
          }
        }
        if (childList.length > 0) part.children = childList;
      }

      parts.push(part);
    }

    return { parts } as unknown as GlossaryDocumentOptions;
  },
};
