/**
 * Glossary document component — stores building block definitions.
 *
 * Generates word/glossary/document.xml containing Quick Parts entries
 * that appear in Word's Insert > Quick Parts gallery.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import {
  parseSectionPropertiesEl,
  sectionPropertiesDesc,
} from "@parts/document/body/section-properties/descriptor";
import type { SectionPropertiesOptions } from "@parts/document/body/section-properties/section-properties";
import { documentNamespaceAttributes } from "@parts/document/document-attributes";
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

/** A section within a building block body (CT_Body section boundary). */
export interface DocPartSectionOptions {
  /** Block-level content in this section. */
  children: SectionChild[];
  /** Section properties carried by w:pPr/w:sectPr or terminal w:sectPr. */
  properties?: SectionPropertiesOptions;
}

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
  description?: string;
  /** GUID for this building block */
  guid?: string;
  /** Whether the name is decorated (built-in) */
  decorated?: boolean;
  /** Style applied to this building block */
  style?: string;
  /** Body sections, including section-break properties. */
  sections: DocPartSectionOptions[];
}

/** Glossary document options */
export interface GlossaryDocumentOptions {
  /** Building blocks */
  parts: DocPartOptions[];
}

// ── Descriptor ──

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { stringifyBodyChild } from "../body";
import type { BodyContext, DocxReadContext } from "../context";
import { parseSectionChild } from "../parse/body";

const GLOSSARY_NS = documentNamespaceAttributes([
  "wpc",
  "mc",
  "o",
  "r",
  "m",
  "v",
  "wp",
  "w10",
  "w",
  "w14",
  "w15",
  "wpg",
  "wpi",
  "wne",
  "wps",
]);

function stringifyDocPartBody(part: DocPartOptions, ctx: BodyContext): string {
  const parts: string[] = [];
  for (let sectionIndex = 0; sectionIndex < part.sections.length; sectionIndex++) {
    const section = part.sections[sectionIndex]!;
    const sectPrXml = section.properties
      ? (sectionPropertiesDesc.stringify(section.properties, ctx) ?? "")
      : "";
    const isLast = sectionIndex === part.sections.length - 1;
    let sectPrHosted = isLast || !sectPrXml;

    for (let childIndex = 0; childIndex < section.children.length; childIndex++) {
      const child = section.children[childIndex]!;
      const inject =
        !isLast &&
        !!sectPrXml &&
        childIndex === section.children.length - 1 &&
        ("paragraph" in child || "toc" in child);
      if (inject) sectPrHosted = true;
      parts.push(stringifyBodyChild(child, ctx, inject ? sectPrXml : undefined));
    }
    if (!isLast && sectPrXml && !sectPrHosted) {
      parts.push(`<w:p><w:pPr>${sectPrXml}</w:pPr></w:p>`);
    }
    if (isLast && sectPrXml) parts.push(sectPrXml);
  }
  return parts.join("");
}

function parseDocPartBody(body: Element, ctx: DocxReadContext): DocPartSectionOptions[] {
  const bodyChildren: Element[] = [];
  const boundaries: { index: number; sectPr: Element }[] = [];
  for (const child of body.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === "w:sectPr") {
      boundaries.push({ index: bodyChildren.length, sectPr: child });
      continue;
    }
    bodyChildren.push(child);
    if (child.name === "w:p") {
      const pPr = findChild(child, "w:pPr");
      const sectPr = findChild(pPr, "w:sectPr");
      if (sectPr) boundaries.push({ index: bodyChildren.length, sectPr });
    }
  }

  if (boundaries.length === 0) {
    return [{ children: bodyChildren.map((child) => parseSectionChild(child, ctx)) }];
  }

  const sections: DocPartSectionOptions[] = [];
  let start = 0;
  for (const boundary of boundaries) {
    sections.push({
      children: bodyChildren
        .slice(start, boundary.index)
        .map((child) => parseSectionChild(child, ctx)),
      properties: parseSectionPropertiesEl(boundary.sectPr),
    });
    start = boundary.index;
  }
  if (start < bodyChildren.length) {
    sections.push({
      children: bodyChildren.slice(start).map((child) => parseSectionChild(child, ctx)),
    });
  }
  return sections;
}

function docPartPrXml(part: GlossaryDocumentOptions["parts"][number]): string {
  const prParts: string[] = [];
  prParts.push(
    `<w:name w:val="${escapeXml(part.name)}"${part.decorated ? ' w:decorated="1"' : ""}/>`,
  );
  if (part.category || part.gallery) {
    const catParts: string[] = [];
    if (part.category) {
      catParts.push(`<w:name w:val="${escapeXml(part.category)}"/>`);
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
    prParts.push(`<w:description w:val="${escapeXml(part.description)}"/>`);
  }
  if (part.guid) {
    prParts.push(`<w:guid w:val="${escapeXml(part.guid)}"/>`);
  }
  if (part.style) {
    prParts.push(`<w:style w:val="${escapeXml(part.style)}"/>`);
  }
  return `<w:docPartPr>${prParts.join("")}</w:docPartPr>`;
}

export const glossaryDesc: CustomDescriptor<GlossaryDocumentOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const partsXml = opts.parts
      .map(
        (part) =>
          `<w:docPart>${docPartPrXml(part)}<w:docPartBody>${stringifyDocPartBody(part, ctx)}</w:docPartBody></w:docPart>`,
      )
      .join("");

    return `<w:glossaryDocument ${GLOSSARY_NS}><w:docParts>${partsXml}</w:docParts></w:glossaryDocument>`;
  },

  parse(el, ctx) {
    const dctx = ctx as DocxReadContext;
    const parts: DocPartOptions[] = [];

    const docPartsEl = findChild(el, "w:docParts");
    if (!docPartsEl) return { parts };

    for (const docPart of docPartsEl.elements ?? []) {
      if (docPart.name !== "w:docPart") continue;
      const part: Partial<DocPartOptions> = {};

      // Parse w:docPartPr
      const pr = findChild(docPart, "w:docPartPr");
      if (pr) {
        // name
        const name = findChild(pr, "w:name");
        if (name) {
          part.name = attr(name, "w:val") ?? "";
          const decorated = attr(name, "w:decorated");
          if (parseOnOff(decorated)) part.decorated = true;
        }

        // category
        const category = findChild(pr, "w:category");
        if (category) {
          const catName = findChild(category, "w:name");
          if (catName) part.category = attr(catName, "w:val");
          const gallery = findChild(category, "w:gallery");
          if (gallery) part.gallery = attr(gallery, "w:val") as DocPartOptions["gallery"];
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
          if (typeList.length > 0) part.types = typeList as DocPartOptions["types"];
          const allAttr = attr(types, "w:all");
          if (parseOnOff(allAttr)) part.allTypes = true;
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
          if (behaviorList.length > 0) part.behaviors = behaviorList as DocPartOptions["behaviors"];
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

        const style = findChild(pr, "w:style");
        if (style) {
          const val = attr(style, "w:val");
          if (val) part.style = val;
        }
      }

      // Parse the CT_Body section boundaries. Paragraph-hosted w:sectPr ends
      // a non-final section; a direct terminal w:sectPr describes the last one.
      const body = findChild(docPart, "w:docPartBody");
      part.sections = body ? parseDocPartBody(body, dctx) : [];

      parts.push(part as DocPartOptions);
    }

    return { parts };
  },
};
