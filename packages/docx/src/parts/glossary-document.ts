/**
 * Glossary document component — stores building block definitions.
 *
 * Generates word/glossary/document.xml containing Quick Parts entries
 * that appear in Word's Insert > Quick Parts gallery.
 *
 * @module
 */

import {
  parseOnOff,
  partPathToRelsPath,
  RELATIONSHIP_TYPES,
  relativePartTarget,
  Relationships,
  resolveRelationshipTarget,
} from "@office-open/core";
import {
  parseSectionPropertiesEl,
  sectionPropertiesDesc,
  type SectionPropertiesDescriptorOptions,
} from "@parts/document/body/section-properties/descriptor";
import type { HeaderFooterGroup } from "@parts/document/body/section-properties/section-properties";
import type { SectionPropertiesOptions } from "@parts/document/body/section-properties/section-properties";
import { documentNamespaceAttributes } from "@parts/document/document-attributes";
import type { EndnoteOptions, EndnoteSeparator } from "@parts/endnotes/descriptor";
import type { EmbeddedFontOptions } from "@parts/fonts/font-table";
import type { FootnoteOptions, FootnoteSeparator } from "@parts/footnotes/descriptor";
import type { HeaderFooterReference } from "@parts/header-footer";
import type { SettingsOptions } from "@parts/settings/settings";
import type { StylesOptions } from "@parts/styles/styles";
import type { WebSettingsOptions } from "@parts/web-settings";
import type { SectionChild } from "@shared/section";

/** Gallery type for building blocks (ST_DocPartGallery) */
export const DocPartGallery = {
  PLACEHOLDER: "placeholder",
  /** Matches any gallery — used by docPartObj references without a specific gallery. */
  ANY: "any",
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

/** Header/footer content slots referenced by one docPart section. Round-trip
 *  pins the source part names; fresh authoring auto-numbers. */
export interface DocPartHeaderFooterOptions {
  default?: SectionChild[];
  first?: SectionChild[];
  even?: SectionChild[];
  /** Source part name per slot relative to word/ (round-trip only). */
  partNames?: { default?: string; first?: string; even?: string };
}

/** A section within a building block body (CT_Body section boundary). */
export interface DocPartSectionOptions {
  /** Block-level content in this section. */
  children: SectionChild[];
  /** Section properties carried by w:pPr/w:sectPr or terminal w:sectPr. */
  properties?: SectionPropertiesOptions;
  /** Headers referenced by this section's w:headerReference elements. */
  headers?: DocPartHeaderFooterOptions;
  /** Footers referenced by this section's w:footerReference elements. */
  footers?: DocPartHeaderFooterOptions;
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
  /**
   * Glossary document path relative to word/ (round-trip, e.g.
   * "glossary/document.xml"). Fresh generation uses glossary/document.xml.
   */
  partName?: string;
  /**
   * Whether source carried w:docParts (round-trip only — do not hand-author).
   * Distinguishes an empty container from the XSD-optional container's absence.
   */
  hasDocParts?: boolean;
  /**
   * Independent numbering part owned by the glossary. Glossary bodies resolve
   * w:numId against this part, not the main document numbering part.
   */
  numbering?: NumberingOptions;
  /** Glossary numbering path relative to word/ (round-trip only). */
  numberingPartName?: string;
  /** Independent settings part owned by the glossary. */
  settings?: SettingsOptions;
  /** Glossary settings path relative to word/ (round-trip only). */
  settingsPartName?: string;
  /** Independent styles part owned by the glossary. */
  styles?: StylesOptions;
  /** Glossary styles path relative to word/ (round-trip only). */
  stylesPartName?: string;
  /** Independent web-settings part owned by the glossary. */
  webSettings?: WebSettingsOptions;
  /** Glossary web-settings path relative to word/ (round-trip only). */
  webSettingsPartName?: string;
  /** Independent font table owned by the glossary. */
  fonts?: EmbeddedFontOptions[];
  /** Glossary font-table path relative to word/ (round-trip only). */
  fontTablePartName?: string;
  /** Independent footnotes part owned by the glossary. */
  footnotes?: FootnoteOptions[];
  /** Glossary footnotes path relative to word/ (round-trip only). */
  footnotesPartName?: string;
  /** Glossary footnote separators — round-tripped verbatim from the source. */
  footnoteSeparators?: {
    separator?: FootnoteSeparator | null;
    continuationSeparator?: FootnoteSeparator | null;
    continuationNotice?: FootnoteSeparator;
  };
  /** Independent endnotes part owned by the glossary. */
  endnotes?: EndnoteOptions[];
  /** Glossary endnotes path relative to word/ (round-trip only). */
  endnotesPartName?: string;
  /** Glossary endnote separators — round-tripped verbatim from the source. */
  endnoteSeparators?: {
    separator?: EndnoteSeparator | null;
    continuationSeparator?: EndnoteSeparator | null;
    continuationNotice?: EndnoteSeparator;
  };
  /** Building blocks */
  parts: DocPartOptions[];
}

// ── Descriptor ──

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { stringifyBodyChild } from "../body";
import type { BodyContext, DocxReadContext } from "../context";
import { parseHeaderFooterPartChildren, parseSectionChild } from "../parse/body";
import type { NumberingOptions } from "../parts/numbering/numbering";

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

/** Claim a glossary header/footer relationship id from the source package
 *  rels, falling back to the next fresh id. */
function claimDocPartReferenceId(
  ctx: BodyContext,
  glossaryPartPath: string,
  type: string,
  partName: string,
  rels: Relationships,
): number {
  const source = (ctx.fileData._options.passthroughRelationships ?? []).find(
    (rel) =>
      rel.source === glossaryPartPath &&
      rel.relationshipType === type &&
      (rel.target === partName || rel.target.endsWith(`/${partName}`)),
  );
  const claimed = source ? /^rId(\d+)$/.exec(source.rId) : undefined;
  if (claimed && !rels.hasId(claimed[0]!)) return Number(claimed[1]!);
  return rels.nextRelationshipId;
}

/** Register one section's header/footer parts on the glossary rels and
 *  collect the entries for the compile phase to serialize. Two slots may
 *  share one part — its relationship and entry are registered once. */
function buildDocPartHeaderFooterRefs(
  section: DocPartSectionOptions,
  kind: "header" | "footer",
  glossaryPartPath: string,
  ctx: BodyContext,
): HeaderFooterGroup<HeaderFooterReference> | undefined {
  const slots = kind === "header" ? section.headers : section.footers;
  if (!slots) return undefined;
  const rels = ctx.viewWrapper.relationships;
  const type = kind === "header" ? RELATIONSHIP_TYPES.header : RELATIONSHIP_TYPES.footer;
  const entries =
    kind === "header" ? ctx.fileData.glossaryHeaderParts : ctx.fileData.glossaryFooterParts;
  const refs: HeaderFooterGroup<HeaderFooterReference> = {};
  for (const slot of ["default", "first", "even"] as const) {
    const children = slots[slot];
    if (!children) continue;
    let partName = slots.partNames?.[slot];
    if (!partName) {
      let n = entries.length + 1;
      partName = `glossary/${kind}${n}.xml`;
      while (entries.some((entry) => entry.partName === partName)) {
        n++;
        partName = `glossary/${kind}${n}.xml`;
      }
    }
    const target = relativePartTarget(glossaryPartPath, `word/${partName}`);
    let referenceId: number;
    const existing = rels.idOf(type, target);
    const existingMatch = existing ? /^rId(\d+)$/.exec(existing) : undefined;
    if (existingMatch) {
      referenceId = Number(existingMatch[1]!);
    } else {
      referenceId = claimDocPartReferenceId(ctx, glossaryPartPath, type, partName, rels);
      rels.addRelationship(referenceId, type, target);
    }
    if (!entries.some((entry) => entry.partName === partName)) {
      entries.push({ children, relationships: new Relationships(), referenceId, partName });
    }
    refs[slot] = { referenceId };
  }
  return Object.keys(refs).length > 0 ? refs : undefined;
}

/** Build the sectPr XML for one docPart section, registering header/footer
 *  content on the glossary rels so the compiler emits the parts. */
function stringifyDocPartSectPr(
  section: DocPartSectionOptions,
  glossaryPartPath: string,
  ctx: BodyContext,
): string {
  if (!section.properties && !section.headers && !section.footers) return "";
  const descriptorOptions: SectionPropertiesDescriptorOptions = {
    ...section.properties,
    headerReferences: section.headers
      ? buildDocPartHeaderFooterRefs(section, "header", glossaryPartPath, ctx)
      : undefined,
    footerReferences: section.footers
      ? buildDocPartHeaderFooterRefs(section, "footer", glossaryPartPath, ctx)
      : undefined,
  };
  return sectionPropertiesDesc.stringify(descriptorOptions, ctx) ?? "";
}

function stringifyDocPartBody(part: DocPartOptions, ctx: BodyContext): string {
  const glossaryPartPath = ctx.viewWrapper?.partName ?? "word/glossary/document.xml";
  const parts: string[] = [];
  for (let sectionIndex = 0; sectionIndex < part.sections.length; sectionIndex++) {
    const section = part.sections[sectionIndex]!;
    const sectPrXml = stringifyDocPartSectPr(section, glossaryPartPath, ctx);
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
      ...parseDocPartHeaderFooterSlots(boundary.sectPr, ctx),
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

/** Resolve one r:id against the current part's own .rels file. */
function resolveCurrentPartRelationship(ctx: DocxReadContext, rId: string): string | undefined {
  const rels = ctx.docx.doc.get(partPathToRelsPath(ctx.currentPart));
  for (const rel of rels?.elements ?? []) {
    if (rel.name !== "Relationship" || attr(rel, "Id") !== rId) continue;
    const target = attr(rel, "Target");
    return target ? resolveRelationshipTarget(ctx.currentPart, target) : undefined;
  }
  return undefined;
}

/** Header/footer content slots referenced by one docPart w:sectPr. Part
 *  names are package paths relative to word/ so rebuild lands in the source
 *  files (word/glossary/header3.xml). */
function parseDocPartHeaderFooterSlots(
  sectPr: Element,
  ctx: DocxReadContext,
): { headers?: DocPartHeaderFooterOptions; footers?: DocPartHeaderFooterOptions } {
  const parsed = {
    header: {} as Record<string, SectionChild[]>,
    footer: {} as Record<string, SectionChild[]>,
  };
  const partNames = {
    header: {} as Record<string, string>,
    footer: {} as Record<string, string>,
  };
  for (const child of sectPr.elements ?? []) {
    if (child.name !== "w:headerReference" && child.name !== "w:footerReference") continue;
    const rId = attr(child, "r:id");
    const slot = attr(child, "w:type");
    if (!rId || !slot) continue;
    const kind = child.name === "w:headerReference" ? "header" : "footer";
    const path = resolveCurrentPartRelationship(ctx, rId);
    if (!path?.startsWith("word/")) continue;
    const children = parseHeaderFooterPartChildren(path, ctx);
    if (children.length === 0) continue;
    ctx.consumedPartPaths.add(path);
    parsed[kind][slot] = children;
    partNames[kind][slot] = path.slice("word/".length);
  }
  const result: { headers?: DocPartHeaderFooterOptions; footers?: DocPartHeaderFooterOptions } = {};
  if (Object.keys(parsed.header).length > 0) {
    result.headers = { ...parsed.header, partNames: partNames.header };
  }
  if (Object.keys(parsed.footer).length > 0) {
    result.footers = { ...parsed.footer, partNames: partNames.footer };
  }
  return result;
}

function docPartPrXml(part: GlossaryDocumentOptions["parts"][number]): string {
  const prParts: string[] = [];
  prParts.push(
    `<w:name w:val="${escapeXml(part.name)}"${
      part.decorated !== undefined ? ` w:decorated="${part.decorated ? 1 : 0}"` : ""
    }/>`,
  );
  if (part.style) {
    prParts.push(`<w:style w:val="${escapeXml(part.style)}"/>`);
  }
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
    const allAttr = part.allTypes !== undefined ? ` w:all="${part.allTypes ? 1 : 0}"` : "";
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
  return `<w:docPartPr>${prParts.join("")}</w:docPartPr>`;
}

export const glossaryDesc: CustomDescriptor<GlossaryDocumentOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    if (opts.hasDocParts === false) {
      return `<w:glossaryDocument ${GLOSSARY_NS}></w:glossaryDocument>`;
    }
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
    if (!docPartsEl) return { hasDocParts: false, parts };

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
          if (decorated !== undefined) part.decorated = parseOnOff(decorated) ?? false;
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
          if (allAttr !== undefined) part.allTypes = parseOnOff(allAttr) ?? false;
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

    return { hasDocParts: true, parts };
  },
};
