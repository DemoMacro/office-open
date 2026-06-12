/**
 * Body-level child descriptors for DOCX.
 *
 * Provides descriptor-based stringification for section children.
 * All types use pure string builders — zero class instantiation, zero toXml().
 *
 * @module
 */

import { uniqueId } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import type { AltChunkOptions } from "@parts/alt-chunk/alt-chunk";
import type { CustomXmlPrOptions } from "@parts/custom-xml/custom-xml";
import type { SubDocOptions } from "@parts/sub-doc/sub-doc";
import type { SdtPropertiesOptions } from "@parts/table-of-contents";
import type { SectionChild } from "@shared/section";

import type { BodyContext } from "../context";

// ── XML helpers ──

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ── AltChunk (pure string — registers relationships + altChunks) ──

const ALTCHUNK_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";

function wrapHtmlDocument(fragment: string): string {
  if (/<(!DOCTYPE|html|HTML)/i.test(fragment)) {
    return fragment;
  }
  return `<!DOCTYPE html>\n<html><head><meta charset="utf-8"></head>\n<body>${fragment}</body></html>`;
}

export const altChunkDesc: CustomDescriptor<AltChunkOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const relId = uniqueId();
    const extension = opts.extension;
    const partPath = `afchunks/afchunk${relId}.${extension}`;
    const rawData = typeof opts.data === "string" ? new TextEncoder().encode(opts.data) : opts.data;
    const data =
      opts.contentType === "text/html" && typeof opts.data === "string"
        ? new TextEncoder().encode(wrapHtmlDocument(opts.data))
        : rawData;

    ctx.fileData.document.relationships.addRelationship(relId, ALTCHUNK_REL_TYPE, partPath);
    ctx.fileData.altChunks.addAltChunk(relId, {
      key: relId,
      data,
      path: partPath,
      extension,
      contentType: opts.contentType,
    });

    const rId = `rId${relId}`;
    if (opts.matchSrc) {
      return `<w:altChunk r:id="${rId}"><w:altChunkPr><w:matchSrc/></w:altChunkPr></w:altChunk>`;
    }
    return `<w:altChunk r:id="${rId}"/>`;
  },

  parse() {
    return {};
  },
};

// ── SubDoc (pure string — registers relationships + subDocs) ──

const SUBDOC_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument";

export const subDocDesc: CustomDescriptor<SubDocOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const relId = uniqueId();
    const partPath = `subdocs/subdoc${relId}.docx`;
    const data = typeof opts.data === "string" ? new TextEncoder().encode(opts.data) : opts.data;

    ctx.fileData.document.relationships.addRelationship(relId, SUBDOC_REL_TYPE, partPath);
    ctx.fileData.subDocs.addSubDoc(relId, {
      data,
      path: partPath,
    });

    return `<w:subDoc r:id="rId${relId}"/>`;
  },

  parse() {
    return {};
  },
};

// ── SDT (pure string — inline sdtPr + stringify children) ──

export interface SdtChildOptions {
  properties: SdtPropertiesOptions;
  children?: SectionChild[];
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function sdtListItemXml(
  item: { displayText?: string; value?: string },
  forceValue?: boolean,
): string {
  const attrs: string[] = [];
  if (item.displayText !== undefined) attrs.push(`w:displayText="${escapeXml(item.displayText)}"`);
  const value = item.value ?? (forceValue ? item.displayText : undefined);
  if (value !== undefined) attrs.push(`w:value="${escapeXml(value)}"`);
  return `<w:listItem ${attrs.join(" ")}/>`;
}

function sdtListTypeXml(
  name: string,
  options: { items?: { displayText?: string; value?: string }[]; lastValue?: string },
): string {
  const parts: string[] = [];
  if (options.items) {
    for (const item of options.items) {
      parts.push(sdtListItemXml(item, name === "w:dropDownList"));
    }
  }
  const attrs: string[] = [];
  if (options.lastValue !== undefined) attrs.push(`w:lastValue="${escapeXml(options.lastValue)}"`);
  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  return parts.length ? `<${name}${attrStr}>${parts.join("")}</${name}>` : `<${name}${attrStr}/>`;
}

function sdtDateXml(options: {
  dateFormat?: string;
  languageId?: string;
  storeMappedDataAs?: string;
  calendar?: string;
  fullDate?: string;
}): string {
  const parts: string[] = [];
  if (options.dateFormat !== undefined)
    parts.push(`<w:dateFormat w:val="${escapeXml(options.dateFormat)}"/>`);
  if (options.languageId !== undefined)
    parts.push(`<w:lid w:val="${escapeXml(options.languageId)}"/>`);
  if (options.storeMappedDataAs !== undefined)
    parts.push(`<w:storeMappedDataAs w:val="${options.storeMappedDataAs}"/>`);
  if (options.calendar !== undefined) parts.push(`<w:calendar w:val="${options.calendar}"/>`);
  const attrs: string[] = [];
  if (options.fullDate !== undefined) attrs.push(`w:fullDate="${options.fullDate}"`);
  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  return parts.length ? `<w:date${attrStr}>${parts.join("")}</w:date>` : `<w:date${attrStr}/>`;
}

function sdtDataBindingXml(options: {
  prefixMappings?: string;
  xpath: string;
  storeItemID: string;
}): string {
  const attrs: string[] = [
    `w:xpath="${escapeXml(options.xpath)}"`,
    `w:storeItemID="${escapeXml(options.storeItemID)}"`,
  ];
  if (options.prefixMappings !== undefined)
    attrs.push(`w:prefixMappings="${escapeXml(options.prefixMappings)}"`);
  return `<w:dataBinding ${attrs.join(" ")}/>`;
}

function sdtDocPartXml(
  name: string,
  options: { gallery?: string; category?: string; unique?: boolean },
): string {
  const parts: string[] = [];
  if (options.gallery !== undefined)
    parts.push(`<w:docPartGallery w:val="${escapeXml(options.gallery)}"/>`);
  if (options.category !== undefined)
    parts.push(`<w:docPartCategory w:val="${escapeXml(options.category)}"/>`);
  if (options.unique !== undefined)
    parts.push(options.unique ? "<w:docPartUnique/>" : '<w:docPartUnique w:val="0"/>');
  return parts.length ? `<${name}>${parts.join("")}</${name}>` : `<${name}/>`;
}

function onOffAttr(name: string, val: boolean): string {
  return val ? `<${name}/>` : `<${name} w:val="0"/>`;
}

function stringifySdtPr(opts: SdtPropertiesOptions): string {
  const parts: string[] = [];

  // rPr is not supported in pure JSON path — skip

  if (opts.alias !== undefined) parts.push(`<w:alias w:val="${escapeXml(opts.alias)}"/>`);
  if (opts.tag !== undefined) parts.push(`<w:tag w:val="${escapeXml(opts.tag)}"/>`);
  if (opts.id !== undefined) parts.push(`<w:id w:val="${opts.id}"/>`);
  if (opts.lock !== undefined) parts.push(`<w:lock w:val="${opts.lock}"/>`);

  // placeholder not supported in pure JSON path — skip

  if (opts.temporary !== undefined) parts.push(onOffAttr("w:temporary", opts.temporary));
  const effectiveShowingPlcHdr = opts.showingPlaceholder ?? false;
  if (opts.showingPlaceholder !== undefined || effectiveShowingPlcHdr) {
    parts.push(onOffAttr("w:showingPlcHdr", effectiveShowingPlcHdr));
  }
  if (opts.dataBinding) parts.push(sdtDataBindingXml(opts.dataBinding));
  if (opts.label !== undefined) parts.push(`<w:label w:val="${opts.label}"/>`);
  if (opts.tabIndex !== undefined) parts.push(`<w:tabIndex w:val="${opts.tabIndex}"/>`);

  // Type discriminator (xsd:choice)
  if (opts.equation) {
    parts.push("<w:equation/>");
  } else if (opts.comboBox) {
    parts.push(sdtListTypeXml("w:comboBox", opts.comboBox));
  } else if (opts.date) {
    parts.push(sdtDateXml(opts.date));
  } else if (opts.docPartObj) {
    parts.push(sdtDocPartXml("w:docPartObj", opts.docPartObj));
  } else if (opts.docPartList) {
    parts.push(sdtDocPartXml("w:docPartList", opts.docPartList));
  } else if (opts.dropDownList) {
    parts.push(sdtListTypeXml("w:dropDownList", opts.dropDownList));
  } else if (opts.picture) {
    parts.push("<w:picture/>");
  } else if (opts.richText) {
    parts.push("<w:richText/>");
  } else if (opts.text !== undefined) {
    const multiLine = opts.text.multiLine ?? false;
    parts.push(`<w:text w:multiLine="${multiLine}"/>`);
  } else if (opts.citation) {
    parts.push("<w:citation/>");
  } else if (opts.group) {
    parts.push("<w:group/>");
  } else if (opts.bibliography) {
    parts.push("<w:bibliography/>");
  }

  return parts.length ? `<w:sdtPr>${parts.join("")}</w:sdtPr>` : "<w:sdtPr/>";
}

export const sdtBlockDesc: CustomDescriptor<SdtChildOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = ["<w:sdt>"];

    // sdtPr
    parts.push(stringifySdtPr(opts.properties));

    // sdtEndPr — typically empty, included for round-trip fidelity with Word
    parts.push("<w:sdtEndPr/>");

    // sdtContent — serialize children directly (no coerce needed)
    if (opts.children && opts.children.length > 0) {
      const contentParts: string[] = [];
      for (const child of opts.children) {
        contentParts.push(ctx.stringifyChild(child, ctx));
      }
      const contentBody = contentParts.join("");
      parts.push(contentBody ? `<w:sdtContent>${contentBody}</w:sdtContent>` : "<w:sdtContent/>");
    }

    parts.push("</w:sdt>");
    return parts.join("");
  },

  parse() {
    return {};
  },
};

// ── Custom XML (pure string — no context side effects) ──

export interface CustomXmlBlockDescriptorOptions {
  element: string;
  uri?: string;
  customXmlPr?: CustomXmlPrOptions;
  children?: SectionChild[];
}

function buildCustomXmlPrXml(pr: CustomXmlPrOptions): string {
  const parts: string[] = ["<w:customXmlPr>"];
  if (pr.placeholder !== undefined) {
    parts.push(`<w:placeholder w:val="${escapeAttr(pr.placeholder)}"/>`);
  }
  if (pr.attributes) {
    for (const attr of pr.attributes) {
      const attrParts: string[] = [
        `w:name="${escapeAttr(attr.name)}"`,
        `w:val="${escapeAttr(attr.val)}"`,
      ];
      if (attr.uri !== undefined) attrParts.push(`w:uri="${escapeAttr(attr.uri)}"`);
      parts.push(`<w:attr ${attrParts.join(" ")}/>`);
    }
  }
  parts.push("</w:customXmlPr>");
  return parts.join("");
}

export const customXmlBlockDesc: CustomDescriptor<CustomXmlBlockDescriptorOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const attrs: string[] = [`w:element="${escapeAttr(opts.element)}"`];
    if (opts.uri !== undefined) attrs.push(`w:uri="${escapeAttr(opts.uri)}"`);
    const parts: string[] = [`<w:customXml ${attrs.join(" ")}>`];

    if (opts.customXmlPr) {
      parts.push(buildCustomXmlPrXml(opts.customXmlPr));
    }

    if (opts.children) {
      for (const child of opts.children) {
        parts.push(ctx.stringifyChild(child, ctx));
      }
    }

    parts.push("</w:customXml>");
    return parts.join("");
  },

  parse() {
    return {};
  },
};
