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
import {
  attr,
  attrBool,
  attrNum,
  children as xmlChildren,
  findChild,
  textOf,
} from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { AltChunkOptions } from "@parts/alt-chunk/alt-chunk";
import type { CustomXmlPrOptions } from "@parts/custom-xml/custom-xml";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import { parseRunProperties } from "@parts/paragraph/run/run-parse";
import { stringifyRunPropertiesInner } from "@parts/paragraph/stringify";
import type { SubDocOptions } from "@parts/sub-doc/sub-doc";
import type { SdtCheckboxOptions, SdtPropertiesOptions } from "@parts/table-of-contents";
import type { SectionChild } from "@shared/section";

import type { BodyContext, DocxReadContext } from "../context";

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

  parse(el, ctx) {
    const rId = attr(el, "r:id");
    const opts: Record<string, unknown> = {};

    // Check for matchSrc
    const altChunkPr = findChild(el, "w:altChunkPr");
    if (altChunkPr && findChild(altChunkPr, "w:matchSrc")) {
      opts.matchSrc = true;
    }

    // Resolve the altChunk data from relationships
    const dctx = ctx as DocxReadContext;
    if (rId) {
      const path = dctx.resolveRelationship(rId);
      if (path) {
        const data = dctx.getRaw(path);
        if (data) {
          opts.data = data;
          const ext = path.split(".").pop() ?? "txt";
          switch (ext) {
            case "html":
              opts.contentType = "text/html";
              opts.extension = "html";
              break;
            case "rtf":
              opts.contentType = "application/rtf";
              opts.extension = "rtf";
              break;
            default:
              opts.contentType = "text/plain";
              opts.extension = "txt";
              break;
          }
        }
      }
    }

    return opts as unknown as AltChunkOptions;
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

  parse(el, ctx) {
    const rId = attr(el, "r:id");
    const dctx = ctx as DocxReadContext;
    if (rId) {
      const path = dctx.resolveRelationship(rId);
      if (path) {
        const data = dctx.getRaw(path);
        if (data) {
          return { data } as SubDocOptions;
        }
      }
    }
    return { data: new Uint8Array(0) } as SubDocOptions;
  },
};

// ── SDT (pure string — inline sdtPr + stringify children) ──

export interface SdtChildOptions {
  properties: SdtPropertiesOptions;
  children?: SectionChild[];
  /** Run properties for the SDT end mark (w:sdtEndPr). */
  endProperties?: RunPropertiesOptions;
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

const W14_NS = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

/** Default symbol font for checkbox content controls (CT_SdtCheckboxSymbol). */
const CHECKBOX_FONT = "MS Gothic";

const DEFAULT_CHECKED: { val: string; font: string } = { val: "2612", font: CHECKBOX_FONT };
const DEFAULT_UNCHECKED: { val: string; font: string } = { val: "2610", font: CHECKBOX_FONT };

/**
 * Build a w14:checkbox element (Word 2010+ content control checkbox).
 *
 * Lives in the w14 extension namespace; emitted with an inline xmlns:w14 so it
 * is valid wherever w:sdtPr appears. validate.ts tolerates the w14 namespace.
 */
function sdtCheckboxXml(opts: SdtCheckboxOptions): string {
  const checked = opts.checkedState ?? DEFAULT_CHECKED;
  const unchecked = opts.uncheckedState ?? DEFAULT_UNCHECKED;
  const inner =
    (opts.checked ? "<w14:checked/>" : '<w14:checked w14:val="0"/>') +
    `<w14:checkedState w14:val="${escapeXml(checked.val)}" w14:font="${escapeXml(checked.font ?? CHECKBOX_FONT)}"/>` +
    `<w14:uncheckedState w14:val="${escapeXml(unchecked.val)}" w14:font="${escapeXml(unchecked.font ?? CHECKBOX_FONT)}"/>`;
  return `<w14:checkbox ${W14_NS}>${inner}</w14:checkbox>`;
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
  } else if (opts.checkbox) {
    parts.push(sdtCheckboxXml(opts.checkbox));
  }

  return parts.length ? `<w:sdtPr>${parts.join("")}</w:sdtPr>` : "<w:sdtPr/>";
}

// ── SDT parse helpers ──

/** Parse w:sdtPr element into SdtPropertiesOptions. */
function parseSdtPr(el: Element): SdtPropertiesOptions {
  const opts: Record<string, unknown> = {};

  const alias = findChild(el, "w:alias");
  if (alias) opts.alias = attr(alias, "w:val");

  const tag = findChild(el, "w:tag");
  if (tag) {
    const val = attr(tag, "w:val");
    if (val) opts.tag = val;
  }

  const id = findChild(el, "w:id");
  if (id) {
    const val = attrNum(id, "w:val");
    if (val !== undefined) opts.id = val;
  }

  const lock = findChild(el, "w:lock");
  if (lock) {
    const val = attr(lock, "w:val");
    if (val) opts.lock = val as SdtPropertiesOptions["lock"];
  }

  const temporary = findChild(el, "w:temporary");
  if (temporary) opts.temporary = attrBool(temporary, "w:val") ?? true;

  const showingPlcHdr = findChild(el, "w:showingPlcHdr");
  if (showingPlcHdr) opts.showingPlaceholder = attrBool(showingPlcHdr, "w:val") ?? true;

  const label = findChild(el, "w:label");
  if (label) {
    const val = attrNum(label, "w:val");
    if (val !== undefined) opts.label = val;
  }

  const tabIndex = findChild(el, "w:tabIndex");
  if (tabIndex) {
    const val = attrNum(tabIndex, "w:val");
    if (val !== undefined) opts.tabIndex = val;
  }

  // Data binding
  const dataBinding = findChild(el, "w:dataBinding");
  if (dataBinding) {
    opts.dataBinding = {
      xpath: attr(dataBinding, "w:xpath") ?? "",
      storeItemID: attr(dataBinding, "w:storeItemID") ?? "",
      prefixMappings: attr(dataBinding, "w:prefixMappings"),
    };
  }

  // Type discriminators (xsd:choice)
  if (findChild(el, "w:equation")) {
    opts.equation = true;
  } else if (findChild(el, "w:comboBox")) {
    const comboBox = findChild(el, "w:comboBox")!;
    const items: { displayText?: string; value?: string }[] = [];
    for (const li of xmlChildren(comboBox, "w:listItem")) {
      items.push({ displayText: attr(li, "w:displayText"), value: attr(li, "w:value") });
    }
    opts.comboBox = {
      items: items.length > 0 ? items : undefined,
      lastValue: attr(comboBox, "w:lastValue"),
    };
  } else if (findChild(el, "w:date")) {
    const date = findChild(el, "w:date")!;
    const dateOpts: Record<string, unknown> = {};
    const dateFormat = findChild(date, "w:dateFormat");
    if (dateFormat) dateOpts.dateFormat = textOf(dateFormat);
    const lid = findChild(date, "w:lid");
    if (lid) dateOpts.languageId = textOf(lid);
    const storeMapped = findChild(date, "w:storeMappedDataAs");
    if (storeMapped) dateOpts.storeMappedDataAs = attr(storeMapped, "w:val");
    const calendar = findChild(date, "w:calendar");
    if (calendar) dateOpts.calendar = attr(calendar, "w:val");
    const fullDate = attr(date, "w:fullDate");
    if (fullDate) dateOpts.fullDate = fullDate;
    opts.date = dateOpts;
  } else if (findChild(el, "w:docPartObj")) {
    const dp = findChild(el, "w:docPartObj")!;
    const dpObj: Record<string, unknown> = {};
    const gallery = findChild(dp, "w:docPartGallery");
    if (gallery) dpObj.gallery = attr(gallery, "w:val");
    const category = findChild(dp, "w:docPartCategory");
    if (category) dpObj.category = attr(category, "w:val");
    if (findChild(dp, "w:docPartUnique")) dpObj.unique = true;
    opts.docPartObj = dpObj;
  } else if (findChild(el, "w:docPartList")) {
    const dp = findChild(el, "w:docPartList")!;
    const dpObj: Record<string, unknown> = {};
    const gallery = findChild(dp, "w:docPartGallery");
    if (gallery) dpObj.gallery = attr(gallery, "w:val");
    const category = findChild(dp, "w:docPartCategory");
    if (category) dpObj.category = attr(category, "w:val");
    if (findChild(dp, "w:docPartUnique")) dpObj.unique = true;
    opts.docPartList = dpObj;
  } else if (findChild(el, "w:dropDownList")) {
    const ddl = findChild(el, "w:dropDownList")!;
    const items: { displayText?: string; value?: string }[] = [];
    for (const li of xmlChildren(ddl, "w:listItem")) {
      items.push({ displayText: attr(li, "w:displayText"), value: attr(li, "w:value") });
    }
    opts.dropDownList = {
      items: items.length > 0 ? items : undefined,
      lastValue: attr(ddl, "w:lastValue"),
    };
  } else if (findChild(el, "w:picture")) {
    opts.picture = true;
  } else if (findChild(el, "w:richText")) {
    opts.richText = true;
  } else if (findChild(el, "w:text")) {
    const text = findChild(el, "w:text")!;
    opts.text = { multiLine: attrBool(text, "w:multiLine") };
  } else if (findChild(el, "w:citation")) {
    opts.citation = true;
  } else if (findChild(el, "w:group")) {
    opts.group = true;
  } else if (findChild(el, "w:bibliography")) {
    opts.bibliography = true;
  } else if (findChild(el, "w14:checkbox")) {
    const cb = findChild(el, "w14:checkbox")!;
    const cbObj: Record<string, unknown> = {};
    const checked = findChild(cb, "w14:checked");
    if (checked) cbObj.checked = attrBool(checked, "w14:val") ?? true;
    const checkedState = findChild(cb, "w14:checkedState");
    if (checkedState)
      cbObj.checkedState = {
        val: attr(checkedState, "w14:val") ?? "",
        font: attr(checkedState, "w14:font"),
      };
    const uncheckedState = findChild(cb, "w14:uncheckedState");
    if (uncheckedState)
      cbObj.uncheckedState = {
        val: attr(uncheckedState, "w14:val") ?? "",
        font: attr(uncheckedState, "w14:font"),
      };
    opts.checkbox = cbObj;
  }

  return opts as SdtPropertiesOptions;
}

/** Parse w:customXmlPr element into CustomXmlPrOptions. */
function parseCustomXmlPr(el: Element): CustomXmlPrOptions {
  const opts: Record<string, unknown> = {};
  const placeholder = findChild(el, "w:placeholder");
  if (placeholder) {
    const val = attr(placeholder, "w:val");
    if (val) opts.placeholder = val;
  }
  const attributes: { name: string; val: string; uri?: string }[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "w:attr") continue;
    const name = attr(child, "w:name");
    const val = attr(child, "w:val");
    if (name && val) {
      const attrOpts: { name: string; val: string; uri?: string } = { name, val };
      const uriVal = attr(child, "w:uri");
      if (uriVal) attrOpts.uri = uriVal;
      attributes.push(attrOpts);
    }
  }
  if (attributes.length > 0) opts.attributes = attributes;
  return opts as unknown as CustomXmlPrOptions;
}

/** Body child element parsing callback for SDT/customXml content. */
let _parseBodyChild: ((el: Element, ctx: DocxReadContext) => SectionChild) | undefined;

/** Register the body child parser (called from parse/body.ts to break circular dependency). */
export function setBodyParseChild(
  parser: (el: Element, ctx: DocxReadContext) => SectionChild,
): void {
  _parseBodyChild = parser;
}

function parseBodyChildren(elements: Element[], ctx: DocxReadContext): SectionChild[] {
  if (!_parseBodyChild) return [];
  const result: SectionChild[] = [];
  for (const el of elements) {
    result.push(_parseBodyChild(el, ctx));
  }
  return result;
}

export const sdtBlockDesc: CustomDescriptor<SdtChildOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = ["<w:sdt>"];

    // sdtPr
    parts.push(stringifySdtPr(opts.properties));

    // sdtEndPr — typically empty, included for round-trip fidelity with Word
    const endPrInner = opts.endProperties
      ? stringifyRunPropertiesInner(opts.endProperties)
      : undefined;
    parts.push(endPrInner ? `<w:sdtEndPr>${endPrInner}</w:sdtEndPr>` : "<w:sdtEndPr/>");

    // sdtContent — checkbox renders its current state symbol; otherwise serialize children
    if (opts.properties.checkbox) {
      const cb = opts.properties.checkbox;
      const symbol =
        (cb.checked ?? false)
          ? (cb.checkedState ?? DEFAULT_CHECKED)
          : (cb.uncheckedState ?? DEFAULT_UNCHECKED);
      const font = escapeXml(symbol.font ?? CHECKBOX_FONT);
      const char = escapeXml(String.fromCodePoint(parseInt(symbol.val, 16)));
      parts.push(
        `<w:sdtContent><w:p><w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/></w:rPr><w:t>${char}</w:t></w:r></w:p></w:sdtContent>`,
      );
    } else if (opts.children && opts.children.length > 0) {
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

  parse(el, ctx) {
    const dctx = ctx as DocxReadContext;

    // Parse sdtPr
    const sdtPr = findChild(el, "w:sdtPr");
    const properties = sdtPr ? parseSdtPr(sdtPr) : {};

    // Parse sdtEndPr (CT_RPr content at the end mark — run properties, no w:rPr wrapper)
    let endProperties: RunPropertiesOptions | undefined;
    const sdtEndPr = findChild(el, "w:sdtEndPr");
    if (sdtEndPr) {
      endProperties = parseRunProperties(sdtEndPr);
    }

    // Parse sdtContent children
    const sdtContent = findChild(el, "w:sdtContent");
    let childList: SectionChild[] | undefined;
    if (sdtContent && sdtContent.elements?.length) {
      childList = parseBodyChildren(sdtContent.elements, dctx);
      if (childList.length === 0) childList = undefined;
    }

    return { properties, children: childList, endProperties } as SdtChildOptions;
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

  parse(el, ctx) {
    const dctx = ctx as DocxReadContext;
    const opts: Record<string, unknown> = {};

    const element = attr(el, "w:element");
    if (element) opts.element = element;

    const uri = attr(el, "w:uri");
    if (uri) opts.uri = uri;

    // Parse w:customXmlPr
    const xmlPr = findChild(el, "w:customXmlPr");
    if (xmlPr) {
      opts.customXmlPr = parseCustomXmlPr(xmlPr);
    }

    // Parse block-level children
    const childList: SectionChild[] = [];
    for (const child of el.elements ?? []) {
      if (child.name === "w:customXmlPr") continue;
      if (_parseBodyChild) {
        childList.push(_parseBodyChild(child, dctx));
      }
    }
    if (childList.length > 0) opts.children = childList;

    return opts as unknown as CustomXmlBlockDescriptorOptions;
  },
};
