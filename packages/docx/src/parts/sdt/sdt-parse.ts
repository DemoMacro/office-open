/**
 * Structured Document Tag parser for DOCX documents.
 *
 * Parses w:sdt elements into SdtPropertiesOptions + children.
 *
 * @module
 */
import { attr, attrBool, attrNum, children, findChild, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type {
  SdtPropertiesOptions,
  SdtListItem,
  SdtDateOptions,
  SdtTextOptions,
  SdtComboBoxOptions,
  SdtDropDownListOptions,
} from "@parts/table-of-contents";

import type { DocxReadContext } from "../../context";

/**
 * Parse w:sdtPr element into SdtPropertiesOptions.
 */
function parseSdtProperties(el: Element): SdtPropertiesOptions {
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
    const items: SdtListItem[] = [];
    for (const li of children(comboBox, "w:listItem")) {
      items.push({
        displayText: attr(li, "w:displayText"),
        value: attr(li, "w:value"),
      });
    }
    opts.comboBox = {
      items: items.length > 0 ? items : undefined,
      lastValue: attr(comboBox, "w:lastValue"),
    } as SdtComboBoxOptions;
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
    opts.date = dateOpts as SdtDateOptions;
  } else if (findChild(el, "w:dropDownList")) {
    const ddl = findChild(el, "w:dropDownList")!;
    const items: SdtListItem[] = [];
    for (const li of children(ddl, "w:listItem")) {
      items.push({
        displayText: attr(li, "w:displayText"),
        value: attr(li, "w:value"),
      });
    }
    opts.dropDownList = {
      items: items.length > 0 ? items : undefined,
      lastValue: attr(ddl, "w:lastValue"),
    } as SdtDropDownListOptions;
  } else if (findChild(el, "w:picture")) {
    opts.picture = true;
  } else if (findChild(el, "w:richText")) {
    opts.richText = true;
  } else if (findChild(el, "w:text")) {
    const text = findChild(el, "w:text")!;
    opts.text = {
      multiLine: attrBool(text, "w:multiLine"),
    } as SdtTextOptions;
  } else if (findChild(el, "w:citation")) {
    opts.citation = true;
  } else if (findChild(el, "w:group")) {
    opts.group = true;
  } else if (findChild(el, "w:bibliography")) {
    opts.bibliography = true;
  } else if (findChild(el, "w:docPartObj")) {
    const dp = findChild(el, "w:docPartObj")!;
    opts.docPartObj = {};
    const gallery = findChild(dp, "w:docPartGallery");
    if (gallery) (opts.docPartObj as Record<string, unknown>).gallery = attr(gallery, "w:val");
    const category = findChild(dp, "w:docPartCategory");
    if (category) (opts.docPartObj as Record<string, unknown>).category = attr(category, "w:val");
  } else if (findChild(el, "w:docPartList")) {
    const dp = findChild(el, "w:docPartList")!;
    opts.docPartList = {};
    const gallery = findChild(dp, "w:docPartGallery");
    if (gallery) (opts.docPartList as Record<string, unknown>).gallery = attr(gallery, "w:val");
    const category = findChild(dp, "w:docPartCategory");
    if (category) (opts.docPartList as Record<string, unknown>).category = attr(category, "w:val");
  }

  return opts as SdtPropertiesOptions;
}

/**
 * Parse a block-level w:sdt element.
 * Returns an object suitable for the { sdt: ... } SectionChild variant.
 */
export function parseSdtBlock(
  el: Element,
  ctx: DocxReadContext,
  parseChildren: (elements: Element[], ctx: DocxReadContext) => unknown[],
): {
  properties: SdtPropertiesOptions;
  children?: unknown[];
} {
  const sdtPr = findChild(el, "w:sdtPr");
  const properties = sdtPr ? parseSdtProperties(sdtPr) : {};

  const sdtContent = findChild(el, "w:sdtContent");
  let childList: unknown[] | undefined;
  if (sdtContent) {
    childList = parseChildren(sdtContent.elements ?? [], ctx);
    if (childList.length === 0) childList = undefined;
  }

  return { properties, children: childList };
}
