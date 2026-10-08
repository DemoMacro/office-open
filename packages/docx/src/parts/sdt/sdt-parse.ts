/**
 * Structured Document Tag parser for DOCX documents.
 *
 * Parses w:sdt elements into SdtPropertiesOptions + children.
 *
 * @module
 */
import { attr, attrBool, attrNum, children, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import { parseRunProperties } from "@parts/paragraph/run/run-parse";
import type {
  SdtChildOrder,
  SdtPropertiesOptions,
  SdtListItem,
  SdtDateOptions,
  SdtTextOptions,
  SdtComboBoxOptions,
  SdtDropDownListOptions,
} from "@parts/table-of-contents";

const SDT_ELEMENT_ORDER = new Map<string, SdtChildOrder>([
  ["w:rPr", "runProperties"],
  ["w:alias", "alias"],
  ["w:tag", "tag"],
  ["w:id", "id"],
  ["w:lock", "lock"],
  ["w:placeholder", "placeholder"],
  ["w:temporary", "temporary"],
  ["w:showingPlcHdr", "showingPlaceholder"],
  ["w:label", "label"],
  ["w:tabIndex", "tabIndex"],
  ["w:dataBinding", "dataBinding"],
  ["w:equation", "equation"],
  ["w:comboBox", "comboBox"],
  ["w:date", "date"],
  ["w:dropDownList", "dropDownList"],
  ["w:picture", "picture"],
  ["w:richText", "richText"],
  ["w:text", "text"],
  ["w:citation", "citation"],
  ["w:group", "group"],
  ["w:bibliography", "bibliography"],
  ["w:docPartObj", "docPartObj"],
  ["w:docPartList", "docPartList"],
  ["w14:checkbox", "checkbox"],
  ["w15:repeatingSection", "repeatingSection"],
  ["w15:repeatingSectionItem", "repeatingSectionItem"],
  ["w14:entityPicker", "entityPicker"],
  ["w15:webExtensionLinked", "webExtensionLinked"],
  ["w15:webExtensionCreated", "webExtensionCreated"],
  ["w15:appearance", "appearance"],
]);

import type { DocxReadContext } from "../../context";

/**
 * Parse w:sdtPr element into SdtPropertiesOptions.
 */
export function parseSdtProperties(el: Element): SdtPropertiesOptions {
  const opts: Partial<SdtPropertiesOptions> = {};

  // rPr — the SDT start mark's run properties (CT_SdtPr's leading element)
  const rPr = findChild(el, "w:rPr");
  if (rPr) opts.runProperties = parseRunProperties(rPr);

  const alias = findChild(el, "w:alias");
  if (alias) opts.alias = attr(alias, "w:val");

  const tag = findChild(el, "w:tag");
  if (tag) {
    // CT_Tag requires w:val; an empty string is legal and must round-trip.
    opts.tag = attr(tag, "w:val") ?? "";
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

  const placeholder = findChild(el, "w:placeholder");
  if (placeholder) {
    const docPart = findChild(placeholder, "w:docPart");
    opts.placeholder = docPart ? { docPart: attr(docPart, "w:val") ?? "" } : {};
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
      // required by CT_DataBinding; tolerate a degenerate source without one
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
    if (dateFormat) dateOpts.dateFormat = attr(dateFormat, "w:val");
    const lid = findChild(date, "w:lid");
    if (lid) dateOpts.languageId = attr(lid, "w:val");
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
    const dpObj: NonNullable<SdtPropertiesOptions["docPartObj"]> = {};
    const gallery = findChild(dp, "w:docPartGallery");
    if (gallery) dpObj.gallery = attr(gallery, "w:val");
    const category = findChild(dp, "w:docPartCategory");
    if (category) dpObj.category = attr(category, "w:val");
    if (findChild(dp, "w:docPartUnique")) dpObj.unique = true;
    opts.docPartObj = dpObj;
  } else if (findChild(el, "w:docPartList")) {
    const dp = findChild(el, "w:docPartList")!;
    const dpObj: NonNullable<SdtPropertiesOptions["docPartList"]> = {};
    const gallery = findChild(dp, "w:docPartGallery");
    if (gallery) dpObj.gallery = attr(gallery, "w:val");
    const category = findChild(dp, "w:docPartCategory");
    if (category) dpObj.category = attr(category, "w:val");
    if (findChild(dp, "w:docPartUnique")) dpObj.unique = true;
    opts.docPartList = dpObj;
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
  } else if (findChild(el, "w15:repeatingSection")) {
    const rs = findChild(el, "w15:repeatingSection")!;
    const rsOpts: NonNullable<SdtPropertiesOptions["repeatingSection"]> = {};
    const sectionTitle = findChild(rs, "w15:sectionTitle");
    if (sectionTitle) rsOpts.sectionTitle = attr(sectionTitle, "w:val");
    const doNotAllow = findChild(rs, "w15:doNotAllowInsertDeleteSection");
    if (doNotAllow) rsOpts.doNotAllowInsertDeleteSection = attrBool(doNotAllow, "w:val") ?? true;
    opts.repeatingSection = rsOpts;
  } else if (findChild(el, "w15:repeatingSectionItem")) {
    opts.repeatingSectionItem = true;
  } else if (findChild(el, "w14:entityPicker")) {
    opts.entityPicker = true;
  }

  // Word 2013+ extension flags (CT_OnOff) — siblings outside the type choice.
  const webExtLinked = findChild(el, "w15:webExtensionLinked");
  if (webExtLinked) opts.webExtensionLinked = attrBool(webExtLinked, "w:val") ?? true;
  const webExtCreated = findChild(el, "w15:webExtensionCreated");
  if (webExtCreated) opts.webExtensionCreated = attrBool(webExtCreated, "w:val") ?? true;

  // w15:appearance (Word 2013+) — ST_SdtAppearance: boundingBox/tags/hidden
  const appearance = findChild(el, "w15:appearance");
  if (appearance) {
    const val = attr(appearance, "w15:val");
    if (val === "boundingBox" || val === "tags" || val === "hidden") {
      opts.appearance = val;
    }
  }

  const childOrder = (el.elements ?? []).flatMap((child) =>
    child.name === undefined ? [] : (SDT_ELEMENT_ORDER.get(child.name) ?? []),
  );
  if (childOrder.length > 0) opts.childOrder = childOrder;

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
  endProperties?: RunPropertiesOptions;
  children?: unknown[];
} {
  const sdtPr = findChild(el, "w:sdtPr");
  const properties = sdtPr ? parseSdtProperties(sdtPr) : {};

  // sdtEndPr (CT_SdtEndPr wraps its run properties in a w:rPr child)
  const sdtEndPr = findChild(el, "w:sdtEndPr");
  const endRPr = sdtEndPr ? findChild(sdtEndPr, "w:rPr") : undefined;
  const endProperties = sdtEndPr ? (endRPr ? parseRunProperties(endRPr) : {}) : undefined;

  const sdtContent = findChild(el, "w:sdtContent");
  let childList: unknown[] | undefined;
  if (sdtContent) {
    childList = parseChildren(sdtContent.elements ?? [], ctx);
    if (childList.length === 0) childList = undefined;
  }

  return { properties, endProperties, children: childList };
}
