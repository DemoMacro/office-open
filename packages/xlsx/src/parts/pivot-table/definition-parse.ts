import type { Element } from "@office-open/xml";

import {
  findPivotChildren,
  parsePivotXmlElement,
  stringifyPivotXmlElement,
} from "../pivot/pivot-xml";
import type { PivotTableDefinitionOptions } from "./definition-types";

const booleanAttributes = [
  "dataOnRows",
  "showError",
  "showMissing",
  "asteriskTotals",
  "showItems",
  "editData",
  "disableFieldList",
  "showCalcMbrs",
  "visualTotals",
  "showMultipleLabel",
  "showDataDropDown",
  "showDrill",
  "printDrill",
  "showMemberPropertyTips",
  "showDataTips",
  "enableWizard",
  "enableDrill",
  "enableFieldProperties",
  "preserveFormatting",
  "useAutoFormatting",
  "pageOverThenDown",
  "subtotalHiddenItems",
  "rowGrandTotals",
  "colGrandTotals",
  "fieldPrintTitles",
  "itemPrintTitles",
  "mergeItem",
  "showDropZones",
  "showEmptyRow",
  "showEmptyCol",
  "showHeaders",
  "compact",
  "outline",
  "outlineData",
  "compactData",
  "published",
  "gridDropZones",
  "immersive",
  "multipleFieldFilters",
  "fieldListSortAscending",
  "mdxSubqueries",
  "customListSort",
] as const;
const numberAttributes = [
  "cacheId",
  "dataPosition",
  "pageWrap",
  "indent",
  "chartFormat",
  "updatedVersion",
  "minRefreshableVersion",
  "createdVersion",
] as const;
const stringAttributes = [
  "name",
  "dataCaption",
  "grandTotalCaption",
  "errorCaption",
  "missingCaption",
  "pageStyle",
  "pivotTableStyle",
  "vacatedStyle",
  "tag",
  "rowHeaderCaption",
  "colHeaderCaption",
] as const;
const childNames = [
  "location",
  "pivotFields",
  "rowFields",
  "rowItems",
  "colFields",
  "colItems",
  "pageFields",
  "dataFields",
  "formats",
  "conditionalFormats",
  "chartFormats",
  "pivotHierarchies",
  "pivotTableStyleInfo",
  "filters",
  "rowHierarchiesUsage",
  "colHierarchiesUsage",
  "extLst",
] as const;

export function parsePivotTableDefinition(part: string, el: Element): PivotTableDefinitionOptions {
  const root = parsePivotXmlElement(part, el);
  const result: PivotTableDefinitionOptions = {
    source: stringifyPivotXmlElement(root),
    sourcePath: part,
  };
  const attributes = root.attributes ?? {};
  for (const name of booleanAttributes) {
    if (attributes[name] !== undefined) result[name] = attributes[name] === "1";
  }
  for (const name of numberAttributes) {
    if (attributes[name] !== undefined) result[name] = Number(attributes[name]);
  }
  for (const name of stringAttributes) {
    if (attributes[name] !== undefined) result[name] = String(attributes[name]);
  }
  for (const name of childNames) {
    const child = findPivotChildren(root, name)[0];
    if (child) result[name] = child;
  }
  return result;
}
