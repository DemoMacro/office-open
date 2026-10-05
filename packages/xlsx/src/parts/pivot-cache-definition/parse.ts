import { attrNum } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import {
  findPivotChildren,
  parsePivotXmlElement,
  stringifyPivotXmlElement,
  type PivotXmlElementOptions,
} from "../pivot/pivot-xml";
import { parseCalculatedItems, parseCalculatedMembers } from "./calculated";
import type {
  PivotCacheDefinitionOptions,
  PivotCacheFieldAttributes,
  PivotCacheFieldOptions,
  PivotSharedItemsAttributes,
} from "./types";

const booleanAttributes = [
  "invalid",
  "saveData",
  "refreshOnLoad",
  "optimizeMemory",
  "enableRefresh",
  "backgroundQuery",
  "upgradeOnRefresh",
  "tupleCache",
  "supportSubquery",
  "supportAdvancedDrill",
] as const;

const numberAttributes = [
  "refreshedDate",
  "missingItemsLimit",
  "createdVersion",
  "refreshedVersion",
  "minRefreshableVersion",
  "recordCount",
] as const;

const cacheFieldNumberAttributes = [
  "numFmtId",
  "sqlType",
  "hierarchy",
  "level",
  "mappingCount",
] as const;

const sharedItemsBooleanAttributes = [
  "containsSemiMixedTypes",
  "containsNonDate",
  "containsDate",
  "containsString",
  "containsBlank",
  "containsMixedTypes",
  "containsNumber",
  "containsInteger",
  "longText",
] as const;

const sharedItemsNumberAttributes = ["minValue", "maxValue", "count"] as const;

export function parsePivotCacheDefinition(part: string, el: Element): PivotCacheDefinitionOptions {
  const root = parsePivotXmlElement(part, el);
  const result: PivotCacheDefinitionOptions = { source: stringifyPivotXmlElement(root) };
  const attributes = root.attributes ?? {};
  for (const name of booleanAttributes) {
    if (attributes[name] !== undefined) result[name] = attributes[name] === "1";
  }
  for (const name of numberAttributes) {
    const value = attrNum(el, name);
    if (value !== undefined) result[name] = value;
  }
  for (const name of ["refreshedBy", "refreshedDateIso"] as const) {
    if (attributes[name] !== undefined) result[name] = String(attributes[name]);
  }
  if (attributes["r:id"] !== undefined) result.recordsRelationshipId = String(attributes["r:id"]);

  const child = (name: string): PivotXmlElementOptions | undefined =>
    findPivotChildren(root, name)[0];
  result.cacheSource = child("cacheSource");
  const fieldsEl = child("cacheFields");
  if (fieldsEl) {
    result.cacheFields = {
      ...(fieldsEl.attributes?.count !== undefined
        ? { count: Number(fieldsEl.attributes?.count) }
        : {}),
      fields: findPivotChildren(fieldsEl, "cacheField").map((field) => parseCacheField(field)),
    };
  }
  result.tupleCacheData = child("tupleCache");
  const calculatedItems = child("calculatedItems");
  if (calculatedItems) {
    result.calculatedItems = parseCalculatedItems(part, "/pivotCacheDefinition", calculatedItems);
  }
  const calculatedMembers = child("calculatedMembers");
  if (calculatedMembers) {
    result.calculatedMembers = parseCalculatedMembers(
      part,
      "/pivotCacheDefinition",
      calculatedMembers,
    );
  }
  result.maps = child("maps");
  result.extLst = child("extLst");
  return result;
}

function parseCacheField(field: PivotXmlElementOptions): PivotCacheFieldOptions {
  const attributes = field.attributes ?? {};
  const result: PivotCacheFieldAttributes = {};
  for (const name of ["name", "caption", "propertyName", "formula"] as const) {
    if (attributes[name] !== undefined) result[name] = String(attributes[name]);
  }
  for (const name of [
    "serverField",
    "uniqueList",
    "databaseField",
    "memberPropertyField",
  ] as const) {
    if (attributes[name] !== undefined) result[name] = attributes[name] === "1";
  }
  for (const name of cacheFieldNumberAttributes) {
    const value = attributes[name];
    if (value !== undefined) result[name] = Number(value);
  }
  const parsed: PivotCacheFieldOptions = result;
  const sharedItems = findPivotChildren(field, "sharedItems")[0];
  if (sharedItems) {
    parsed.sharedItems = sharedItems;
    parsed.sharedItemsAttributes = parseSharedItemsAttributes(sharedItems);
  }
  parsed.fieldGroup = findPivotChildren(field, "fieldGroup")[0];
  parsed.mpMaps = findPivotChildren(field, "mpMap");
  parsed.extLst = findPivotChildren(field, "extLst")[0];
  return parsed;
}

function parseSharedItemsAttributes(el: PivotXmlElementOptions): PivotSharedItemsAttributes {
  const attributes = el.attributes ?? {};
  const result: PivotSharedItemsAttributes = {};
  for (const name of sharedItemsBooleanAttributes) {
    if (attributes[name] !== undefined) result[name] = attributes[name] === "1";
  }
  for (const name of sharedItemsNumberAttributes) {
    if (attributes[name] !== undefined) result[name] = Number(attributes[name]);
  }
  for (const name of ["minDate", "maxDate"] as const) {
    if (attributes[name] !== undefined) result[name] = String(attributes[name]);
  }
  return result;
}
