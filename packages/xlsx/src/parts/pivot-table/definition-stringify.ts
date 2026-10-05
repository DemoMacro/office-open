import type { PivotXmlElementOptions } from "../pivot/pivot-xml";
import { stringifyPivotXmlElement } from "../pivot/pivot-xml";
import type { PivotTableDefinitionOptions } from "./definition-types";

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

export function stringifyPivotTableDefinitionOptions(options: PivotTableDefinitionOptions): string {
  if (options.source) return options.source;
  const childKeySet = new Set<string>(childNames);
  childKeySet.add("source");
  childKeySet.add("sourcePath");
  childKeySet.add("cacheRelationshipTarget");
  childKeySet.add("cacheRelationshipType");
  const root: PivotXmlElementOptions = {
    name: "pivotTableDefinition",
    attributes: {
      xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    },
    children: childNames
      .map((name) => options[name])
      .filter((child): child is PivotXmlElementOptions => !!child),
  };
  for (const [key, value] of Object.entries(options)) {
    if (value === undefined || childKeySet.has(key)) continue;
    root.attributes![key === "cacheRelationshipId" ? "r:id" : key] = value;
  }
  return stringifyPivotXmlElement(root);
}
