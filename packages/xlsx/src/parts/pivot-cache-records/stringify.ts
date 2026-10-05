import { escapeXml } from "@office-open/xml";

import { profilePivotFields } from "../pivot/pivot-utils";
import type { PivotSourceData } from "../pivot/pivot-utils";
import type { PivotCacheRecordEntry, PivotCacheRecordsOptions } from "./types";

export function stringifyPivotCacheRecords(sourceData: PivotSourceData): string {
  const profiles = profilePivotFields(sourceData);
  const fieldIndexMaps = profiles.map((field) => {
    if (field.numeric) return new Map<string, number>();
    const map = new Map<string, number>();
    for (let j = 0; j < field.unique.length; j++) map.set(String(field.unique[j]), j);
    return map;
  });
  const parts: string[] = [
    `<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sourceData.records.length}">`,
  ];
  for (const row of sourceData.records) {
    parts.push("<r>");
    for (let i = 0; i < row.length; i++) {
      const value = row[i];
      if (value === null) parts.push("<m/>");
      else if (value instanceof Date)
        parts.push(`<d v="${value.toISOString().replace(/\.\d{3}Z$/, "Z")}"/>`);
      else if (profiles[i]!.numeric) parts.push(`<n v="${value}"/>`);
      else parts.push(`<x v="${fieldIndexMaps[i]?.get(String(value)) ?? 0}"/>`);
    }
    parts.push("</r>");
  }
  parts.push("</pivotCacheRecords>");
  return parts.join("");
}

export function stringifyPivotCacheRecordsOptions(options: PivotCacheRecordsOptions): string {
  if (options.source) return options.source;
  const parts: string[] = [
    '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
  ];
  if (options.count !== undefined) parts.push(` count="${options.count}"`);
  parts.push(">");
  for (const record of options.records) {
    parts.push("<r>");
    for (const entry of record.entries) parts.push(stringifyEntry(entry));
    parts.push("</r>");
  }
  parts.push("</pivotCacheRecords>");
  return parts.join("");
}

function stringifyEntry(entry: PivotCacheRecordEntry): string {
  const elementName =
    entry.type === "sharedItem"
      ? "x"
      : entry.type === "date"
        ? "d"
        : entry.type === "boolean"
          ? "b"
          : entry.type === "string"
            ? "s"
            : entry.type[0]!;
  const attrs: string[] = [];
  if (entry.value !== undefined) attrs.push(`v="${escapeXml(String(entry.value))}"`);
  if (entry.unformatted) attrs.push('u="1"');
  if (entry.formula) attrs.push('f="1"');
  if (entry.caption !== undefined) attrs.push(`c="${escapeXml(entry.caption)}"`);
  if (entry.captionLength !== undefined) attrs.push(`cp="${entry.captionLength}"`);
  return `<${elementName}${attrs.length ? ` ${attrs.join(" ")}` : ""}/>`;
}
