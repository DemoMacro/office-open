import type { Element } from "@office-open/xml";

import {
  findPivotChildren,
  parsePivotXmlElement,
  stringifyPivotXmlElement,
} from "../pivot/pivot-xml";
import type {
  PivotCacheRecordEntry,
  PivotCacheRecordOptions,
  PivotCacheRecordsOptions,
} from "./types";

export function parsePivotCacheRecords(part: string, el: Element): PivotCacheRecordsOptions {
  const root = parsePivotXmlElement(part, el);
  const records = findPivotChildren(root, "r").map((record) => {
    const parsed: PivotCacheRecordOptions = {
      entries: (record.children ?? [])
        .map(parseEntry)
        .filter((entry): entry is PivotCacheRecordEntry => !!entry),
    };
    parsed.extLst = findPivotChildren(record, "extLst")[0];
    return parsed;
  });
  const result: PivotCacheRecordsOptions = {
    source: stringifyPivotXmlElement(root),
    records,
  };
  const count = root.attributes?.count;
  if (count !== undefined) result.count = Number(count);
  result.extLst = findPivotChildren(root, "extLst")[0];
  return result;
}

function parseEntry(
  el: ReturnType<typeof parsePivotXmlElement>,
): PivotCacheRecordEntry | undefined {
  const attributes = el.attributes ?? {};
  const base = {
    ...(attributes.u !== undefined ? { unformatted: attributes.u === "1" } : {}),
    ...(attributes.f !== undefined ? { formula: attributes.f === "1" } : {}),
    ...(attributes.c !== undefined ? { caption: String(attributes.c) } : {}),
    ...(attributes.cp !== undefined ? { captionLength: Number(attributes.cp) } : {}),
    tuples: findPivotChildren(el, "tpls"),
    indexes: findPivotChildren(el, "x"),
  };
  const value = attributes.v;
  switch (el.name) {
    case "m":
      return { type: "missing", ...base };
    case "n":
      return { type: "number", value: Number(value), ...base };
    case "b":
      return { type: "boolean", value: value === "1", ...base };
    case "e":
      return { type: "error", value: String(value), ...base };
    case "s":
      return { type: "string", value: String(value), ...base };
    case "d":
      return { type: "date", value: String(value), ...base };
    case "x":
      return { type: "sharedItem", value: Number(value), ...base };
    default:
      return undefined;
  }
}
