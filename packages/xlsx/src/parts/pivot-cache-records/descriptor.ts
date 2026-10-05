import type { CustomDescriptor, ReadContext, WriteContext } from "@office-open/core/descriptor";
import type { Element } from "@office-open/xml";

import type { PivotSourceData } from "../pivot/pivot-utils";
import { parsePivotCacheRecords } from "./parse";
import { stringifyPivotCacheRecords, stringifyPivotCacheRecordsOptions } from "./stringify";
import type { PivotCacheRecordsOptions } from "./types";

export type PivotCacheRecordsDescriptorOptions =
  | { sourceData: PivotSourceData }
  | { records: PivotCacheRecordsOptions };

export const pivotCacheRecordsDesc: CustomDescriptor<
  PivotCacheRecordsDescriptorOptions,
  WriteContext,
  PivotCacheRecordsOptions
> = {
  kind: "custom",
  stringify(opts) {
    return "sourceData" in opts
      ? stringifyPivotCacheRecords(opts.sourceData)
      : stringifyPivotCacheRecordsOptions(opts.records);
  },
  parse(el: Element, ctx: ReadContext) {
    const part = (ctx as { currentPart?: string }).currentPart ?? "pivotCacheRecords.xml";
    return parsePivotCacheRecords(part, el);
  },
};
