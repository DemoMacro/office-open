import type { CustomDescriptor, ReadContext, WriteContext } from "@office-open/core/descriptor";
import type { Element } from "@office-open/xml";

import type { PivotSourceData } from "../pivot/pivot-utils";
import { parsePivotCacheDefinition } from "./parse";
import { stringifyPivotCacheDefinition, stringifyPivotCacheDefinitionOptions } from "./stringify";
import type { PivotCacheDefinitionOptions } from "./types";

export type PivotCacheDefDescriptorOptions =
  | {
      sourceRef: string;
      sourceSheet: string;
      sourceData: PivotSourceData;
      recordsRid: string;
      cacheDefOpts?: PivotCacheDefinitionOptions;
    }
  | { definition: PivotCacheDefinitionOptions };

export const pivotCacheDefDesc: CustomDescriptor<
  PivotCacheDefDescriptorOptions,
  WriteContext,
  PivotCacheDefinitionOptions
> = {
  kind: "custom",
  stringify(opts) {
    return "definition" in opts
      ? stringifyPivotCacheDefinitionOptions(opts.definition)
      : stringifyPivotCacheDefinition(
          opts.sourceRef,
          opts.sourceSheet,
          opts.sourceData,
          opts.recordsRid,
          opts.cacheDefOpts,
        );
  },
  parse(el: Element, ctx: ReadContext) {
    const part = (ctx as { currentPart?: string }).currentPart ?? "pivotCacheDefinition.xml";
    return parsePivotCacheDefinition(part, el);
  },
};
