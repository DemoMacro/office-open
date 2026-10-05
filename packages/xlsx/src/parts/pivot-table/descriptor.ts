import type { CustomDescriptor, ReadContext, WriteContext } from "@office-open/core/descriptor";
import type { Element } from "@office-open/xml";

import { parsePivotTableDefinition } from "./definition-parse";
import { stringifyPivotTableDefinitionOptions } from "./definition-stringify";
import type { PivotTableDefinitionOptions } from "./definition-types";
import { stringifyPivotTable } from "./stringify";
import type { PivotTableDescriptorOptions } from "./types";

export const pivotTableDesc: CustomDescriptor<
  PivotTableDescriptorOptions,
  WriteContext,
  PivotTableDefinitionOptions
> = {
  kind: "custom",
  stringify(opts) {
    return "definition" in opts
      ? stringifyPivotTableDefinitionOptions(opts.definition)
      : stringifyPivotTable(opts.options, opts.sourceData, opts.cacheId);
  },
  parse(el: Element, ctx: ReadContext) {
    const part = (ctx as { currentPart?: string }).currentPart ?? "pivotTableDefinition.xml";
    return parsePivotTableDefinition(part, el);
  },
};
