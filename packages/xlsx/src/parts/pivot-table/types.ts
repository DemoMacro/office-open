import type { PivotSourceData, SourcePivotTableOptions } from "../pivot/pivot-utils";
import type { PivotTableDefinitionOptions } from "./definition-types";

export type PivotTableDescriptorOptions =
  | {
      options: SourcePivotTableOptions;
      sourceData: PivotSourceData;
      cacheId: number;
    }
  | { definition: PivotTableDefinitionOptions };
