/**
 * Pivot table module — exports all pivot table types and utilities.
 *
 * XML generation has been migrated to descriptors (compile/descriptors/).
 *
 * @module
 */
export type {
  PivotCacheDefinitionOptions,
  CacheFieldExtraAttrs,
  PivotTableOptions,
  PivotDataField,
  PivotSourceData,
  ConsolidateFunction,
  PivotFilterOptions,
  PivotFilterType,
  OlapPrOptions,
} from "./pivot-utils";
export {
  ConsolidateFunction as ConsolidateFunctionType,
  PivotFilterType as PivotFilterTypeValue,
  collectUniqueValues,
  isNumericField,
  aggregate,
} from "./pivot-utils";
