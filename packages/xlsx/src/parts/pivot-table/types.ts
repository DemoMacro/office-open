/**
 * PivotTable — parse-result types for CT_pivotTableDefinition.
 *
 * @module
 */

import type { PivotAreaOptions, PivotSourceData, PivotTableOptions } from "../pivot/pivot-utils";

// ── Types ──

export interface PivotTableDescriptorOptions {
  options: PivotTableOptions;
  sourceData: PivotSourceData;
  cacheId: number;
}

/** Parsed pivotField (CT_PivotField) — index-based, inspect-only. */
export interface PivotFieldParseResult {
  axis?: string;
  showAll?: boolean;
  dataField?: boolean;
  hierarchy?: string;
  dragToRow?: boolean;
  dragToCol?: boolean;
  dragToPage?: boolean;
  dragToData?: boolean;
  dragOff?: boolean;
  showDropDowns?: boolean;
  insertBlankRow?: boolean;
  showPropCell?: boolean;
  showPropTip?: boolean;
  showPropAsCaption?: boolean;
  compact?: boolean;
  outline?: boolean;
  subtotalTop?: boolean;
  includeNewItemsInFilter?: boolean;
  avgSubtotal?: boolean;
  countASubtotal?: boolean;
  maxSubtotal?: boolean;
  minSubtotal?: boolean;
  sumSubtotal?: boolean;
}

/** Parsed dataField (CT_DataField). */
export interface DataFieldParseResult {
  name?: string;
  fld?: number;
  subtotal?: string;
  showDataAs?: string;
  baseField?: number;
  baseItem?: number;
  numFmtId?: string;
}

/** Parsed pageField (CT_PageField). */
export interface PageFieldParseResult {
  fld?: number;
  hier?: number;
  cap?: string;
  item?: number;
}

/** Parsed format entry (CT_Format). */
export interface PivotFormatParseResult {
  action?: string;
  dxfId?: number;
  pivotArea?: Partial<PivotAreaOptions>;
}

/** Parsed conditionalFormat (CT_ConditionalFormat) entry. */
export interface ConditionalFormatParseResult {
  scope?: string;
  type?: string;
  priority?: number;
  pivotAreas?: Partial<PivotAreaOptions>[];
}

/** Parsed chartFormat (CT_ChartFormat). */
export interface ChartFormatParseResult {
  chart?: number;
  format?: number;
  series?: boolean;
  pivotArea?: Partial<PivotAreaOptions>;
}

/** Parsed pivotHierarchy (CT_PivotHierarchy). */
export interface PivotHierarchyParseResult {
  outline?: boolean;
  multipleItemSelectionAllowed?: boolean;
  subtotalTop?: boolean;
  showInFieldList?: boolean;
  dragToRow?: boolean;
  dragToCol?: boolean;
  dragToPage?: boolean;
  dragToData?: boolean;
  dragOff?: boolean;
  includeNewItemsInFilter?: boolean;
  caption?: string;
}

/** Parsed pivotFilter (CT_PivotFilter). */
export interface PivotFilterParseResult {
  fld?: number;
  type?: string;
  id?: number;
  mpFld?: number;
  evalOrder?: number;
  iMeasureHier?: number;
  iMeasureFld?: number;
  stringValue1?: string;
  stringValue2?: string;
}

/** Parsed row/col hierarchyUsage entry (CT_RowColHierarchyUsage). */
export interface HierarchyUsageParseResult {
  hierarchyUsage: number;
}

/** Parsed calculatedItem (CT_CalculatedItem). */
export interface CalculatedItemParseResult {
  field?: number;
  formula?: string;
  pivotArea?: Partial<PivotAreaOptions>;
}

/** Parsed calculatedMember (CT_CalculatedMember). */
export interface CalculatedMemberParseResult {
  name?: string;
  mdx?: string;
  memberName?: string;
  hierarchy?: string;
  parent?: string;
  solveOrder?: number;
  set?: boolean;
}

/**
 * Structured output of {@link pivotTableDesc}.parse — CT_pivotTableDefinition as
 * parsed (index-based pivotFields/dataFields, flat attributes). This is the
 * CT-layer shape, NOT {@link PivotTableOptions} (the name-based user shape),
 * so it cannot be fed straight back into stringify; the compiler regenerates a
 * pivot table from `sourceData` + user `PivotTableOptions` instead.
 */
export interface PivotTableParseResult {
  name?: string;
  cacheId?: number;
  dataOnRows?: boolean;
  showHeaders?: boolean;
  showEmptyRow?: boolean;
  showEmptyCol?: boolean;
  grandTotalCaption?: string;
  errorCaption?: string;
  showError?: boolean;
  missingCaption?: string;
  showMissing?: boolean;
  pageStyle?: string;
  pivotTableStyle?: string;
  tag?: string;
  showItems?: boolean;
  editData?: boolean;
  disableFieldList?: boolean;
  showCalcMbrs?: boolean;
  visualTotals?: boolean;
  showMultipleLabel?: boolean;
  showDataDropDown?: boolean;
  showDrill?: boolean;
  printDrill?: boolean;
  showMemberPropertyTips?: boolean;
  showDataTips?: boolean;
  enableWizard?: boolean;
  enableDrill?: boolean;
  enableFieldProperties?: boolean;
  pageWrap?: number;
  pageOverThenDown?: boolean;
  subtotalHiddenItems?: boolean;
  fieldPrintTitles?: boolean;
  mergeItem?: boolean;
  showDropZones?: boolean;
  published?: boolean;
  gridDropZones?: boolean;
  multipleFieldFilters?: boolean;
  rowHeaderCaption?: string;
  colHeaderCaption?: string;
  fieldListSortAscending?: boolean;
  mdxSubqueries?: boolean;
  customListSort?: boolean;
  asteriskTotals?: boolean;
  dataPosition?: number;
  immersive?: boolean;
  vacatedStyle?: string;
  chartFormat?: boolean;
  preserveFormatting?: boolean;
  dataCaption?: string;
  location?: string;
  locationRowPageCount?: number;
  locationColPageCount?: number;
  pivotFields?: PivotFieldParseResult[];
  dataFields?: DataFieldParseResult[];
  rowFields?: number[];
  colFields?: number[];
  pageFields?: PageFieldParseResult[];
  formats?: PivotFormatParseResult[];
  conditionalFormats?: ConditionalFormatParseResult[];
  chartFormats?: ChartFormatParseResult[];
  pivotHierarchies?: PivotHierarchyParseResult[];
  filters?: PivotFilterParseResult[];
  rowHierarchiesUsage?: HierarchyUsageParseResult[];
  colHierarchiesUsage?: HierarchyUsageParseResult[];
  calculatedItems?: CalculatedItemParseResult[];
  calculatedMembers?: CalculatedMemberParseResult[];
  style?: string;
}
