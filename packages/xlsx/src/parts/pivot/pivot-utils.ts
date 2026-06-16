/**
 * Pivot table utility types and helper functions.
 *
 * @module
 */

/** Aggregation function for data fields (maps to ST_DataConsolidateFunction). */
export const ConsolidateFunction = {
  SUM: "sum",
  COUNT: "count",
  AVERAGE: "average",
  MAX: "max",
  MIN: "min",
  PRODUCT: "product",
  COUNT_NUMS: "countNums",
  STD_DEV: "stdDev",
  STD_DEV_P: "stdDevp",
  VAR: "var",
  VAR_P: "varp",
} as const;

export type ConsolidateFunction = (typeof ConsolidateFunction)[keyof typeof ConsolidateFunction];

/** A data field definition for pivot table aggregation. */
export interface PivotDataField {
  /** Source column name to aggregate */
  field: string;
  /** Aggregation function (default: "sum") */
  summarize?: ConsolidateFunction;
  /** Custom name for the data field (default: "Sum of {field}") */
  name?: string;
  /** Show data as (CT_DataField @showDataAs) */
  showDataAs?: string;
  /** Base field index for "show data as" calculations */
  baseField?: number;
  /** Base item index for "show data as" calculations */
  baseItem?: number;
  /** Sort by tuple items (CT_Tuples in sortByTuple) */
  sortByTupleItems?: number[];
}

/** Per-field overrides for pivotField XML attributes (CT_PivotField). */
export interface PivotFieldOverrideOptions {
  /** Field name to match (required) */
  field: string;
  /** All drilled (CT_PivotField @allDrilled) */
  allDrilled?: boolean;
  /** Auto show (CT_PivotField @autoShow) */
  autoShow?: boolean;
  /** Count subtotal (CT_PivotField @countSubtotal) */
  countSubtotal?: boolean;
  /** Data source sort (CT_PivotField @dataSourceSort) */
  dataSourceSort?: boolean;
  /** Default attribute drill state (CT_PivotField @defaultAttributeDrillState) */
  defaultAttributeDrillState?: boolean;
  /** Hidden level (CT_PivotField @hiddenLevel) */
  hiddenLevel?: boolean;
  /** Hide new items (CT_PivotField @hideNewItems) */
  hideNewItems?: boolean;
  /** Insert blank row (CT_PivotField @insertBlankRow) */
  insertBlankRow?: boolean;
  /** Insert page break (CT_PivotField @insertPageBreak) */
  insertPageBreak?: boolean;
  /** Item page count (CT_PivotField @itemPageCount) */
  itemPageCount?: boolean;
  /** Measure filter (CT_PivotField @measureFilter) */
  measureFilter?: boolean;
  /** Non auto sort default (CT_PivotField @nonAutoSortDefault) */
  nonAutoSortDefault?: boolean;
  /** Product subtotal (CT_PivotField @productSubtotal) */
  productSubtotal?: boolean;
  /** Rank by (CT_PivotField @rankBy) */
  rankBy?: number;
  /** Server field (CT_PivotField @serverField) */
  serverField?: boolean;
  /** Show drop downs (CT_PivotField @showDropDowns) */
  showDropDowns?: boolean;
  /** Show property as caption (CT_PivotField @showPropAsCaption) */
  showPropAsCaption?: boolean;
  /** Show property cell (CT_PivotField @showPropCell) */
  showPropCell?: boolean;
  /** Show property tip (CT_PivotField @showPropTip) */
  showPropTip?: boolean;
  /** StdDevP subtotal (CT_PivotField @stdDevPSubtotal) */
  stdDevPSubtotal?: boolean;
  /** StdDev subtotal (CT_PivotField @stdDevSubtotal) */
  stdDevSubtotal?: boolean;
  /** Subtotal caption (CT_PivotField @subtotalCaption) */
  subtotalCaption?: string;
  /** Top auto show (CT_PivotField @topAutoShow) */
  topAutoShow?: boolean;
  /** Unique member property (CT_PivotField @uniqueMemberProperty) */
  uniqueMemberProperty?: boolean;
  /** VarP subtotal (CT_PivotField @varPSubtotal) */
  varPSubtotal?: boolean;
  /** Var subtotal (CT_PivotField @varSubtotal) */
  varSubtotal?: boolean;
  /** Show detail for default item (CT_Item @sd) */
  defaultItemSd?: boolean;
}

/** Pivot filter type (ST_PivotFilterType) */
export const PivotFilterType = {
  UNKNOWN: "unknown",
  COUNT: "count",
  PERCENT: "percent",
  SUM: "sum",
  CAPTION_EQUAL: "captionEqual",
  CAPTION_NOT_EQUAL: "captionNotEqual",
  CAPTION_BEGINS_WITH: "captionBeginsWith",
  CAPTION_NOT_BEGINS_WITH: "captionNotBeginsWith",
  CAPTION_ENDS_WITH: "captionEndsWith",
  CAPTION_NOT_ENDS_WITH: "captionNotEndsWith",
  CAPTION_CONTAINS: "captionContains",
  CAPTION_NOT_CONTAINS: "captionNotContains",
  CAPTION_GREATER_THAN: "captionGreaterThan",
  CAPTION_GREATER_THAN_OR_EQUAL: "captionGreaterThanOrEqual",
  CAPTION_LESS_THAN: "captionLessThan",
  CAPTION_LESS_THAN_OR_EQUAL: "captionLessThanOrEqual",
  CAPTION_BETWEEN: "captionBetween",
  CAPTION_NOT_BETWEEN: "captionNotBetween",
  VALUE_EQUAL: "valueEqual",
  VALUE_NOT_EQUAL: "valueNotEqual",
  VALUE_GREATER_THAN: "valueGreaterThan",
  VALUE_GREATER_THAN_OR_EQUAL: "valueGreaterThanOrEqual",
  VALUE_LESS_THAN: "valueLessThan",
  VALUE_LESS_THAN_OR_EQUAL: "valueLessThanOrEqual",
  VALUE_BETWEEN: "valueBetween",
  VALUE_NOT_BETWEEN: "valueNotBetween",
  DATE_EQUAL: "dateEqual",
  DATE_NOT_EQUAL: "dateNotEqual",
  DATE_OLDER_THAN: "dateOlderThan",
  DATE_OLDER_THAN_OR_EQUAL: "dateOlderThanOrEqual",
  DATE_NEWER_THAN: "dateNewerThan",
  DATE_NEWER_THAN_OR_EQUAL: "dateNewerThanOrEqual",
  DATE_BETWEEN: "dateBetween",
  DATE_NOT_BETWEEN: "dateNotBetween",
  TOMORROW: "tomorrow",
  TODAY: "today",
  YESTERDAY: "yesterday",
  NEXT_WEEK: "nextWeek",
  THIS_WEEK: "thisWeek",
  LAST_WEEK: "lastWeek",
  NEXT_MONTH: "nextMonth",
  THIS_MONTH: "thisMonth",
  LAST_MONTH: "lastMonth",
  NEXT_QUARTER: "nextQuarter",
  THIS_QUARTER: "thisQuarter",
  LAST_QUARTER: "lastQuarter",
  NEXT_YEAR: "nextYear",
  THIS_YEAR: "thisYear",
  LAST_YEAR: "lastYear",
  YEAR_TO_DATE: "yearToDate",
  Q1: "Q1",
  Q2: "Q2",
  Q3: "Q3",
  Q4: "Q4",
  M1: "M1",
  M2: "M2",
  M3: "M3",
  M4: "M4",
  M5: "M5",
  M6: "M6",
  M7: "M7",
  M8: "M8",
  M9: "M9",
  M10: "M10",
  M11: "M11",
  M12: "M12",
} as const;

export type PivotFilterType = (typeof PivotFilterType)[keyof typeof PivotFilterType];

/** A single pivot filter (CT_PivotFilter) */
export interface PivotFilterOptions {
  /** Field index to filter on (required) */
  fld: number;
  /** Filter type (required) */
  type: PivotFilterType;
  /** Filter ID — unique within this pivot table (required) */
  id: number;
  /** Measure field index for OLAP filters */
  mpFld?: number;
  /** Evaluation order */
  evalOrder?: number;
  /** Measure hierarchy */
  iMeasureHier?: number;
  /** Measure field */
  iMeasureFld?: number;
  /** Filter name */
  name?: string;
  /** Filter description */
  description?: string;
  /** First string value for caption/date filters */
  stringValue1?: string;
  /** Second string value for between filters */
  stringValue2?: string;
}

/** Options for a single pivot table on a worksheet. */
export interface PivotTableOptions {
  /** Pivot table name (default: "PivotTable{N}") */
  name?: string;
  /** Source data range, e.g. "A1:D11" — must be on the same or a different sheet */
  source: string;
  /** Source sheet name (default: current sheet) */
  sourceSheet?: string;
  /** Target cell for the pivot table output (default: "A3") */
  location?: string;
  /** Field names to use as row labels */
  rows: string[];
  /** Field names to use as column labels */
  columns?: string[];
  /** Data fields with aggregation settings */
  data: PivotDataField[];
  /** Pivot style name (default: "PivotStyleLight16") */
  style?: string;
  /** Pivot filters (CT_PivotFilters) */
  filters?: PivotFilterOptions[];
  /** Field names to use as page/report filters */
  pages?: string[];
  /** Page field captions (maps by index to pages array) */
  pageCaptions?: string[];
  /** Data fields on rows instead of columns (CT_PivotTableDefinition @dataOnRows) */
  dataOnRows?: boolean;
  /** Grand total caption text */
  grandTotalCaption?: string;
  /** Error caption text */
  errorCaption?: string;
  /** Show error messages */
  showError?: boolean;
  /** Missing caption text */
  missingCaption?: string;
  /** Show missing items */
  showMissing?: boolean;
  /** Custom page style name */
  pageStyle?: string;
  /** Custom pivot table style name */
  pivotTableStyle?: string;
  /** Tag string */
  tag?: string;
  /** Show items with no data */
  showItems?: boolean;
  /** Edit data in-place */
  editData?: boolean;
  /** Disable field list */
  disableFieldList?: boolean;
  /** Show calculated members */
  showCalcMbrs?: boolean;
  /** Visual totals */
  visualTotals?: boolean;
  /** Show multiple labels */
  showMultipleLabel?: boolean;
  /** Show data drop-down */
  showDataDropDown?: boolean;
  /** Show drill indicators */
  showDrill?: boolean;
  /** Print drill indicators */
  printDrill?: boolean;
  /** Show member property tips */
  showMemberPropertyTips?: boolean;
  /** Show data tips */
  showDataTips?: boolean;
  /** Enable layout wizard */
  enableWizard?: boolean;
  /** Enable drill-down */
  enableDrill?: boolean;
  /** Enable field properties */
  enableFieldProperties?: boolean;
  /** Number of page fields per row/column */
  pageWrap?: number;
  /** Page layout over then down */
  pageOverThenDown?: boolean;
  /** Subtotal hidden items */
  subtotalHiddenItems?: boolean;
  /** Field print titles */
  fieldPrintTitles?: boolean;
  /** Merge item labels */
  mergeItem?: boolean;
  /** Show drop zones */
  showDropZones?: boolean;
  /** Show empty row */
  showEmptyRow?: boolean;
  /** Show empty column */
  showEmptyCol?: boolean;
  /** Show headers */
  showHeaders?: boolean;
  /** Published to server */
  published?: boolean;
  /** Grid drop zones */
  gridDropZones?: boolean;
  /** Multiple field filters */
  multipleFieldFilters?: boolean;
  /** Row header caption */
  rowHeaderCaption?: string;
  /** Column header caption */
  colHeaderCaption?: string;
  /** Sort field list ascending */
  fieldListSortAscending?: boolean;
  /** MDX subqueries enabled */
  mdxSubqueries?: boolean;
  /** Custom list sort */
  customListSort?: boolean;
  /** Asterisk totals (CT_PivotTableDefinition @asteriskTotals) */
  asteriskTotals?: boolean;
  /** Data position (CT_PivotTableDefinition @dataPosition) */
  dataPosition?: number;
  /** Immersive (CT_PivotTableDefinition @immersive) */
  immersive?: boolean;
  /** Vacated style (CT_PivotTableDefinition @vacatedStyle) */
  vacatedStyle?: string;
  /** Calculated items (CT_CalculatedItems) */
  calculatedItems?: CalculatedItemOptions[];
  /** Calculated members (CT_CalculatedMembers) */
  calculatedMembers?: CalculatedMemberOptions[];
  /** Pivot hierarchies (CT_PivotHierarchies) */
  pivotHierarchies?: PivotHierarchyOptions[];
  /** Conditional formats (CT_ConditionalFormats for pivot) */
  pivotConditionalFormats?: PivotConditionalFormatOptions[];
  /** Chart formats (CT_ChartFormats) */
  chartFormats?: ChartFormatOptions[];
  /** Auto sort scope (CT_AutoSortScope) */
  autoSortScope?: PivotAreaOptions;
  /** Member properties per field (CT_MemberProperties → mps/mp) */
  memberProperties?: MemberPropertyOptions[];
  /** Pivot format areas (CT_Formats → format) */
  formats?: PivotFormatOptions[];
  /** Row hierarchy usage (CT_RowHierarchiesUsage) */
  rowHierarchiesUsage?: HierarchyUsageOptions[];
  /** Column hierarchy usage (CT_ColHierarchiesUsage) */
  colHierarchiesUsage?: HierarchyUsageOptions[];
  /** Location column page count (CT_Location @colPageCount) */
  locationColPageCount?: number;
  /** Location row page count (CT_Location @rowPageCount) */
  locationRowPageCount?: number;
  /** Per-field overrides for pivotField (CT_PivotField attributes) */
  fieldOverrides?: PivotFieldOverrideOptions[];
}

/** Pivot format (CT_Format). */
export interface PivotFormatOptions {
  /** Action type (default: "formatting") */
  action?: "formatting" | "drill" | "formula" | "blank" | "subtotal" | "report";
  /** Differential format index */
  dxfId?: number;
  /** Pivot area */
  pivotArea: PivotAreaOptions;
}

/** Calculated item in a pivot table (CT_CalculatedItem) */
export interface CalculatedItemOptions {
  /** Field index */
  field?: number;
  /** Formula */
  formula?: string;
  /** Pivot area reference */
  pivotArea?: PivotAreaOptions;
}

/** Calculated member in a pivot table (CT_CalculatedMember) */
export interface CalculatedMemberOptions {
  /** Name (required) */
  name: string;
  /** MDX expression (required) */
  mdx: string;
  /** Member name */
  memberName?: string;
  /** Hierarchy */
  hierarchy?: string;
  /** Parent member */
  parent?: string;
  /** Solve order (default: 0) */
  solveOrder?: number;
  /** Is a set (default: false) */
  set?: boolean;
}

/** Pivot hierarchy (CT_PivotHierarchy) */
export interface PivotHierarchyOptions {
  /** Outline mode (default: false) */
  outline?: boolean;
  /** Allow multiple item selection (default: false) */
  multipleItemSelectionAllowed?: boolean;
  /** Subtotal on top (default: false) */
  subtotalTop?: boolean;
  /** Show in field list (default: true) */
  showInFieldList?: boolean;
  /** Drag to row (default: true) */
  dragToRow?: boolean;
  /** Drag to column (default: true) */
  dragToCol?: boolean;
  /** Drag to page (default: true) */
  dragToPage?: boolean;
  /** Drag to data (default: false) */
  dragToData?: boolean;
  /** Drag off (default: true) */
  dragOff?: boolean;
  /** Include new items in filter (default: false) */
  includeNewItemsInFilter?: boolean;
  /** Caption */
  caption?: string;
  /** Members */
  members?: MemberOptions[];
  /** Member properties (CT_MemberProperties → mps/mp) */
  memberProperties?: MemberPropertyOptions[];
}

/** Member in pivot hierarchy (CT_Member) */
export interface MemberOptions {
  /** Member name (required) */
  name: string;
  /** Level (CT_Member @level) */
  level?: number;
}

/** Member property (CT_MemberProperty) */
export interface MemberPropertyOptions {
  /** Field index */
  field: number;
  /** Property name */
  name?: string;
  /** Show cell? */
  showCell?: boolean;
  /** Show tip? */
  showTip?: boolean;
  /** Show as caption (CT_MemberProperty @showAsCaption) */
  showAsCaption?: boolean;
  /** Name length (CT_MemberProperty @nameLen) */
  nameLen?: number;
  /** Property position (CT_MemberProperty @pPos) */
  pPos?: number;
  /** Property length (CT_MemberProperty @pLen) */
  pLen?: number;
}

/** Pivot area for conditional formats, chart formats, etc. (CT_PivotArea) */
export interface PivotAreaOptions {
  /** Field index */
  field?: number;
  /** Area type (default: "normal") */
  type?: "none" | "normal" | "data" | "all" | "origin" | "button" | "topEnd" | "topRight";
  /** Data only (default: true) */
  dataOnly?: boolean;
  /** Label only (default: false) */
  labelOnly?: boolean;
  /** Grand row (default: false) */
  grandRow?: boolean;
  /** Grand column (default: false) */
  grandCol?: boolean;
  /** Cache index (default: false) */
  cacheIndex?: boolean;
  /** Outline (default: true) */
  outline?: boolean;
  /** Offset reference */
  offset?: string;
  /** Collapsed levels are subtotals (default: false) */
  collapsedLevelsAreSubtotals?: boolean;
  /** Axis */
  axis?: "axisRow" | "axisCol" | "axisPage" | "axisValues";
  /** Field position */
  fieldPosition?: number;
  /** References */
  references?: PivotAreaReferenceOptions[];
}

/** Pivot area reference (CT_PivotAreaReference) */
export interface PivotAreaReferenceOptions {
  /** Field index */
  field?: number;
  /** Count */
  count?: number;
  /** Selected (default: true) */
  selected?: boolean;
  /** By position (default: false) */
  byPosition?: boolean;
  /** Relative (default: false) */
  relative?: boolean;
  /** Default subtotal */
  defaultSubtotal?: boolean;
  /** Sum subtotal */
  sumSubtotal?: boolean;
  /** CountA subtotal */
  countASubtotal?: boolean;
  /** Average subtotal */
  avgSubtotal?: boolean;
  /** Max subtotal */
  maxSubtotal?: boolean;
  /** Min subtotal */
  minSubtotal?: boolean;
  /** Count subtotal (CT_Reference @countSubtotal) */
  countSubtotal?: boolean;
  /** Product subtotal (CT_Reference @productSubtotal) */
  productSubtotal?: boolean;
  /** StdDevP subtotal (CT_Reference @stdDevPSubtotal) */
  stdDevPSubtotal?: boolean;
  /** StdDev subtotal (CT_Reference @stdDevSubtotal) */
  stdDevSubtotal?: boolean;
  /** VarP subtotal (CT_Reference @varPSubtotal) */
  varPSubtotal?: boolean;
  /** Var subtotal (CT_Reference @varSubtotal) */
  varSubtotal?: boolean;
  /** X indices */
  x?: number[];
}

/** Pivot conditional format (CT_ConditionalFormat for pivot) */
export interface PivotConditionalFormatOptions {
  /** Scope (default: "selection") */
  scope?: "selection" | "data" | "field";
  /** Type (default: "none") */
  type?: "none" | "all" | "row" | "column";
  /** Priority (required) */
  priority: number;
  /** Pivot areas */
  pivotAreas?: PivotAreaOptions[];
}

/** Chart format for pivot table (CT_ChartFormat) */
export interface ChartFormatOptions {
  /** Chart index (required) */
  chart: number;
  /** Format index (required) */
  format: number;
  /** Is series (default: false) */
  series?: boolean;
  /** Pivot area */
  pivotArea?: PivotAreaOptions;
}

/** Cache hierarchy for OLAP pivot caches (CT_CacheHierarchy) */
export interface CacheHierarchyOptions {
  /** Unique name (required) */
  uniqueName: string;
  /** Caption */
  caption?: string;
  /** Is measure (default: false) */
  measure?: boolean;
  /** Is set (default: false) */
  set?: boolean;
  /** Parent set index */
  parentSet?: number;
  /** Icon set (default: 0) */
  iconSet?: number;
  /** Is attribute (default: false) */
  attribute?: boolean;
  /** Is time dimension (default: false) */
  time?: boolean;
  /** Key attribute (default: false) */
  keyAttribute?: boolean;
  /** Default member unique name */
  defaultMemberUniqueName?: string;
  /** All unique name */
  allUniqueName?: string;
  /** All caption */
  allCaption?: string;
  /** Dimension unique name */
  dimensionUniqueName?: string;
  /** Display folder */
  displayFolder?: string;
  /** Measure group */
  measureGroup?: string;
  /** Is measures (default: false) */
  measures?: boolean;
  /** Count (required) */
  count: number;
  /** One field (default: false) */
  oneField?: boolean;
  /** Hidden (default: false) */
  hidden?: boolean;
  /** Member value datatype (CT_CacheHierarchy @memberValueDatatype) */
  memberValueDatatype?: "string" | "number" | "integer" | "boolean" | "error";
  /** Unbalanced (CT_CacheHierarchy @unbalanced) */
  unbalanced?: boolean;
  /** Unbalanced group (CT_CacheHierarchy @unbalancedGroup) */
  unbalancedGroup?: boolean;
  /** Group levels (CT_GroupLevels) */
  groupLevels?: GroupLevelOptions[];
  /** Fields usage (CT_FieldsUsage) */
  fieldsUsage?: FieldUsageOptions[];
}

/** KPI definition for pivot cache (CT_PCDKPI) */
export interface KpiOptions {
  /** Unique name (required) */
  uniqueName: string;
  /** Caption */
  caption?: string;
  /** Display folder */
  displayFolder?: string;
  /** Measure group */
  measureGroup?: string;
  /** Parent */
  parent?: string;
  /** Value expression (required) */
  value: string;
  /** Goal expression */
  goal?: string;
  /** Status expression */
  status?: string;
  /** Trend expression */
  trend?: string;
  /** Weight expression */
  weight?: string;
  /** Time expression */
  time?: string;
}

/** Measure group (CT_MeasureGroup) */
export interface MeasureGroupOptions {
  /** Name (required) */
  name: string;
  /** Caption (required) */
  caption: string;
}

/** Set definition (CT_Set) */
export interface SetOptions {
  /** Count */
  count?: number;
  /** Max rank (required) */
  maxRank: number;
  /** Set definition MDX (required) */
  setDefinition: string;
  /** Sort type (default: "none") */
  sortType?:
    | "none"
    | "ascending"
    | "descending"
    | "ascendingAlpha"
    | "descendingAlpha"
    | "ascendingNatural"
    | "descendingNatural";
  /** Query failed (default: false) */
  queryFailed?: boolean;
}

/** Server format (CT_ServerFormat) */
export interface ServerFormatOptions {
  /** Culture */
  culture?: string;
  /** Format string */
  format?: string;
}

/** Field group for pivot cache (CT_FieldGroup) */
export interface FieldGroupOptions {
  /** Parent field index */
  parent?: number;
  /** Base field index */
  base?: number;
  /** Range properties */
  rangePr?: RangePropertiesOptions;
  /** Discrete properties */
  discretePr?: number[];
  /** Group items names */
  groupItems?: string[];
}

/** Range properties for field grouping (CT_RangePr) */
export interface RangePropertiesOptions {
  /** Auto start (default: true) */
  autoStart?: boolean;
  /** Auto end (default: true) */
  autoEnd?: boolean;
  /** Group by (default: "range") */
  groupBy?: "range" | "seconds" | "minutes" | "hours" | "days" | "months" | "quarters" | "years";
  /** Start number */
  startNum?: number;
  /** End number */
  endNum?: number;
  /** Start date ISO string */
  startDate?: string;
  /** End date ISO string */
  endDate?: string;
  /** Group interval (default: 1) */
  groupInterval?: number;
}

/** Pivot dimension (CT_PivotDimension) */
export interface PivotDimensionOptions {
  /** Is measure (default: false) */
  measure?: boolean;
  /** Name (required) */
  name: string;
  /** Unique name (required) */
  uniqueName: string;
  /** Caption (required) */
  caption: string;
}

/** Range set for consolidation source (CT_RangeSet) */
export interface RangeSetOptions {
  /** Index for page field 1 */
  i1?: number;
  /** Index for page field 2 */
  i2?: number;
  /** Index for page field 3 */
  i3?: number;
  /** Index for page field 4 */
  i4?: number;
  /** Cell reference */
  ref?: string;
  /** Named range */
  name?: string;
  /** Sheet name */
  sheet?: string;
  /** Relationship ID to external workbook */
  rId?: string;
}

/** Page item for consolidation (CT_PageItem) */
export interface ConsolidationPageItemOptions {
  /** Page item name */
  name: string;
}

/** Page for consolidation (CT_PCDSCPage) */
export interface ConsolidationPageOptions {
  /** Page items */
  items?: ConsolidationPageItemOptions[];
}

/** Consolidation source (CT_Consolidation) */
export interface ConsolidationOptions {
  /** Auto page (default: true) */
  autoPage?: boolean;
  /** Pages (max 4) */
  pages?: ConsolidationPageOptions[];
  /** Range sets (required) */
  rangeSets: RangeSetOptions[];
}

/** Tuple cache entry (CT_PCDSDTCEntries choice: m/n/e/s) */
export interface TupleCacheEntryOptions {
  /** Entry type */
  type: "m" | "n" | "e" | "s";
  /** Value (required for n/s, optional for e) */
  value?: number | string;
}

/** Deleted field (CT_DeletedField) */
export interface DeletedFieldOptions {
  /** Field name */
  name: string;
}

/** Group member (CT_GroupMember) */
export interface GroupMemberOptions {
  /** Unique name (required) */
  uniqueName: string;
  /** Is group */
  group?: boolean;
}

/** Level group (CT_LevelGroup) */
export interface LevelGroupOptions {
  /** Name (required) */
  name: string;
  /** Unique name (required) */
  uniqueName: string;
  /** Caption (required) */
  caption: string;
  /** Unique parent */
  uniqueParent?: string;
  /** Group ID */
  id?: number;
  /** Members */
  members: GroupMemberOptions[];
}

/** Group level (CT_GroupLevel) */
export interface GroupLevelOptions {
  /** Unique name (required) */
  uniqueName: string;
  /** Caption (required) */
  caption: string;
  /** User-defined */
  user?: boolean;
  /** Custom roll-up */
  customRollUp?: boolean;
  /** Groups */
  groups?: LevelGroupOptions[];
}

/** Field usage (CT_FieldUsage) */
export interface FieldUsageOptions {
  /** Field index */
  value: number;
}

/** Hierarchy usage (CT_HierarchyUsage) */
export interface HierarchyUsageOptions {
  /** Hierarchy usage value (required) */
  hierarchyUsage: number;
}

/** Query cache entry (CT_Query) */
export interface QueryCacheEntryOptions {
  /** MDX query string (required) */
  mdx: string;
  /** Tuples */
  tpls?: TupleOptions[];
}

/** Tuple (CT_Tuple) */
export interface TupleOptions {
  /** Tuple items (field indices) */
  items?: number[];
}

/** Member property map (CT_X, used as mpMap child) */
export interface MpMapOptions {
  /** Field index */
  x: number;
}

/** Measure dimension map (CT_MeasureDimensionMap) */
export interface MeasureDimensionMapOptions {
  /** Measure group index */
  measureGroup?: number;
  /** Dimension index */
  dimension?: number;
}

/** OLAP properties for pivot cache (CT_OlapPr) */
export interface OLAPPropertiesOptions {
  /** Local cube connection string */
  local?: string;
  /** Local connection string */
  localConnection?: string;
  /** Send locale info to OLAP server */
  sendLocale?: boolean;
  /** Row dimensions */
  rowDrillCount?: number;
  /** Column dimensions */
  colDrillCount?: number;
  /** Local refresh (CT_OlapPr @localRefresh) */
  localRefresh?: boolean;
  /** Use server fill formatting */
  serverFill?: boolean;
  /** Use server number formatting */
  serverNumberFormat?: boolean;
  /** Use server font formatting */
  serverFont?: boolean;
  /** Use server font color */
  serverFontColor?: boolean;
}

/** Parsed source data for pivot cache generation. */
export interface PivotSourceData {
  fieldNames: string[];
  records: (string | number | null | Date)[][];
}

/**
 * Extract unique values from source data for a given field index.
 */
export function collectUniqueValues(
  records: (string | number | null | Date)[][],
  fieldIdx: number,
): (string | number | null | Date)[] {
  const seen = new Set<string>();
  const result: (string | number | null | Date)[] = [];
  for (const row of records) {
    const val = row[fieldIdx];
    const key = val instanceof Date ? val.toISOString() : String(val);
    if (!seen.has(key)) {
      seen.add(key);
      result.push(val);
    }
  }
  return result;
}

/**
 * Check if a field is numeric (all non-empty values are numbers).
 */
export function isNumericField(
  records: (string | number | null | Date)[][],
  fieldIdx: number,
): boolean {
  for (const row of records) {
    const val = row[fieldIdx];
    if (typeof val === "string" && val !== "") return false;
  }
  return true;
}

/**
 * Aggregate values using the specified function.
 */
export function aggregate(values: number[], func: ConsolidateFunction): number {
  if (values.length === 0) return 0;
  switch (func) {
    case "sum":
      return values.reduce((a, b) => a + b, 0);
    case "count":
    case "countNums":
      return values.length;
    case "average":
      return values.reduce((a, b) => a + b, 0) / values.length;
    case "max":
      // Two-arg reduce: Math.max is associative, so Math.max(a,b,c) ===
      // Math.max(Math.max(a,b),c) for all values incl. NaN/Infinity — fully
      // equivalent to Math.max(...values) without spreading onto the stack.
      return values.reduce((a, b) => Math.max(a, b));
    case "min":
      return values.reduce((a, b) => Math.min(a, b));
    case "product":
      return values.reduce((a, b) => a * b, 1);
    case "var":
      return sampleVariance(values);
    case "varp":
      return populationVariance(values);
    case "stdDev":
      return Math.sqrt(sampleVariance(values));
    case "stdDevp":
      return Math.sqrt(populationVariance(values));
    default:
      return values.reduce((a, b) => a + b, 0);
  }
}

function populationVariance(values: number[]): number {
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  return values.reduce((sum, v) => sum + (v - mean) ** 2, 0) / values.length;
}

function sampleVariance(values: number[]): number {
  if (values.length < 2) return 0;
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  return values.reduce((sum, v) => sum + (v - mean) ** 2, 0) / (values.length - 1);
}

// ── PivotCacheDefinition types (extracted from pivot-cache-definition-xml) ──

export interface CacheFieldExtraAttrs {
  /** Database field (CT_CacheField @databaseField) */
  databaseField?: boolean;
  /** Level (CT_CacheField @level) */
  level?: number;
  /** Mapping count (CT_CacheField @mappingCount) */
  mappingCount?: number;
  /** Member property field (CT_CacheField @memberPropertyField) */
  memberPropertyField?: number;
  /** Property name (CT_CacheField @propertyName) */
  propertyName?: string;
  /** Server field (CT_CacheField @serverField) */
  serverField?: boolean;
  /** Unique list (CT_CacheField @uniqueList) */
  uniqueList?: boolean;
  /** Shared items contains mixed types (CT_SharedItems @containsMixedTypes) */
  containsMixedTypes?: boolean;
  /** Shared items contains non-date (CT_SharedItems @containsNonDate) */
  containsNonDate?: boolean;
  /** Shared items long text (CT_SharedItems @longText) */
  longText?: boolean;
  /** Shared items max date (CT_SharedItems @maxDate) */
  maxDate?: string;
  /** Shared items min date (CT_SharedItems @minDate) */
  minDate?: string;
}

export interface PivotCacheDefinitionOptions {
  /** Cache is invalid (CT_PivotCacheDefinition @invalid) */
  invalid?: boolean;
  /** Save data with cache (CT_PivotCacheDefinition @saveData) */
  saveData?: boolean;
  /** Optimize memory usage (CT_PivotCacheDefinition @optimizeMemory) */
  optimizeMemory?: boolean;
  /** Enable refresh (CT_PivotCacheDefinition @enableRefresh) */
  enableRefresh?: boolean;
  /** User who last refreshed */
  refreshedBy?: string;
  /** Date of last refresh (decimal) */
  refreshedDate?: number;
  /** Date of last refresh (ISO 8601) */
  refreshedDateIso?: string;
  /** Background query (CT_PivotCacheDefinition @backgroundQuery) */
  backgroundQuery?: boolean;
  /** Missing items limit */
  missingItemsLimit?: number;
  /** Upgrade on refresh */
  upgradeOnRefresh?: boolean;
  /** Support subquery */
  supportSubquery?: boolean;
  /** Support advanced drill */
  supportAdvancedDrill?: boolean;
  /** Cache hierarchies (CT_CacheHierarchies) */
  cacheHierarchies?: CacheHierarchyOptions[];
  /** KPIs (CT_PCDKPIs) */
  kpis?: KpiOptions[];
  /** Measure groups (CT_MeasureGroups) */
  measureGroups?: MeasureGroupOptions[];
  /** Dimensions (CT_Dimensions) */
  dimensions?: PivotDimensionOptions[];
  /** Sets (CT_Sets in tupleCache) */
  sets?: SetOptions[];
  /** Server formats (CT_ServerFormats) */
  serverFormats?: ServerFormatOptions[];
  /** Field groups per field index (CT_FieldGroup inside cacheField) */
  fieldGroups?: ReadonlyMap<number, FieldGroupOptions>;
  /** Consolidation source (alternative to worksheetSource) */
  consolidation?: ConsolidationOptions;
  /** Tuple cache entries (CT_PCDSDTCEntries) */
  entries?: TupleCacheEntryOptions[];
  /** Query cache (CT_QueryCache in tupleCache) */
  queryCache?: QueryCacheEntryOptions[];
  /** Member property map per cache field (mpMap) */
  mpMaps?: MpMapOptions[];
  /** Measure dimension maps (CT_MeasureDimensionMaps) */
  measureDimensionMaps?: MeasureDimensionMapOptions[];
  /** Per-field cache field overrides (mapped by field index) */
  cacheFieldOverrides?: ReadonlyMap<number, CacheFieldExtraAttrs>;
  /** OLAP properties (CT_OlapPr) */
  olapPr?: OLAPPropertiesOptions;
}
