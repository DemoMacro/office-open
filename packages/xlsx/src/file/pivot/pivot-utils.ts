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
  readonly field: string;
  /** Aggregation function (default: "sum") */
  readonly summarize?: ConsolidateFunction;
  /** Custom name for the data field (default: "Sum of {field}") */
  readonly name?: string;
  /** Show data as (CT_DataField @showDataAs) */
  readonly showDataAs?: string;
  /** Base field index for "show data as" calculations */
  readonly baseField?: number;
  /** Base item index for "show data as" calculations */
  readonly baseItem?: number;
  /** Sort by tuple items (CT_Tuples in sortByTuple) */
  readonly sortByTupleItems?: readonly number[];
}

/** Per-field overrides for pivotField XML attributes (CT_PivotField). */
export interface PivotFieldOverrideOptions {
  /** Field name to match (required) */
  readonly field: string;
  /** All drilled (CT_PivotField @allDrilled) */
  readonly allDrilled?: boolean;
  /** Auto show (CT_PivotField @autoShow) */
  readonly autoShow?: boolean;
  /** Count subtotal (CT_PivotField @countSubtotal) */
  readonly countSubtotal?: boolean;
  /** Data source sort (CT_PivotField @dataSourceSort) */
  readonly dataSourceSort?: boolean;
  /** Default attribute drill state (CT_PivotField @defaultAttributeDrillState) */
  readonly defaultAttributeDrillState?: boolean;
  /** Hidden level (CT_PivotField @hiddenLevel) */
  readonly hiddenLevel?: boolean;
  /** Hide new items (CT_PivotField @hideNewItems) */
  readonly hideNewItems?: boolean;
  /** Insert blank row (CT_PivotField @insertBlankRow) */
  readonly insertBlankRow?: boolean;
  /** Insert page break (CT_PivotField @insertPageBreak) */
  readonly insertPageBreak?: boolean;
  /** Item page count (CT_PivotField @itemPageCount) */
  readonly itemPageCount?: boolean;
  /** Measure filter (CT_PivotField @measureFilter) */
  readonly measureFilter?: boolean;
  /** Non auto sort default (CT_PivotField @nonAutoSortDefault) */
  readonly nonAutoSortDefault?: boolean;
  /** Product subtotal (CT_PivotField @productSubtotal) */
  readonly productSubtotal?: boolean;
  /** Rank by (CT_PivotField @rankBy) */
  readonly rankBy?: number;
  /** Server field (CT_PivotField @serverField) */
  readonly serverField?: boolean;
  /** Show drop downs (CT_PivotField @showDropDowns) */
  readonly showDropDowns?: boolean;
  /** Show property as caption (CT_PivotField @showPropAsCaption) */
  readonly showPropAsCaption?: boolean;
  /** Show property cell (CT_PivotField @showPropCell) */
  readonly showPropCell?: boolean;
  /** Show property tip (CT_PivotField @showPropTip) */
  readonly showPropTip?: boolean;
  /** StdDevP subtotal (CT_PivotField @stdDevPSubtotal) */
  readonly stdDevPSubtotal?: boolean;
  /** StdDev subtotal (CT_PivotField @stdDevSubtotal) */
  readonly stdDevSubtotal?: boolean;
  /** Subtotal caption (CT_PivotField @subtotalCaption) */
  readonly subtotalCaption?: string;
  /** Top auto show (CT_PivotField @topAutoShow) */
  readonly topAutoShow?: boolean;
  /** Unique member property (CT_PivotField @uniqueMemberProperty) */
  readonly uniqueMemberProperty?: boolean;
  /** VarP subtotal (CT_PivotField @varPSubtotal) */
  readonly varPSubtotal?: boolean;
  /** Var subtotal (CT_PivotField @varSubtotal) */
  readonly varSubtotal?: boolean;
  /** Show detail for default item (CT_Item @sd) */
  readonly defaultItemSd?: boolean;
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
  readonly fld: number;
  /** Filter type (required) */
  readonly type: PivotFilterType;
  /** Filter ID — unique within this pivot table (required) */
  readonly id: number;
  /** Measure field index for OLAP filters */
  readonly mpFld?: number;
  /** Evaluation order */
  readonly evalOrder?: number;
  /** Measure hierarchy */
  readonly iMeasureHier?: number;
  /** Measure field */
  readonly iMeasureFld?: number;
  /** Filter name */
  readonly name?: string;
  /** Filter description */
  readonly description?: string;
  /** First string value for caption/date filters */
  readonly stringValue1?: string;
  /** Second string value for between filters */
  readonly stringValue2?: string;
}

/** Options for a single pivot table on a worksheet. */
export interface PivotTableOptions {
  /** Pivot table name (default: "PivotTable{N}") */
  readonly name?: string;
  /** Source data range, e.g. "A1:D11" — must be on the same or a different sheet */
  readonly source: string;
  /** Source sheet name (default: current sheet) */
  readonly sourceSheet?: string;
  /** Target cell for the pivot table output (default: "A3") */
  readonly location?: string;
  /** Field names to use as row labels */
  readonly rows: readonly string[];
  /** Field names to use as column labels */
  readonly columns?: readonly string[];
  /** Data fields with aggregation settings */
  readonly data: readonly PivotDataField[];
  /** Pivot style name (default: "PivotStyleLight16") */
  readonly style?: string;
  /** Pivot filters (CT_PivotFilters) */
  readonly filters?: readonly PivotFilterOptions[];
  /** Field names to use as page/report filters */
  readonly pages?: readonly string[];
  /** Page field captions (maps by index to pages array) */
  readonly pageCaptions?: readonly string[];
  /** Data fields on rows instead of columns (CT_PivotTableDefinition @dataOnRows) */
  readonly dataOnRows?: boolean;
  /** Grand total caption text */
  readonly grandTotalCaption?: string;
  /** Error caption text */
  readonly errorCaption?: string;
  /** Show error messages */
  readonly showError?: boolean;
  /** Missing caption text */
  readonly missingCaption?: string;
  /** Show missing items */
  readonly showMissing?: boolean;
  /** Custom page style name */
  readonly pageStyle?: string;
  /** Custom pivot table style name */
  readonly pivotTableStyle?: string;
  /** Tag string */
  readonly tag?: string;
  /** Show items with no data */
  readonly showItems?: boolean;
  /** Edit data in-place */
  readonly editData?: boolean;
  /** Disable field list */
  readonly disableFieldList?: boolean;
  /** Show calculated members */
  readonly showCalcMbrs?: boolean;
  /** Visual totals */
  readonly visualTotals?: boolean;
  /** Show multiple labels */
  readonly showMultipleLabel?: boolean;
  /** Show data drop-down */
  readonly showDataDropDown?: boolean;
  /** Show drill indicators */
  readonly showDrill?: boolean;
  /** Print drill indicators */
  readonly printDrill?: boolean;
  /** Show member property tips */
  readonly showMemberPropertyTips?: boolean;
  /** Show data tips */
  readonly showDataTips?: boolean;
  /** Enable layout wizard */
  readonly enableWizard?: boolean;
  /** Enable drill-down */
  readonly enableDrill?: boolean;
  /** Enable field properties */
  readonly enableFieldProperties?: boolean;
  /** Number of page fields per row/column */
  readonly pageWrap?: number;
  /** Page layout over then down */
  readonly pageOverThenDown?: boolean;
  /** Subtotal hidden items */
  readonly subtotalHiddenItems?: boolean;
  /** Field print titles */
  readonly fieldPrintTitles?: boolean;
  /** Merge item labels */
  readonly mergeItem?: boolean;
  /** Show drop zones */
  readonly showDropZones?: boolean;
  /** Show empty row */
  readonly showEmptyRow?: boolean;
  /** Show empty column */
  readonly showEmptyCol?: boolean;
  /** Show headers */
  readonly showHeaders?: boolean;
  /** Published to server */
  readonly published?: boolean;
  /** Grid drop zones */
  readonly gridDropZones?: boolean;
  /** Multiple field filters */
  readonly multipleFieldFilters?: boolean;
  /** Row header caption */
  readonly rowHeaderCaption?: string;
  /** Column header caption */
  readonly colHeaderCaption?: string;
  /** Sort field list ascending */
  readonly fieldListSortAscending?: boolean;
  /** MDX subqueries enabled */
  readonly mdxSubqueries?: boolean;
  /** Custom list sort */
  readonly customListSort?: boolean;
  /** Asterisk totals (CT_PivotTableDefinition @asteriskTotals) */
  readonly asteriskTotals?: boolean;
  /** Data position (CT_PivotTableDefinition @dataPosition) */
  readonly dataPosition?: number;
  /** Immersive (CT_PivotTableDefinition @immersive) */
  readonly immersive?: boolean;
  /** Vacated style (CT_PivotTableDefinition @vacatedStyle) */
  readonly vacatedStyle?: string;
  /** Calculated items (CT_CalculatedItems) */
  readonly calculatedItems?: readonly CalculatedItemOptions[];
  /** Calculated members (CT_CalculatedMembers) */
  readonly calculatedMembers?: readonly CalculatedMemberOptions[];
  /** Pivot hierarchies (CT_PivotHierarchies) */
  readonly pivotHierarchies?: readonly PivotHierarchyOptions[];
  /** Conditional formats (CT_ConditionalFormats for pivot) */
  readonly pivotConditionalFormats?: readonly PivotConditionalFormatOptions[];
  /** Chart formats (CT_ChartFormats) */
  readonly chartFormats?: readonly ChartFormatOptions[];
  /** Auto sort scope (CT_AutoSortScope) */
  readonly autoSortScope?: PivotAreaOptions;
  /** Member properties per field (CT_MemberProperties → mps/mp) */
  readonly memberProperties?: readonly MemberPropertyOptions[];
  /** Pivot format areas (CT_Formats → format) */
  readonly formats?: readonly PivotFormatOptions[];
  /** Row hierarchy usage (CT_RowHierarchiesUsage) */
  readonly rowHierarchiesUsage?: readonly HierarchyUsageOptions[];
  /** Column hierarchy usage (CT_ColHierarchiesUsage) */
  readonly colHierarchiesUsage?: readonly HierarchyUsageOptions[];
  /** Location column page count (CT_Location @colPageCount) */
  readonly locationColPageCount?: number;
  /** Location row page count (CT_Location @rowPageCount) */
  readonly locationRowPageCount?: number;
  /** Per-field overrides for pivotField (CT_PivotField attributes) */
  readonly fieldOverrides?: readonly PivotFieldOverrideOptions[];
}

/** Pivot format (CT_Format). */
export interface PivotFormatOptions {
  /** Action type (default: "formatting") */
  readonly action?: "formatting" | "drill" | "formula" | "blank" | "subtotal" | "report";
  /** Differential format index */
  readonly dxfId?: number;
  /** Pivot area */
  readonly pivotArea: PivotAreaOptions;
}

/** Calculated item in a pivot table (CT_CalculatedItem) */
export interface CalculatedItemOptions {
  /** Field index */
  readonly field?: number;
  /** Formula */
  readonly formula?: string;
  /** Pivot area reference */
  readonly pivotArea?: PivotAreaOptions;
}

/** Calculated member in a pivot table (CT_CalculatedMember) */
export interface CalculatedMemberOptions {
  /** Name (required) */
  readonly name: string;
  /** MDX expression (required) */
  readonly mdx: string;
  /** Member name */
  readonly memberName?: string;
  /** Hierarchy */
  readonly hierarchy?: string;
  /** Parent member */
  readonly parent?: string;
  /** Solve order (default: 0) */
  readonly solveOrder?: number;
  /** Is a set (default: false) */
  readonly set?: boolean;
}

/** Pivot hierarchy (CT_PivotHierarchy) */
export interface PivotHierarchyOptions {
  /** Outline mode (default: false) */
  readonly outline?: boolean;
  /** Allow multiple item selection (default: false) */
  readonly multipleItemSelectionAllowed?: boolean;
  /** Subtotal on top (default: false) */
  readonly subtotalTop?: boolean;
  /** Show in field list (default: true) */
  readonly showInFieldList?: boolean;
  /** Drag to row (default: true) */
  readonly dragToRow?: boolean;
  /** Drag to column (default: true) */
  readonly dragToCol?: boolean;
  /** Drag to page (default: true) */
  readonly dragToPage?: boolean;
  /** Drag to data (default: false) */
  readonly dragToData?: boolean;
  /** Drag off (default: true) */
  readonly dragOff?: boolean;
  /** Include new items in filter (default: false) */
  readonly includeNewItemsInFilter?: boolean;
  /** Caption */
  readonly caption?: string;
  /** Members */
  readonly members?: readonly MemberOptions[];
  /** Member properties (CT_MemberProperties → mps/mp) */
  readonly memberProperties?: readonly MemberPropertyOptions[];
}

/** Member in pivot hierarchy (CT_Member) */
export interface MemberOptions {
  /** Member name (required) */
  readonly name: string;
  /** Level (CT_Member @level) */
  readonly level?: number;
}

/** Member property (CT_MemberProperty) */
export interface MemberPropertyOptions {
  /** Field index */
  readonly field: number;
  /** Property name */
  readonly name?: string;
  /** Show cell? */
  readonly showCell?: boolean;
  /** Show tip? */
  readonly showTip?: boolean;
  /** Show as caption (CT_MemberProperty @showAsCaption) */
  readonly showAsCaption?: boolean;
  /** Name length (CT_MemberProperty @nameLen) */
  readonly nameLen?: number;
  /** Property position (CT_MemberProperty @pPos) */
  readonly pPos?: number;
  /** Property length (CT_MemberProperty @pLen) */
  readonly pLen?: number;
}

/** Pivot area for conditional formats, chart formats, etc. (CT_PivotArea) */
export interface PivotAreaOptions {
  /** Field index */
  readonly field?: number;
  /** Area type (default: "normal") */
  readonly type?: "none" | "normal" | "data" | "all" | "origin" | "button" | "topEnd" | "topRight";
  /** Data only (default: true) */
  readonly dataOnly?: boolean;
  /** Label only (default: false) */
  readonly labelOnly?: boolean;
  /** Grand row (default: false) */
  readonly grandRow?: boolean;
  /** Grand column (default: false) */
  readonly grandCol?: boolean;
  /** Cache index (default: false) */
  readonly cacheIndex?: boolean;
  /** Outline (default: true) */
  readonly outline?: boolean;
  /** Offset reference */
  readonly offset?: string;
  /** Collapsed levels are subtotals (default: false) */
  readonly collapsedLevelsAreSubtotals?: boolean;
  /** Axis */
  readonly axis?: "axisRow" | "axisCol" | "axisPage" | "axisValues";
  /** Field position */
  readonly fieldPosition?: number;
  /** References */
  readonly references?: readonly PivotAreaReferenceOptions[];
}

/** Pivot area reference (CT_PivotAreaReference) */
export interface PivotAreaReferenceOptions {
  /** Field index */
  readonly field?: number;
  /** Count */
  readonly count?: number;
  /** Selected (default: true) */
  readonly selected?: boolean;
  /** By position (default: false) */
  readonly byPosition?: boolean;
  /** Relative (default: false) */
  readonly relative?: boolean;
  /** Default subtotal */
  readonly defaultSubtotal?: boolean;
  /** Sum subtotal */
  readonly sumSubtotal?: boolean;
  /** CountA subtotal */
  readonly countASubtotal?: boolean;
  /** Average subtotal */
  readonly avgSubtotal?: boolean;
  /** Max subtotal */
  readonly maxSubtotal?: boolean;
  /** Min subtotal */
  readonly minSubtotal?: boolean;
  /** Count subtotal (CT_Reference @countSubtotal) */
  readonly countSubtotal?: boolean;
  /** Product subtotal (CT_Reference @productSubtotal) */
  readonly productSubtotal?: boolean;
  /** StdDevP subtotal (CT_Reference @stdDevPSubtotal) */
  readonly stdDevPSubtotal?: boolean;
  /** StdDev subtotal (CT_Reference @stdDevSubtotal) */
  readonly stdDevSubtotal?: boolean;
  /** VarP subtotal (CT_Reference @varPSubtotal) */
  readonly varPSubtotal?: boolean;
  /** Var subtotal (CT_Reference @varSubtotal) */
  readonly varSubtotal?: boolean;
  /** X indices */
  readonly x?: readonly number[];
}

/** Pivot conditional format (CT_ConditionalFormat for pivot) */
export interface PivotConditionalFormatOptions {
  /** Scope (default: "selection") */
  readonly scope?: "selection" | "data" | "field";
  /** Type (default: "none") */
  readonly type?: "none" | "all" | "row" | "column";
  /** Priority (required) */
  readonly priority: number;
  /** Pivot areas */
  readonly pivotAreas?: readonly PivotAreaOptions[];
}

/** Chart format for pivot table (CT_ChartFormat) */
export interface ChartFormatOptions {
  /** Chart index (required) */
  readonly chart: number;
  /** Format index (required) */
  readonly format: number;
  /** Is series (default: false) */
  readonly series?: boolean;
  /** Pivot area */
  readonly pivotArea?: PivotAreaOptions;
}

/** Cache hierarchy for OLAP pivot caches (CT_CacheHierarchy) */
export interface CacheHierarchyOptions {
  /** Unique name (required) */
  readonly uniqueName: string;
  /** Caption */
  readonly caption?: string;
  /** Is measure (default: false) */
  readonly measure?: boolean;
  /** Is set (default: false) */
  readonly set?: boolean;
  /** Parent set index */
  readonly parentSet?: number;
  /** Icon set (default: 0) */
  readonly iconSet?: number;
  /** Is attribute (default: false) */
  readonly attribute?: boolean;
  /** Is time dimension (default: false) */
  readonly time?: boolean;
  /** Key attribute (default: false) */
  readonly keyAttribute?: boolean;
  /** Default member unique name */
  readonly defaultMemberUniqueName?: string;
  /** All unique name */
  readonly allUniqueName?: string;
  /** All caption */
  readonly allCaption?: string;
  /** Dimension unique name */
  readonly dimensionUniqueName?: string;
  /** Display folder */
  readonly displayFolder?: string;
  /** Measure group */
  readonly measureGroup?: string;
  /** Is measures (default: false) */
  readonly measures?: boolean;
  /** Count (required) */
  readonly count: number;
  /** One field (default: false) */
  readonly oneField?: boolean;
  /** Hidden (default: false) */
  readonly hidden?: boolean;
  /** Member value datatype (CT_CacheHierarchy @memberValueDatatype) */
  readonly memberValueDatatype?: "string" | "number" | "integer" | "boolean" | "error";
  /** Unbalanced (CT_CacheHierarchy @unbalanced) */
  readonly unbalanced?: boolean;
  /** Unbalanced group (CT_CacheHierarchy @unbalancedGroup) */
  readonly unbalancedGroup?: boolean;
  /** Group levels (CT_GroupLevels) */
  readonly groupLevels?: readonly GroupLevelOptions[];
  /** Fields usage (CT_FieldsUsage) */
  readonly fieldsUsage?: readonly FieldUsageOptions[];
}

/** KPI definition for pivot cache (CT_PCDKPI) */
export interface KpiOptions {
  /** Unique name (required) */
  readonly uniqueName: string;
  /** Caption */
  readonly caption?: string;
  /** Display folder */
  readonly displayFolder?: string;
  /** Measure group */
  readonly measureGroup?: string;
  /** Parent */
  readonly parent?: string;
  /** Value expression (required) */
  readonly value: string;
  /** Goal expression */
  readonly goal?: string;
  /** Status expression */
  readonly status?: string;
  /** Trend expression */
  readonly trend?: string;
  /** Weight expression */
  readonly weight?: string;
  /** Time expression */
  readonly time?: string;
}

/** Measure group (CT_MeasureGroup) */
export interface MeasureGroupOptions {
  /** Name (required) */
  readonly name: string;
  /** Caption (required) */
  readonly caption: string;
}

/** Set definition (CT_Set) */
export interface SetOptions {
  /** Count */
  readonly count?: number;
  /** Max rank (required) */
  readonly maxRank: number;
  /** Set definition MDX (required) */
  readonly setDefinition: string;
  /** Sort type (default: "none") */
  readonly sortType?:
    | "none"
    | "ascending"
    | "descending"
    | "ascendingAlpha"
    | "descendingAlpha"
    | "ascendingNatural"
    | "descendingNatural";
  /** Query failed (default: false) */
  readonly queryFailed?: boolean;
}

/** Server format (CT_ServerFormat) */
export interface ServerFormatOptions {
  /** Culture */
  readonly culture?: string;
  /** Format string */
  readonly format?: string;
}

/** Field group for pivot cache (CT_FieldGroup) */
export interface FieldGroupOptions {
  /** Parent field index */
  readonly parent?: number;
  /** Base field index */
  readonly base?: number;
  /** Range properties */
  readonly rangePr?: RangePrOptions;
  /** Discrete properties */
  readonly discretePr?: readonly number[];
  /** Group items names */
  readonly groupItems?: readonly string[];
}

/** Range properties for field grouping (CT_RangePr) */
export interface RangePrOptions {
  /** Auto start (default: true) */
  readonly autoStart?: boolean;
  /** Auto end (default: true) */
  readonly autoEnd?: boolean;
  /** Group by (default: "range") */
  readonly groupBy?:
    | "range"
    | "seconds"
    | "minutes"
    | "hours"
    | "days"
    | "months"
    | "quarters"
    | "years";
  /** Start number */
  readonly startNum?: number;
  /** End number */
  readonly endNum?: number;
  /** Start date ISO string */
  readonly startDate?: string;
  /** End date ISO string */
  readonly endDate?: string;
  /** Group interval (default: 1) */
  readonly groupInterval?: number;
}

/** Pivot dimension (CT_PivotDimension) */
export interface PivotDimensionOptions {
  /** Is measure (default: false) */
  readonly measure?: boolean;
  /** Name (required) */
  readonly name: string;
  /** Unique name (required) */
  readonly uniqueName: string;
  /** Caption (required) */
  readonly caption: string;
}

/** Range set for consolidation source (CT_RangeSet) */
export interface RangeSetOptions {
  /** Index for page field 1 */
  readonly i1?: number;
  /** Index for page field 2 */
  readonly i2?: number;
  /** Index for page field 3 */
  readonly i3?: number;
  /** Index for page field 4 */
  readonly i4?: number;
  /** Cell reference */
  readonly ref?: string;
  /** Named range */
  readonly name?: string;
  /** Sheet name */
  readonly sheet?: string;
  /** Relationship ID to external workbook */
  readonly rId?: string;
}

/** Page item for consolidation (CT_PageItem) */
export interface ConsolidationPageItemOptions {
  /** Page item name */
  readonly name: string;
}

/** Page for consolidation (CT_PCDSCPage) */
export interface ConsolidationPageOptions {
  /** Page items */
  readonly items?: readonly ConsolidationPageItemOptions[];
}

/** Consolidation source (CT_Consolidation) */
export interface ConsolidationOptions {
  /** Auto page (default: true) */
  readonly autoPage?: boolean;
  /** Pages (max 4) */
  readonly pages?: readonly ConsolidationPageOptions[];
  /** Range sets (required) */
  readonly rangeSets: readonly RangeSetOptions[];
}

/** Tuple cache entry (CT_PCDSDTCEntries choice: m/n/e/s) */
export interface TupleCacheEntryOptions {
  /** Entry type */
  readonly type: "m" | "n" | "e" | "s";
  /** Value (required for n/s, optional for e) */
  readonly value?: number | string;
}

/** Deleted field (CT_DeletedField) */
export interface DeletedFieldOptions {
  /** Field name */
  readonly name: string;
}

/** Group member (CT_GroupMember) */
export interface GroupMemberOptions {
  /** Unique name (required) */
  readonly uniqueName: string;
  /** Is group */
  readonly group?: boolean;
}

/** Level group (CT_LevelGroup) */
export interface LevelGroupOptions {
  /** Name (required) */
  readonly name: string;
  /** Unique name (required) */
  readonly uniqueName: string;
  /** Caption (required) */
  readonly caption: string;
  /** Unique parent */
  readonly uniqueParent?: string;
  /** Group ID */
  readonly id?: number;
  /** Members */
  readonly members: readonly GroupMemberOptions[];
}

/** Group level (CT_GroupLevel) */
export interface GroupLevelOptions {
  /** Unique name (required) */
  readonly uniqueName: string;
  /** Caption (required) */
  readonly caption: string;
  /** User-defined */
  readonly user?: boolean;
  /** Custom roll-up */
  readonly customRollUp?: boolean;
  /** Groups */
  readonly groups?: readonly LevelGroupOptions[];
}

/** Field usage (CT_FieldUsage) */
export interface FieldUsageOptions {
  /** Field index */
  readonly value: number;
}

/** Hierarchy usage (CT_HierarchyUsage) */
export interface HierarchyUsageOptions {
  /** Hierarchy usage value (required) */
  readonly hierarchyUsage: number;
}

/** Query cache entry (CT_Query) */
export interface QueryCacheEntryOptions {
  /** MDX query string (required) */
  readonly mdx: string;
  /** Tuples */
  readonly tpls?: readonly TupleOptions[];
}

/** Tuple (CT_Tuple) */
export interface TupleOptions {
  /** Tuple items (field indices) */
  readonly items?: readonly number[];
}

/** Member property map (CT_X, used as mpMap child) */
export interface MpMapOptions {
  /** Field index */
  readonly x: number;
}

/** Measure dimension map (CT_MeasureDimensionMap) */
export interface MeasureDimensionMapOptions {
  /** Measure group index */
  readonly measureGroup?: number;
  /** Dimension index */
  readonly dimension?: number;
}

/** OLAP properties for pivot cache (CT_OlapPr) */
export interface OlapPrOptions {
  /** Local cube connection string */
  readonly local?: string;
  /** Local connection string */
  readonly localConnection?: string;
  /** Send locale info to OLAP server */
  readonly sendLocale?: boolean;
  /** Row dimensions */
  readonly rowDrillCount?: number;
  /** Column dimensions */
  readonly colDrillCount?: number;
  /** Local refresh (CT_OlapPr @localRefresh) */
  readonly localRefresh?: boolean;
  /** Use server fill formatting */
  readonly serverFill?: boolean;
  /** Use server number formatting */
  readonly serverNumberFormat?: boolean;
  /** Use server font formatting */
  readonly serverFont?: boolean;
  /** Use server font color */
  readonly serverFontColor?: boolean;
}

/** Parsed source data for pivot cache generation. */
export interface PivotSourceData {
  readonly fieldNames: readonly string[];
  readonly records: readonly (string | number | null | Date)[][];
}

/**
 * Extract unique values from source data for a given field index.
 */
export function collectUniqueValues(
  records: readonly (string | number | null | Date)[][],
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
  records: readonly (string | number | null | Date)[][],
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
export function aggregate(values: readonly number[], func: ConsolidateFunction): number {
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
      return Math.max(...values);
    case "min":
      return Math.min(...values);
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

function populationVariance(values: readonly number[]): number {
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  return values.reduce((sum, v) => sum + (v - mean) ** 2, 0) / values.length;
}

function sampleVariance(values: readonly number[]): number {
  if (values.length < 2) return 0;
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  return values.reduce((sum, v) => sum + (v - mean) ** 2, 0) / (values.length - 1);
}
