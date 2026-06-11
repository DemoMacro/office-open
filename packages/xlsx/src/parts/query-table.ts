/**
 * QueryTable options for SpreadsheetML documents.
 *
 * Reference: OOXML transitional, sml.xsd, CT_QueryTable
 *
 * @module
 */

export interface QueryTableDeletedFieldOptions {
  /** Field name that was deleted (required) */
  name: string;
}

export interface QueryTableFieldOptions {
  /** Field ID (required, 1-based) */
  id: number;
  /** Field name */
  name?: string;
  /** Table column index (1-based) */
  tableColumnId?: number;
  /** Row number (for data layout) */
  row?: number;
  /** Fill formatting */
  fillFormatting?: boolean;
  /** Text formatting */
  textFormatting?: boolean;
  /** Number formatting */
  numberFormatting?: boolean;
  /** Border formatting */
  borderFormatting?: boolean;
  /** Width */
  width?: number;
  /** CLipped */
  clipped?: boolean;
  /** Data bound (CT_QueryTableField @dataBound) */
  dataBound?: boolean;
}

export interface QueryTableRefreshOptions {
  /** Next unique ID for new rows */
  nextId?: number;
  /** Minimum refresh version */
  minimumVersion?: number;
  /** Preserve column sort/filter/layout on refresh */
  preserveFormatting?: boolean;
  /** Adjust column width on refresh */
  adjustColumnWidth?: boolean;
  /** Refresh data on load */
  refreshOnLoad?: boolean;
  /** Background refresh */
  backgroundRefresh?: boolean;
  /** Deleted fields */
  deletedFields?: QueryTableDeletedFieldOptions[];
  /** Query table fields */
  queryTableFields?: QueryTableFieldOptions[];
  /** Row count */
  rowCount?: number;
  /** Field ID wrapped (CT_QueryTableRefresh @fieldIdWrapped) */
  fieldIdWrapped?: boolean;
  /** Headers in last refresh (CT_QueryTableRefresh @headersInLastRefresh) */
  headersInLastRefresh?: boolean;
  /** Preserve sort/filter layout (CT_QueryTableRefresh @preserveSortFilterLayout) */
  preserveSortFilterLayout?: boolean;
  /** Unbound columns left (CT_QueryTableRefresh @unboundColumnsLeft) */
  unboundColumnsLeft?: number;
  /** Unbound columns right (CT_QueryTableRefresh @unboundColumnsRight) */
  unboundColumnsRight?: number;
}

export interface QueryTableOptions {
  /** Query table name */
  name?: string;
  /** Connection ID referencing the workbook connection */
  connectionId: number;
  /** Auto-refresh on open */
  autoFormat?: boolean;
  /** Preserve column sort/filter/layout on refresh */
  preserveFormatting?: boolean;
  /** Adjust column width on refresh */
  adjustColumnWidth?: boolean;
  /** Refresh data on load */
  refreshOnLoad?: boolean;
  /** Background refresh */
  backgroundRefresh?: boolean;
  /** Show row numbers (CT_QueryTable @rowNumbers) */
  rowNumbers?: boolean;
  /** Disable refresh (CT_QueryTable @disableRefresh) */
  disableRefresh?: boolean;
  /** First background refresh (CT_QueryTable @firstBackgroundRefresh) */
  firstBackgroundRefresh?: boolean;
  /** Grow/shrink type (CT_QueryTable @growShrinkType) */
  growShrinkType?: boolean;
  /** Fill formulas on refresh (CT_QueryTable @fillFormulas) */
  fillFormulas?: boolean;
  /** Remove data on save (CT_QueryTable @removeDataOnSave) */
  removeDataOnSave?: boolean;
  /** Disable edit (CT_QueryTable @disableEdit) */
  disableEdit?: boolean;
  /** Intermediate (CT_QueryTable @intermediate) */
  intermediate?: boolean;
  /** Query table refresh info (CT_QueryTableRefresh) */
  queryTableRefresh?: QueryTableRefreshOptions;
}
