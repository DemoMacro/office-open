/**
 * QueryTable XML generator — produces xl/queryTables/queryTable{n}.xml.
 *
 * A query table represents data retrieved from an external data source.
 *
 * Reference: OOXML transitional, sml.xsd, CT_QueryTable
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs, escapeXml } from "@office-open/xml";

// ── Options ──

export interface QueryTableDeletedFieldOptions {
  /** Field name that was deleted (required) */
  readonly name: string;
}

export interface QueryTableFieldOptions {
  /** Field ID (required, 1-based) */
  readonly id: number;
  /** Field name */
  readonly name?: string;
  /** Table column index (1-based) */
  readonly tableColumnId?: number;
  /** Row number (for data layout) */
  readonly row?: number;
  /** Fill formatting */
  readonly fillFormatting?: boolean;
  /** Text formatting */
  readonly textFormatting?: boolean;
  /** Number formatting */
  readonly numberFormatting?: boolean;
  /** Border formatting */
  readonly borderFormatting?: boolean;
  /** Width */
  readonly width?: number;
  /** CLipped */
  readonly clipped?: boolean;
}

export interface QueryTableRefreshOptions {
  /** Next unique ID for new rows */
  readonly nextId?: number;
  /** Minimum refresh version */
  readonly minimumVersion?: number;
  /** Preserve column sort/filter/layout on refresh */
  readonly preserveFormatting?: boolean;
  /** Adjust column width on refresh */
  readonly adjustColumnWidth?: boolean;
  /** Refresh data on load */
  readonly refreshOnLoad?: boolean;
  /** Background refresh */
  readonly backgroundRefresh?: boolean;
  /** Deleted fields */
  readonly deletedFields?: readonly QueryTableDeletedFieldOptions[];
  /** Query table fields */
  readonly queryTableFields?: readonly QueryTableFieldOptions[];
  /** Row count */
  readonly rowCount?: number;
}

export interface QueryTableOptions {
  /** Query table name */
  readonly name?: string;
  /** Connection ID referencing the workbook connection */
  readonly connectionId: number;
  /** Auto-refresh on open */
  readonly autoFormat?: boolean;
  /** Preserve column sort/filter/layout on refresh */
  readonly preserveFormatting?: boolean;
  /** Adjust column width on refresh */
  readonly adjustColumnWidth?: boolean;
  /** Refresh data on load */
  readonly refreshOnLoad?: boolean;
  /** Background refresh */
  readonly backgroundRefresh?: boolean;
  /** Show row numbers (CT_QueryTable @rowNumbers) */
  readonly rowNumbers?: boolean;
  /** Disable refresh (CT_QueryTable @disableRefresh) */
  readonly disableRefresh?: boolean;
  /** First background refresh (CT_QueryTable @firstBackgroundRefresh) */
  readonly firstBackgroundRefresh?: boolean;
  /** Grow/shrink type (CT_QueryTable @growShrinkType) */
  readonly growShrinkType?: boolean;
  /** Fill formulas on refresh (CT_QueryTable @fillFormulas) */
  readonly fillFormulas?: boolean;
  /** Remove data on save (CT_QueryTable @removeDataOnSave) */
  readonly removeDataOnSave?: boolean;
  /** Disable edit (CT_QueryTable @disableEdit) */
  readonly disableEdit?: boolean;
  /** Intermediate (CT_QueryTable @intermediate) */
  readonly intermediate?: boolean;
  /** Query table refresh info (CT_QueryTableRefresh) */
  readonly queryTableRefresh?: QueryTableRefreshOptions;
}

// ── Component ──

export class QueryTableXml extends BaseXmlComponent {
  private readonly opts: QueryTableOptions;

  public constructor(options: QueryTableOptions) {
    super("queryTable");
    this.opts = options;
  }

  public override toXml(_context: Context): string {
    const o = this.opts;
    const a: Record<string, string | number | boolean | undefined> = {
      name: o.name ?? "QueryTable1",
      connectionId: o.connectionId,
      autoFormat: o.autoFormat ? 1 : undefined,
      preserveFormatting: o.preserveFormatting ? 1 : undefined,
      adjustColumnWidth: o.adjustColumnWidth !== false ? 1 : 0,
      refreshOnLoad: o.refreshOnLoad ? 1 : undefined,
      backgroundRefresh: o.backgroundRefresh ? 1 : undefined,
      rowNumbers: o.rowNumbers ? 1 : undefined,
      disableRefresh: o.disableRefresh ? 1 : undefined,
      firstBackgroundRefresh: o.firstBackgroundRefresh ? 1 : undefined,
      growShrinkType: o.growShrinkType ? 1 : undefined,
      fillFormulas: o.fillFormulas ? 1 : undefined,
      removeDataOnSave: o.removeDataOnSave ? 1 : undefined,
      disableEdit: o.disableEdit ? 1 : undefined,
      intermediate: o.intermediate ? 1 : undefined,
    };

    const children: string[] = [];

    // queryTableRefresh
    if (o.queryTableRefresh) {
      children.push(buildQueryTableRefresh(o.queryTableRefresh));
    }

    if (children.length > 0) {
      return `<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"${attrs(a)}>${children.join("")}</queryTable>`;
    }
    return `<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"${attrs(a)}/>`;
  }
}

function buildQueryTableRefresh(opts: QueryTableRefreshOptions): string {
  const a: Record<string, string | number | boolean | undefined> = {};
  if (opts.nextId !== undefined) a.nextId = opts.nextId;
  if (opts.minimumVersion !== undefined) a.minimumVersion = opts.minimumVersion;
  if (opts.preserveFormatting) a.preserveFormatting = 1;
  if (opts.adjustColumnWidth !== undefined) a.adjustColumnWidth = opts.adjustColumnWidth ? 1 : 0;
  if (opts.refreshOnLoad) a.refreshOnLoad = 1;
  if (opts.backgroundRefresh) a.backgroundRefresh = 1;
  if (opts.rowCount !== undefined) a.rowCount = opts.rowCount;

  const children: string[] = [];

  // queryTableDeletedFields
  if (opts.deletedFields && opts.deletedFields.length > 0) {
    const dfParts = opts.deletedFields.map((df) => `<deletedField name="${escapeXml(df.name)}"/>`);
    children.push(
      `<queryTableDeletedFields count="${opts.deletedFields.length}">${dfParts.join("")}</queryTableDeletedFields>`,
    );
  }

  // queryTableFields
  if (opts.queryTableFields && opts.queryTableFields.length > 0) {
    const fParts: string[] = [`<queryTableFields count="${opts.queryTableFields.length}">`];
    for (const f of opts.queryTableFields) {
      const fAttrs: Record<string, string | number | boolean | undefined> = { id: f.id };
      if (f.name !== undefined) fAttrs.name = f.name;
      if (f.tableColumnId !== undefined) fAttrs.tableColumnId = f.tableColumnId;
      if (f.row !== undefined) fAttrs.row = f.row;
      if (f.fillFormatting) fAttrs.fillFormatting = 1;
      if (f.textFormatting) fAttrs.textFormatting = 1;
      if (f.numberFormatting) fAttrs.numberFormatting = 1;
      if (f.borderFormatting) fAttrs.borderFormatting = 1;
      if (f.width !== undefined) fAttrs.width = f.width;
      if (f.clipped) fAttrs.clipped = 1;
      fParts.push(`<queryTableField${attrs(fAttrs)}/>`);
    }
    fParts.push("</queryTableFields>");
    children.push(fParts.join(""));
  }

  // Sort-by-row column (tr — CT_TableRefresh)
  if (children.length > 0) {
    return `<queryTableRefresh${attrs(a)}>${children.join("")}</queryTableRefresh>`;
  }
  return `<queryTableRefresh${attrs(a)}/>`;
}
