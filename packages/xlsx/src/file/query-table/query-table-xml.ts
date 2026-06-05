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
import { attrs } from "@office-open/xml";

// ── Options ──

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

    return `<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"${attrs(a)}/>`;
  }
}
