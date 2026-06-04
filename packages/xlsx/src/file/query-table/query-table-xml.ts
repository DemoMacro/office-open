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
    };

    return `<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"${attrs(a)}/>`;
  }
}
