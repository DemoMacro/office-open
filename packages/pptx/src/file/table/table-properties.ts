import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

/**
 * a:tblPr — Table properties (firstRow, bandRow, etc.).
 * Lazy: stores options, builds XML in toXml().
 */
export class TableProperties extends BaseXmlComponent {
  private readonly options?: {
    readonly firstRow?: boolean;
    readonly lastRow?: boolean;
    readonly bandRow?: boolean;
    readonly firstCol?: boolean;
    readonly lastCol?: boolean;
    readonly bandCol?: boolean;
    readonly tableStyleId?: string;
  };

  public constructor(options?: {
    readonly firstRow?: boolean;
    readonly lastRow?: boolean;
    readonly bandRow?: boolean;
    readonly firstCol?: boolean;
    readonly lastCol?: boolean;
    readonly bandCol?: boolean;
    readonly tableStyleId?: string;
  }) {
    super("a:tblPr");
    this.options = options;
  }

  public override toXml(_context: Context): string {
    if (!this.options) return "<a:tblPr/>";
    const opts = this.options;
    const attrs: string[] = [];
    if (opts.firstRow !== undefined) attrs.push(`firstRow="${opts.firstRow ? 1 : 0}"`);
    if (opts.lastRow !== undefined) attrs.push(`lastRow="${opts.lastRow ? 1 : 0}"`);
    if (opts.bandRow !== undefined) attrs.push(`bandRow="${opts.bandRow ? 1 : 0}"`);
    if (opts.firstCol !== undefined) attrs.push(`firstCol="${opts.firstCol ? 1 : 0}"`);
    if (opts.lastCol !== undefined) attrs.push(`lastCol="${opts.lastCol ? 1 : 0}"`);
    if (opts.bandCol !== undefined) attrs.push(`bandCol="${opts.bandCol ? 1 : 0}"`);
    if (attrs.length === 0 && !opts.tableStyleId) return "<a:tblPr/>";
    const styleId = opts.tableStyleId
      ? `<a:tableStyleId>${opts.tableStyleId}</a:tableStyleId>`
      : "";
    return `<a:tblPr ${attrs.join(" ")}>${styleId}</a:tblPr>`;
  }
}
