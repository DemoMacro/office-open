import type { Context } from "@file/xml-components";
import { BaseXmlComponent } from "@file/xml-components";

export interface SlideSyncOptions {
  readonly serverSldId: string;
  readonly serverSldModifiedTime: string;
  readonly clientInsertedTime: string;
}

/**
 * p:sldSyncPr — Slide synchronization properties (CT_SlideSyncProperties).
 * Standalone top-level element, used as a part root (not nested inside p:sld).
 */
export class SlideSyncProperties extends BaseXmlComponent {
  private readonly opts: SlideSyncOptions;

  public constructor(opts: SlideSyncOptions) {
    super("p:sldSyncPr");
    this.opts = opts;
  }

  public override toXml(_context: Context): string {
    return (
      `<p:sldSyncPr serverSldId="${this.opts.serverSldId}" ` +
      `serverSldModifiedTime="${this.opts.serverSldModifiedTime}" ` +
      `clientInsertedTime="${this.opts.clientInsertedTime}"/>`
    );
  }
}
