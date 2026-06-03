import type { Context } from "@file/xml-components";
import { BaseXmlComponent } from "@file/xml-components";
import type { TableStyleListOptions } from "@office-open/core";
import { createTableStyleList } from "@office-open/core";

const DEFAULT_STYLE_ID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

export class TableStyles extends BaseXmlComponent {
  private readonly opts?: TableStyleListOptions;

  public constructor(opts?: TableStyleListOptions) {
    super("a:tblStyleLst");
    this.opts = opts;
  }

  public override toXml(context: Context): string {
    if (!this.opts) {
      return `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${DEFAULT_STYLE_ID}"/>`;
    }
    return createTableStyleList(this.opts).toXml(context);
  }
}
