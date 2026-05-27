import type { IXmlableObject } from "@file/xml-components";
import { XmlComponent } from "@file/xml-components";

import type { ChangedAttributesProperties } from "../track-revision";

export function buildInsertedTableRowObj(options: ChangedAttributesProperties): IXmlableObject {
  return {
    "w:ins": { _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id } },
  };
}

export class InsertedTableRow extends XmlComponent {
  public constructor(options: ChangedAttributesProperties) {
    super("w:ins");
    this.root.push({
      _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id },
    });
  }
}
