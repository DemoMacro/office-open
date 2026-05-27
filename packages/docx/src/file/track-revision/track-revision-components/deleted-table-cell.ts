import type { IXmlableObject } from "@file/xml-components";
import { XmlComponent } from "@file/xml-components";

import type { ChangedAttributesProperties } from "../track-revision";

export function buildDeletedTableCellObj(options: ChangedAttributesProperties): IXmlableObject {
  return {
    "w:cellDel": {
      _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id },
    },
  };
}

export class DeletedTableCell extends XmlComponent {
  public constructor(options: ChangedAttributesProperties) {
    super("w:cellDel");
    this.root.push({
      _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id },
    });
  }
}
