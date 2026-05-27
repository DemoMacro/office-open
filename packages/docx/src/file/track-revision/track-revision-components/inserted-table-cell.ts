import type { IXmlableObject } from "@file/xml-components";
import { XmlComponent } from "@file/xml-components";

import type { ChangedAttributesProperties } from "../track-revision";

export function buildInsertedTableCellObj(options: ChangedAttributesProperties): IXmlableObject {
  return {
    "w:cellIns": {
      _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id },
    },
  };
}

export class InsertedTableCell extends XmlComponent {
  public constructor(options: ChangedAttributesProperties) {
    super("w:cellIns");
    this.root.push({
      _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id },
    });
  }
}
