import type { IXmlableObject } from "@file/xml-components";
import { XmlComponent } from "@file/xml-components";

import { ChangeAttributes } from "../track-revision";
import type { ChangedAttributesProperties } from "../track-revision";

export function buildDeletedTableRowObj(options: ChangedAttributesProperties): IXmlableObject {
  return {
    "w:del": { _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id } },
  };
}

export class DeletedTableRow extends XmlComponent {
  public constructor(options: ChangedAttributesProperties) {
    super("w:del");
    this.root.push(
      new ChangeAttributes({
        author: options.author,
        date: options.date,
        id: options.id,
      }),
    );
  }
}
