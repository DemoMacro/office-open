import { XmlComponent } from "@file/xml-components";

import { ChangeAttributes } from "../track-revision";
import type { ChangedAttributesProperties } from "../track-revision";

export class InsertedTableCell extends XmlComponent {
  public constructor(options: ChangedAttributesProperties) {
    super("w:cellIns");
    this.root.push(
      new ChangeAttributes({
        author: options.author,
        date: options.date,
        id: options.id,
      }),
    );
  }
}
