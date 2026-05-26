import { XmlComponent } from "@file/xml-components";

import { ChangeAttributes } from "../track-revision";
import type { ChangedAttributesProperties } from "../track-revision";

export class InsertedTableRow extends XmlComponent {
  public constructor(options: ChangedAttributesProperties) {
    super("w:ins");
    this.root.push(
      new ChangeAttributes({
        author: options.author,
        date: options.date,
        id: options.id,
      }),
    );
  }
}
