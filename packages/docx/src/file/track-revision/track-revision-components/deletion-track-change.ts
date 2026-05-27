import { XmlComponent } from "@file/xml-components";

import type { ChangedAttributesProperties } from "../track-revision";

export class DeletionTrackChange extends XmlComponent {
  public constructor(options: ChangedAttributesProperties) {
    super("w:del");
    this.root.push({
      _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id },
    });
  }
}
