import { XmlComponent } from "@file/xml-components";

import { ChangeAttributes } from "../track-revision";
import type { IChangedAttributesProperties } from "../track-revision";

export class DeletedTableRow extends XmlComponent {
    public constructor(options: IChangedAttributesProperties) {
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
