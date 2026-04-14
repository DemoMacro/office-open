import { XmlComponent } from "@file/xml-components";

import { ChangeAttributes } from "../track-revision";
import type { IChangedAttributesProperties } from "../track-revision";

export class DeletedTableCell extends XmlComponent {
    public constructor(options: IChangedAttributesProperties) {
        super("w:cellDel");
        this.root.push(
            new ChangeAttributes({
                author: options.author,
                date: options.date,
                id: options.id,
            }),
        );
    }
}
