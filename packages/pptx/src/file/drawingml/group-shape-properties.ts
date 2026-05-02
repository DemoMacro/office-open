import { XmlComponent } from "@file/xml-components";

import { GroupTransform2D, type IGroupTransform2DOptions } from "./group-transform-2d";

/**
 * p:grpSpPr — Group shape properties with CT_GroupTransform2D.
 */
export class GroupShapeProperties extends XmlComponent {
    public constructor(options?: IGroupTransform2DOptions) {
        super("p:grpSpPr");
        this.root.push(new GroupTransform2D(options ?? {}));
    }
}
