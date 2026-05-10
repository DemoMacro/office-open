import { XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import {
    createGroupTransform2D,
    type GroupTransform2DOptions as CoreGroupTransform2DOptions,
} from "@office-open/core/drawingml";

export type IGroupTransform2DOptions = CoreGroupTransform2DOptions;

/**
 * a:xfrm — Group transform (CT_GroupTransform2D).
 * Extends regular Transform2D with child offset (chOff) and child extent (chExt).
 * Delegates to core createGroupTransform2D.
 */
export class GroupTransform2D extends XmlComponent {
    private readonly core: XmlComponent;

    public constructor(options: IGroupTransform2DOptions, prefix: "a" | "p" = "a") {
        super(`${prefix}:xfrm`);
        this.core = createGroupTransform2D(options, `${prefix}:xfrm`);
    }

    public override prepForXml(context: IContext): IXmlableObject | undefined {
        return this.core["prepForXml"]?.(context);
    }
}
