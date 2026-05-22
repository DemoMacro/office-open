import { XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import {
  createTransform2D,
  type Transform2DOptions as CoreTransform2DOptions,
} from "@office-open/core/drawingml";

export type ITransform2DOptions = CoreTransform2DOptions;

/**
 * a:xfrm / p:xfrm — 2D transform for shapes and graphic frames (position + size in EMUs).
 * Delegates to core createTransform2D.
 */
export class Transform2D extends XmlComponent {
  private readonly core: XmlComponent;

  public constructor(options: ITransform2DOptions, prefix: "a" | "p" = "a") {
    super(`${prefix}:xfrm`);
    this.core = createTransform2D(options, `${prefix}:xfrm`);
  }

  public override prepForXml(context: IContext): IXmlableObject | undefined {
    return this.core["prepForXml"]?.(context);
  }
}
