/**
 * 2D transform (form) for DrawingML shapes.
 *
 * Delegates to core createTransform2D.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_Transform2D
 *
 * @module
 */
import type { IMediaDataTransformation } from "@file/media";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { XmlComponent } from "@file/xml-components";
import { createTransform2D } from "@office-open/core/drawingml";

/**
 * Represents a 2D transformation for DrawingML objects.
 *
 * @example
 * ```typescript
 * const form = new Form({
 *   emus: { x: 914400, y: 914400 },
 *   flip: { horizontal: true },
 *   rotation: 450000 // 7.5 degrees in 60000ths
 * });
 * ```
 */
export class Form extends XmlComponent {
  private readonly core: XmlComponent;

  public constructor(options: IMediaDataTransformation) {
    super("a:xfrm");
    this.core = createTransform2D({
      x: options.offset?.emus?.x ?? 0,
      y: options.offset?.emus?.y ?? 0,
      width: options.emus.x,
      height: options.emus.y,
      flipHorizontal: options.flip?.horizontal,
      flipVertical: options.flip?.vertical,
      rotation: options.rotation,
    });
  }

  public override prepForXml(context: IContext): IXmlableObject | undefined {
    return this.core["prepForXml"]?.(context);
  }
}
