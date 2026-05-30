import {
  buildAttrObject,
  BuilderElement,
  XmlComponent,
  stringContainerObj,
} from "@file/xml-components";

export interface SlideHeaderFooterOptions {
  readonly slideNumber?: boolean;
  readonly dateTime?: boolean;
  readonly footer?: string | boolean;
  readonly header?: boolean;
}

/**
 * p:hf — Slide header/footer settings.
 *
 * CT_HeaderFooter has boolean attributes (sldNum, hdr, ftr, dt) for visibility.
 * Footer text content is passed as p:ftr child element (legacy support).
 */
export class HeaderFooter extends XmlComponent {
  public constructor(options: SlideHeaderFooterOptions = {}) {
    super("p:hf");

    // Boolean visibility attributes
    const attrs: Record<string, number> = {};
    if (options.slideNumber !== false) attrs.sldNum = 1;
    if (options.dateTime !== false) attrs.dt = 1;
    // Only set hdr when explicitly true (no header placeholder in default slide master)
    if (options.header === true) attrs.hdr = 1;
    // Only set ftr when explicitly true or footer has string content
    if (options.footer !== false && options.footer !== undefined) attrs.ftr = 1;

    if (Object.keys(attrs).length > 0) {
      this.root.push(buildAttrObject(attrs));
    }

    // Footer text content as child element
    if (typeof options.footer === "string") {
      this.root.push(
        new BuilderElement({
          name: "p:ftr",
          children: [
            new BuilderElement({
              name: "p:txBody",
              children: [
                new BuilderElement({ name: "a:bodyPr" }),
                new BuilderElement({ name: "a:lstStyle" }),
                new BuilderElement({
                  name: "a:p",
                  children: [
                    new BuilderElement({
                      name: "a:r",
                      children: [stringContainerObj("a:t", options.footer)],
                    }),
                    new BuilderElement({
                      name: "a:endParaRPr",
                      attributes: { lang: { key: "lang", value: "en-US" } },
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      );
    }
  }
}
