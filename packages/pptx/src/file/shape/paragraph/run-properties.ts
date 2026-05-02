import type { ShapeFill } from "@file/drawingml/shape-properties";
import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

export interface IRunPropertiesOptions {
    readonly fontSize?: number;
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly underline?: "sng" | "dbl" | "none";
    readonly font?: string;
    readonly lang?: string;
    readonly fill?: ShapeFill;
}

/**
 * a:rPr — Run properties (font, size, color, etc.).
 */
export class RunProperties extends XmlComponent {
    public constructor(options: IRunPropertiesOptions = {}) {
        super("a:rPr");

        const attrs: Record<
            string,
            { readonly key: string; readonly value: string | number | boolean | undefined }
        > = {};
        if (options.fontSize) attrs.sz = { key: "sz", value: options.fontSize * 100 };
        if (options.bold !== undefined) attrs.b = { key: "b", value: options.bold };
        if (options.italic !== undefined) attrs.i = { key: "i", value: options.italic };
        if (options.underline) attrs.u = { key: "u", value: options.underline };
        attrs.dirty = { key: "dirty", value: true };
        if (options.lang) attrs.lang = { key: "lang", value: options.lang };

        this.root.push(new NextAttributeComponent(attrs));

        if (options.font) {
            this.root.push(
                new BuilderElement({
                    name: "a:latin",
                    attributes: { typeface: { key: "typeface", value: options.font } },
                }),
            );
            this.root.push(
                new BuilderElement({
                    name: "a:ea",
                    attributes: { typeface: { key: "typeface", value: options.font } },
                }),
            );
        }

        if (options.fill) {
            this.root.push(options.fill);
        }
    }
}
