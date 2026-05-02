import { BuilderElement } from "@file/xml-components";

/**
 * a:solidFill — Solid color fill.
 */
export class SolidFill extends BuilderElement<{ readonly val: string }> {
    public constructor(color: string) {
        super({
            name: "a:solidFill",
            children: [
                new BuilderElement<{ readonly val: string }>({
                    name: "a:srgbClr",
                    attributes: { val: { key: "val", value: color.replace("#", "") } },
                }),
            ],
        });
    }
}
