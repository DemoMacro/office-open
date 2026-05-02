import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

export interface OutlineOptions {
    readonly width?: number;
    readonly color?: string;
    readonly dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
}

/**
 * a:ln — Line/outline properties for shapes.
 */
export class Outline extends XmlComponent {
    public constructor(options: OutlineOptions = {}) {
        super("a:ln");

        if (options.width !== undefined) {
            this.root.push(
                new NextAttributeComponent({
                    w: { key: "w", value: options.width },
                }),
            );
        }

        if (options.color) {
            this.root.push(
                new BuilderElement({
                    name: "a:solidFill",
                    children: [
                        new BuilderElement({
                            name: "a:srgbClr",
                            attributes: {
                                val: { key: "val", value: options.color.replace("#", "") },
                            },
                        }),
                    ],
                }),
            );
        } else {
            this.root.push(new BuilderElement({ name: "a:noFill" }));
        }

        if (options.dashStyle) {
            this.root.push(
                new BuilderElement({
                    name: "a:prstDash",
                    attributes: {
                        val: { key: "val", value: options.dashStyle },
                    },
                }),
            );
        }
    }
}
