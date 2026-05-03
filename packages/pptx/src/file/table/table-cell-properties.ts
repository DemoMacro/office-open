import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

import { buildFill } from "../drawingml/fill";
import type { FillOptions } from "../drawingml/fill";

export interface ICellBorderOptions {
    readonly width?: number;
    readonly color?: string;
    readonly dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
}

/**
 * a:tcPr — Table cell properties (borders + fill).
 * XSD order: lnL → lnR → lnT → lnB → lnTlToBr → lnBlToTr → cell3D → EG_FillProperties → headers → extLst
 * anchor is an attribute, not a child element.
 */
export class TableCellProperties extends XmlComponent {
    public constructor(options?: {
        readonly fill?: FillOptions;
        readonly borders?: {
            readonly top?: ICellBorderOptions;
            readonly bottom?: ICellBorderOptions;
            readonly left?: ICellBorderOptions;
            readonly right?: ICellBorderOptions;
        };
        readonly verticalAlign?: "t" | "ctr" | "b";
    }) {
        super("a:tcPr");

        if (options?.verticalAlign) {
            this.root.push(
                new NextAttributeComponent({
                    anchor: { key: "anchor", value: options.verticalAlign },
                }),
            );
        }

        if (options?.borders) {
            if (options.borders.left) {
                this.root.push(new BorderLine("a:lnL", options.borders.left));
            }
            if (options.borders.right) {
                this.root.push(new BorderLine("a:lnR", options.borders.right));
            }
            if (options.borders.top) {
                this.root.push(new BorderLine("a:lnT", options.borders.top));
            }
            if (options.borders.bottom) {
                this.root.push(new BorderLine("a:lnB", options.borders.bottom));
            }
        }

        if (options?.fill !== undefined) {
            this.root.push(buildFill(options.fill));
        }
    }
}

class BorderLine extends BuilderElement<{}> {
    public constructor(name: string, options: ICellBorderOptions) {
        const children: BuilderElement<{}>[] = [];
        const attributes: Record<
            string,
            { readonly key: string; readonly value: string | number | undefined }
        > = {};

        if (options.width !== undefined) {
            attributes.w = { key: "w", value: options.width };
        }
        if (options.color) {
            children.push(
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
        }
        if (options.dashStyle) {
            children.push(
                new BuilderElement({
                    name: "a:prstDash",
                    attributes: { val: { key: "val", value: options.dashStyle } },
                }),
            );
        }

        super({
            name,
            attributes: Object.keys(attributes).length > 0 ? attributes : undefined,
            children: children.length > 0 ? children : undefined,
        });
    }
}
