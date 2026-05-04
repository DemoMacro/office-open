import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

import { buildFill } from "../drawingml/fill";
import type { FillOptions } from "../drawingml/fill";

export interface ICellBorderOptions {
    readonly width?: number;
    readonly color?: string;
    readonly dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
}

function buildBorderLine(name: string, options: ICellBorderOptions): IXmlableObject {
    const attrs: Record<string, string | number> = {};
    const children: IXmlableObject[] = [];

    if (options.width !== undefined) attrs.w = options.width;
    if (options.color) {
        children.push({
            "a:solidFill": [{ "a:srgbClr": { _attr: { val: options.color.replace("#", "") } } }],
        });
    }
    if (options.dashStyle) {
        children.push({ "a:prstDash": { _attr: { val: options.dashStyle } } });
    }

    return { [name]: Object.keys(attrs).length > 0 ? [{ _attr: attrs }, ...children] : children };
}

/**
 * a:tcPr — Table cell properties (borders + fill).
 * Lazy: stores options, builds IXmlableObject in prepForXml.
 */
export class TableCellProperties extends BaseXmlComponent {
    private readonly options?: {
        readonly fill?: FillOptions;
        readonly borders?: {
            readonly top?: ICellBorderOptions;
            readonly bottom?: ICellBorderOptions;
            readonly left?: ICellBorderOptions;
            readonly right?: ICellBorderOptions;
        };
        readonly verticalAlign?: "t" | "ctr" | "b";
    };

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
        this.options = options;
    }

    public override prepForXml(context: IContext): IXmlableObject {
        const children: IXmlableObject[] = [];
        const opts = this.options;

        if (opts?.verticalAlign) {
            children.push({ _attr: { anchor: opts.verticalAlign } });
        }

        if (opts?.borders) {
            if (opts.borders.left) children.push(buildBorderLine("a:lnL", opts.borders.left));
            if (opts.borders.right) children.push(buildBorderLine("a:lnR", opts.borders.right));
            if (opts.borders.top) children.push(buildBorderLine("a:lnT", opts.borders.top));
            if (opts.borders.bottom) children.push(buildBorderLine("a:lnB", opts.borders.bottom));
        }

        if (opts?.fill !== undefined) {
            const fillComp = buildFill(opts.fill);
            const fillObj = fillComp.prepForXml(context);
            if (fillObj) children.push(fillObj);
        }

        return {
            "a:tcPr":
                children.length === 0
                    ? {}
                    : children.length === 1 && "_attr" in children[0]
                      ? children[0]
                      : children,
        };
    }
}
