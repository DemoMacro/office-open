import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";

export interface IGroupShapeOptions {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly rotation?: number;
    readonly flipH?: boolean;
    readonly children: readonly BaseXmlComponent[];
}

/**
 * p:grpSp — Group shape containing child shapes.
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class GroupShape extends BaseXmlComponent {
    private static nextId = 100;
    private readonly id: number;
    private readonly options: IGroupShapeOptions;

    public constructor(options: IGroupShapeOptions) {
        super("p:grpSp");
        this.id = GroupShape.nextId++;
        this.options = options;
    }

    public override prepForXml(context: IContext): IXmlableObject {
        const opts = this.options;
        const id = this.id;
        const name = `Group ${id}`;
        const children: IXmlableObject[] = [];

        // p:nvGrpSpPr
        children.push({
            "p:nvGrpSpPr": [
                { "p:cNvPr": { _attr: { id, name } } },
                { "p:cNvGrpSpPr": {} },
                { "p:nvPr": {} },
            ],
        });

        // p:grpSpPr
        const xfrmChildren: IXmlableObject[] = [];
        const xfrmAttrs: Record<string, string | number> = {};
        if (opts.flipH !== undefined) xfrmAttrs.flipH = opts.flipH ? 1 : 0;
        if (opts.rotation !== undefined) xfrmAttrs.rot = opts.rotation;
        if (Object.keys(xfrmAttrs).length > 0) xfrmChildren.push({ _attr: xfrmAttrs });
        xfrmChildren.push({
            "a:off": {
                _attr: {
                    x: opts.x !== undefined ? pixelsToEmus(opts.x) : 0,
                    y: opts.y !== undefined ? pixelsToEmus(opts.y) : 0,
                },
            },
        });
        xfrmChildren.push({
            "a:ext": {
                _attr: {
                    cx: opts.width !== undefined ? pixelsToEmus(opts.width) : 0,
                    cy: opts.height !== undefined ? pixelsToEmus(opts.height) : 0,
                },
            },
        });
        xfrmChildren.push({ "a:chOff": { _attr: { x: 0, y: 0 } } });
        xfrmChildren.push({ "a:chExt": { _attr: { cx: 0, cy: 0 } } });
        children.push({ "p:grpSpPr": { "a:xfrm": xfrmChildren } });

        // Children
        for (const child of opts.children) {
            const obj = child.prepForXml(context);
            if (obj) children.push(obj);
        }

        return { "p:grpSp": children };
    }
}
