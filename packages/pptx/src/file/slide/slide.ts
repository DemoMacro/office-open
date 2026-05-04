import { SlideTiming } from "@file/animation/timing";
import type { Background } from "@file/background/background";
import type { IHeaderFooterOptions } from "@file/header-footer/header-footer";
import { Shape } from "@file/shape/shape";
import type { ITransitionOptions } from "@file/transition/transition";
import { buildTransition } from "@file/transition/transition";
import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

/**
 * p:nvGrpSpPr singleton — non-visual properties for shape tree.
 */
const NV_GRP_SP_PR: IXmlableObject = {
    "p:nvGrpSpPr": [
        { "p:cNvPr": { _attr: { id: 1, name: "" } } },
        { "p:cNvGrpSpPr": {} },
        { "p:nvPr": {} },
    ],
};

/**
 * p:grpSpPr singleton — group shape properties with default transform.
 */
const GRP_SP_PR: IXmlableObject = {
    "p:grpSpPr": {
        "a:xfrm": [
            { "a:off": { _attr: { x: 0, y: 0 } } },
            { "a:ext": { _attr: { cx: 0, cy: 0 } } },
            { "a:chOff": { _attr: { x: 0, y: 0 } } },
            { "a:chExt": { _attr: { cx: 0, cy: 0 } } },
        ],
    },
};

/**
 * p:clrMapOvr singleton — color map override.
 */
const CLR_MAP_OVR: IXmlableObject = {
    "p:clrMapOvr": [{ "a:masterClrMapping": {} }],
};

const XMLNS_ATTRS: IXmlableObject = {
    _attr: {
        "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    },
};

function collectAnimations(children: readonly BaseXmlComponent[]): Array<{
    readonly spid: number;
    readonly options: import("@file/animation/types").IAnimationOptions;
}> {
    const entries: Array<{
        readonly spid: number;
        readonly options: import("@file/animation/types").IAnimationOptions;
    }> = [];
    for (const child of children) {
        if (child instanceof Shape) {
            const anim = child.Animation;
            if (anim) {
                entries.push({ spid: child.ShapeId, options: anim });
            }
        }
    }
    return entries;
}

/**
 * p:sld — A slide in a presentation.
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class Slide extends BaseXmlComponent {
    private readonly children: readonly BaseXmlComponent[];
    private readonly background?: Background;
    private readonly transition?: ITransitionOptions;
    public readonly HeaderFooter?: IHeaderFooterOptions;

    public constructor(
        children: readonly BaseXmlComponent[],
        background?: Background,
        transition?: ITransitionOptions,
        headerFooter?: IHeaderFooterOptions,
    ) {
        super("p:sld");
        this.children = children;
        this.background = background;
        this.transition = transition;
        this.HeaderFooter = headerFooter;
    }

    public override prepForXml(context: IContext): IXmlableObject {
        const children: IXmlableObject[] = [];

        // xmlns attributes
        children.push(XMLNS_ATTRS);

        // p:cSld — common slide data (background + shape tree)
        const cSldChildren: IXmlableObject[] = [];
        if (this.background) {
            const bgObj = this.background.prepForXml(context);
            if (bgObj) cSldChildren.push(bgObj);
        }

        // p:spTree — shape tree
        const spTreeChildren: IXmlableObject[] = [NV_GRP_SP_PR, GRP_SP_PR];
        for (const child of this.children) {
            const obj = child.prepForXml(context);
            if (obj) spTreeChildren.push(obj);
        }
        cSldChildren.push({ "p:spTree": spTreeChildren });

        children.push({ "p:cSld": cSldChildren });

        // p:clrMapOvr
        children.push(CLR_MAP_OVR);

        // p:transition (optional)
        if (this.transition) {
            const transObj = buildTransition(this.transition);
            if (transObj) children.push(transObj);
        }

        // p:timing (animation)
        const animations = collectAnimations(this.children);
        if (animations.length > 0) {
            const timing = new SlideTiming(animations);
            const timingObj = timing.prepForXml(context);
            if (timingObj) children.push(timingObj);
        }

        return { "p:sld": children };
    }
}
