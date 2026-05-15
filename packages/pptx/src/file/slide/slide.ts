import { SlideTiming } from "@file/animation/timing";
import type { IAnimationOptions } from "@file/animation/types";
import type { Background } from "@file/background/background";
import type { IHeaderFooterOptions } from "@file/header-footer/header-footer";
import { AudioFrame } from "@file/media/audio-frame";
import { VideoFrame } from "@file/media/video-frame";
import { Shape } from "@file/shape/shape";
import type { ITransitionOptions } from "@file/transition/transition";
import { buildTransition } from "@file/transition/transition";
import { BaseXmlComponent, XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { xml } from "@office-open/xml";

const NV_GRP_SP_PR = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>';

const GRP_SP_PR =
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

const CLR_MAP_OVR = "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>";

const XMLNS =
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

function collectAnimations(children: readonly BaseXmlComponent[]): Array<{
    readonly spid: number;
    readonly options: IAnimationOptions;
}> {
    const entries: Array<{ readonly spid: number; readonly options: IAnimationOptions }> = [];
    for (const child of children) {
        let anim: IAnimationOptions | undefined;
        let spid: number | undefined;
        if (child instanceof Shape) {
            anim = child.Animation;
            spid = child.ShapeId;
        } else if (child instanceof VideoFrame) {
            anim = child.Animation;
            spid = child.ShapeId;
        } else if (child instanceof AudioFrame) {
            anim = child.Animation;
            spid = child.ShapeId;
        }
        if (anim && spid !== undefined) {
            entries.push({ spid, options: anim });
        }
    }
    return entries;
}

/**
 * p:sld — A slide in a presentation.
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class Slide extends XmlComponent {
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
        children.push({
            _attr: {
                "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "xmlns:p": "http://schemas.openxmlformats.org/presentationml/2006/main",
            },
        });

        // p:cSld — common slide data (background + shape tree)
        const cSldChildren: IXmlableObject[] = [];
        if (this.background) {
            const bgObj = this.background.prepForXml(context);
            if (bgObj) cSldChildren.push(bgObj);
        }

        // p:spTree — shape tree
        const spTreeChildren: IXmlableObject[] = [
            {
                "p:nvGrpSpPr": [
                    { "p:cNvPr": { _attr: { id: 1, name: "" } } },
                    { "p:cNvGrpSpPr": {} },
                    { "p:nvPr": {} },
                ],
            },
            {
                "p:grpSpPr": {
                    "a:xfrm": [
                        { "a:off": { _attr: { x: 0, y: 0 } } },
                        { "a:ext": { _attr: { cx: 0, cy: 0 } } },
                        { "a:chOff": { _attr: { x: 0, y: 0 } } },
                        { "a:chExt": { _attr: { cx: 0, cy: 0 } } },
                    ],
                },
            },
        ];
        for (const child of this.children) {
            const obj = child.prepForXml(context);
            if (obj) spTreeChildren.push(obj);
        }
        cSldChildren.push({ "p:spTree": spTreeChildren });

        children.push({ "p:cSld": cSldChildren });

        // p:clrMapOvr
        children.push({ "p:clrMapOvr": [{ "a:masterClrMapping": {} }] });

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

    public toXml(context: IContext): string {
        let s = `<p:sld${XMLNS}>`;

        // p:cSld
        s += "<p:cSld>";
        if (this.background) {
            const bgObj = this.background.prepForXml(context);
            if (bgObj) s += xml(bgObj);
        }

        // p:spTree
        s += "<p:spTree>";
        s += NV_GRP_SP_PR;
        s += GRP_SP_PR;
        for (const child of this.children) {
            if (typeof (child as unknown as { toXml?: Function }).toXml === "function") {
                s += (child as unknown as { toXml(ctx: IContext): string }).toXml(context);
            } else {
                const obj = child.prepForXml(context);
                if (obj) s += xml(obj);
            }
        }
        s += "</p:spTree></p:cSld>";

        s += CLR_MAP_OVR;

        if (this.transition) {
            const transObj = buildTransition(this.transition);
            if (transObj) s += xml(transObj);
        }

        const animations = collectAnimations(this.children);
        if (animations.length > 0) {
            const timing = new SlideTiming(animations);
            const timingObj = timing.prepForXml(context);
            if (timingObj) s += xml(timingObj);
        }

        s += "</p:sld>";
        return s;
    }
}
