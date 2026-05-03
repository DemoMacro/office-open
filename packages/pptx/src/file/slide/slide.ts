import { SlideTiming } from "@file/animation/timing";
import type { Background } from "@file/background/background";
import type { IHeaderFooterOptions } from "@file/header-footer/header-footer";
import { Shape } from "@file/shape/shape";
import type { ITransitionOptions } from "@file/transition/transition";
import { Transition } from "@file/transition/transition";
import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

import { CommonSlideData } from "./common-slide-data";

class ColorMapOverride extends BuilderElement<{}> {
    public constructor() {
        super({
            name: "p:clrMapOvr",
            children: [new BuilderElement({ name: "a:masterClrMapping" })],
        });
    }
}

function collectAnimations(children: readonly XmlComponent[]): Array<{
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
 */
export class Slide extends XmlComponent {
    public readonly HeaderFooter?: IHeaderFooterOptions;

    public constructor(
        children: readonly XmlComponent[],
        background?: Background,
        transition?: ITransitionOptions,
        headerFooter?: IHeaderFooterOptions,
    ) {
        super("p:sld");
        this.HeaderFooter = headerFooter;
        this.root.push(
            new NextAttributeComponent({
                "xmlns:a": {
                    key: "xmlns:a",
                    value: "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                "xmlns:r": {
                    key: "xmlns:r",
                    value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                },
                "xmlns:p": {
                    key: "xmlns:p",
                    value: "http://schemas.openxmlformats.org/presentationml/2006/main",
                },
            }),
        );
        this.root.push(new CommonSlideData(children, background));
        this.root.push(new ColorMapOverride());
        if (transition) {
            this.root.push(new Transition(transition));
        }

        const animations = collectAnimations(children);
        if (animations.length > 0) {
            this.root.push(new SlideTiming(animations));
        }
    }
}
