import { SlideTiming } from "@file/animation/timing";
import type { IAnimationOptions } from "@file/animation/types";
import type { Background } from "@file/background/background";
import type { IHeaderFooterOptions } from "@file/header-footer/header-footer";
import { coerceChild } from "@file/slide/coerce";
import type { SlideChild } from "@file/slide/slide-child";
import type { ITransitionOptions } from "@file/transition/transition";
import { buildTransition } from "@file/transition/transition";
import { BaseXmlComponent, XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

interface IAnimatable {
  readonly ShapeId: number;
  readonly Animation?: IAnimationOptions;
}

function isAnimatable(child: BaseXmlComponent): child is BaseXmlComponent & IAnimatable {
  return "ShapeId" in child && "Animation" in child;
}

function collectAnimations(children: readonly BaseXmlComponent[]): Array<{
  readonly spid: number;
  readonly options: IAnimationOptions;
}> {
  const entries: Array<{ readonly spid: number; readonly options: IAnimationOptions }> = [];
  for (const child of children) {
    if (isAnimatable(child)) {
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
export class Slide extends XmlComponent {
  private readonly children: readonly SlideChild[];
  private readonly background?: Background;
  private readonly transition?: ITransitionOptions;
  public readonly HeaderFooter?: IHeaderFooterOptions;

  public constructor(
    children: readonly SlideChild[],
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
        "p:grpSpPr": [
          {
            "a:xfrm": [
              { "a:off": { _attr: { x: 0, y: 0 } } },
              { "a:ext": { _attr: { cx: 0, cy: 0 } } },
              { "a:chOff": { _attr: { x: 0, y: 0 } } },
              { "a:chExt": { _attr: { cx: 0, cy: 0 } } },
            ],
          },
        ],
      },
    ];
    const coercedChildren = this.children.map(coerceChild);
    for (const child of coercedChildren) {
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
    const animations = collectAnimations(coercedChildren);
    if (animations.length > 0) {
      const timing = new SlideTiming(animations);
      const timingObj = timing.prepForXml(context);
      if (timingObj) children.push(timingObj);
    }

    return { "p:sld": children };
  }
}
