import { SlideTiming } from "@file/animation/timing";
import type { AnimationOptions } from "@file/animation/types";
import type { Background } from "@file/background/background";
import type { SlideHeaderFooterOptions } from "@file/header-footer/header-footer";
import { coerceChild } from "@file/slide/coerce";
import type { SlideChild } from "@file/slide/slide-child";
import type { TransitionOptions } from "@file/transition/transition";
import { buildTransition } from "@file/transition/transition";
import { BaseXmlComponent, XmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";
import { xml } from "@office-open/xml";

interface Animatable {
  readonly shapeId: number;
  readonly animation?: AnimationOptions;
}

function isAnimatable(child: BaseXmlComponent): child is BaseXmlComponent & Animatable {
  return "shapeId" in child && "animation" in child;
}

function collectAnimations(children: readonly BaseXmlComponent[]): Array<{
  readonly spid: number;
  readonly options: AnimationOptions;
}> {
  const entries: Array<{ readonly spid: number; readonly options: AnimationOptions }> = [];
  for (const child of children) {
    if (isAnimatable(child)) {
      const anim = child.animation;
      if (anim) {
        entries.push({ spid: child.shapeId, options: anim });
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
  private readonly transition?: TransitionOptions;
  public readonly HeaderFooter?: SlideHeaderFooterOptions;

  public constructor(
    children: readonly SlideChild[],
    background?: Background,
    transition?: TransitionOptions,
    headerFooter?: SlideHeaderFooterOptions,
  ) {
    super("p:sld");
    this.children = children;
    this.background = background;
    this.transition = transition;
    this.HeaderFooter = headerFooter;
  }

  public override prepForXml(context: Context): IXmlableObject {
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

  // Direct XML serialization — builds the <p:sld> wrapper as string, but uses
  // prepForXml for child components that have lazy root[] (empty) patterns.
  public override toXml(context: Context): string {
    const parts: string[] = [];

    // Opening tag with namespace declarations
    parts.push(
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">',
    );

    // p:cSld — common slide data
    parts.push("<p:cSld>");
    if (this.background) {
      const bgObj = this.background.prepForXml(context);
      if (bgObj) parts.push(xml(bgObj));
    }

    // p:spTree — shape tree (fixed wrapper + dynamic children)
    parts.push("<p:spTree>");
    parts.push('<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>');
    parts.push(
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>',
    );

    const coercedChildren = this.children.map(coerceChild);
    for (const child of coercedChildren) {
      const s = child.toXml(context);
      if (s) parts.push(s);
    }

    parts.push("</p:spTree>");
    parts.push("</p:cSld>");

    // p:clrMapOvr
    parts.push("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>");

    // p:transition (optional)
    if (this.transition) {
      const transObj = buildTransition(this.transition);
      if (transObj) parts.push(xml(transObj));
    }

    // p:timing (animation)
    const animations = collectAnimations(coercedChildren);
    if (animations.length > 0) {
      const timing = new SlideTiming(animations);
      const timingObj = timing.prepForXml(context);
      if (timingObj) parts.push(xml(timingObj));
    }

    parts.push("</p:sld>");
    return parts.join("");
  }
}
