import { SlideTiming } from "@file/animation/timing";
import type { AnimationOptions } from "@file/animation/types";
import type { Background } from "@file/background/background";
import type { SlideHeaderFooterOptions } from "@file/header-footer/header-footer";
import { coerceChild } from "@file/slide/coerce";
import type { SlideChild } from "@file/slide/slide-child";
import type { TransitionOptions } from "@file/transition/transition";
import { buildTransition } from "@file/transition/transition";
import { BaseXmlComponent, XmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

interface Animatable {
  readonly shapeId: number;
  readonly animation?: AnimationOptions;
}

export interface ControlOptions {
  readonly shapeId?: number;
  readonly name?: string;
  readonly showAsIcon?: boolean;
  readonly rId?: string;
  readonly imageWidth?: number;
  readonly imageHeight?: number;
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
 * Lazy: stores options, builds XML in toXml().
 */
export class Slide extends XmlComponent {
  private readonly children: readonly SlideChild[];
  private readonly background?: Background;
  private readonly transition?: TransitionOptions;
  public readonly HeaderFooter?: SlideHeaderFooterOptions;
  private readonly showMasterSp?: boolean;
  private readonly showMasterPhAnim?: boolean;
  private readonly controls?: readonly ControlOptions[];
  private readonly customerData?: readonly { readonly rId: string }[];

  public constructor(
    children: readonly SlideChild[],
    background?: Background,
    transition?: TransitionOptions,
    headerFooter?: SlideHeaderFooterOptions,
    showMasterSp?: boolean,
    showMasterPhAnim?: boolean,
    extra?: {
      readonly controls?: readonly ControlOptions[];
      readonly customerData?: readonly { readonly rId: string }[];
    },
  ) {
    super("p:sld");
    this.children = children;
    this.background = background;
    this.transition = transition;
    this.HeaderFooter = headerFooter;
    this.showMasterSp = showMasterSp;
    this.showMasterPhAnim = showMasterPhAnim;
    this.controls = extra?.controls;
    this.customerData = extra?.customerData;
  }

  // Direct XML serialization — builds the <p:sld> wrapper as string, but uses
  // toXml() for child components that have lazy root[] (empty) patterns.
  public override toXml(context: Context): string {
    const parts: string[] = [];

    // Opening tag with namespace declarations
    const sldAttrs: string[] = [];
    if (this.showMasterSp === false) sldAttrs.push(' showMasterSp="0"');
    if (this.showMasterPhAnim === false) sldAttrs.push(' showMasterPhAnim="0"');
    parts.push(
      `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${sldAttrs.join("")}>`,
    );

    // p:cSld — common slide data
    parts.push("<p:cSld>");
    if (this.background) {
      parts.push(this.background.toXml(context));
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

    // custDataLst (E2)
    if (this.customerData && this.customerData.length > 0) {
      const cdItems = this.customerData.map((d) => `<p:custData r:id="${d.rId}"/>`).join("");
      parts.push(`<p:custDataLst>${cdItems}</p:custDataLst>`);
    }

    // controls (E1)
    if (this.controls && this.controls.length > 0) {
      const ctrlItems = this.controls
        .map((c) => {
          const attrs: string[] = [];
          if (c.shapeId !== undefined) attrs.push(`spid="${c.shapeId}"`);
          if (c.name) attrs.push(`name="${c.name}"`);
          if (c.showAsIcon) attrs.push('showAsIcon="1"');
          if (c.rId) attrs.push(`r:id="${c.rId}"`);
          if (c.imageWidth !== undefined) attrs.push(`imgW="${c.imageWidth}"`);
          if (c.imageHeight !== undefined) attrs.push(`imgH="${c.imageHeight}"`);
          return `<p:control ${attrs.join(" ")}/>`;
        })
        .join("");
      parts.push(`<p:controls>${ctrlItems}</p:controls>`);
    }

    parts.push("</p:cSld>");

    // p:clrMapOvr
    parts.push("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>");

    // p:transition (optional)
    if (this.transition) {
      parts.push(buildTransition(this.transition));
    }

    // p:timing (animation)
    const animations = collectAnimations(coercedChildren);
    if (animations.length > 0) {
      const timing = new SlideTiming(animations);
      parts.push(timing.toXml(context));
    }

    parts.push("</p:sld>");
    return parts.join("");
  }
}
