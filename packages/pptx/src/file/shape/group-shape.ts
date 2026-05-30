import { coerceChild } from "@file/slide/coerce";
import type { SlideChild } from "@file/slide/slide-child";
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { convertPixelsToEmu } from "@office-open/core";

export interface GroupShapeOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly rotation?: number;
  readonly flipHorizontal?: boolean;
  readonly children: readonly SlideChild[];
}

/**
 * p:grpSp — Group shape containing child shapes.
 * Lazy: stores options, builds XML in toXml().
 */
export class GroupShape extends BaseXmlComponent {
  private static nextId = 100;
  private readonly id: number;
  private readonly options: GroupShapeOptions;

  public constructor(options: GroupShapeOptions) {
    super("p:grpSp");
    this.id = GroupShape.nextId++;
    this.options = options;
  }

  public override toXml(context: Context): string {
    const opts = this.options;
    const id = this.id;
    const parts: string[] = [];

    // p:nvGrpSpPr
    parts.push(
      `<p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`,
    );

    // p:grpSpPr
    const xfrmAttrs: string[] = [];
    if (opts.flipHorizontal !== undefined) xfrmAttrs.push(`flipH="${opts.flipHorizontal ? 1 : 0}"`);
    if (opts.rotation !== undefined) xfrmAttrs.push(`rot="${opts.rotation}"`);
    const xfrmAttrStr = xfrmAttrs.length > 0 ? ` ${xfrmAttrs.join(" ")}` : "";
    parts.push(
      `<p:grpSpPr><a:xfrm${xfrmAttrStr}><a:off x="${convertPixelsToEmu(opts.x ?? 0)}" y="${convertPixelsToEmu(opts.y ?? 0)}"/><a:ext cx="${convertPixelsToEmu(opts.width ?? 0)}" cy="${convertPixelsToEmu(opts.height ?? 0)}"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`,
    );

    // Child shapes
    for (const rawChild of opts.children) {
      const child = coerceChild(rawChild);
      const s = child.toXml(context);
      if (s) parts.push(s);
    }

    return `<p:grpSp>${parts.join("")}</p:grpSp>`;
  }
}
