import type { TextBodyOptions } from "@file/shape/text-body";
import { TextBody } from "@file/shape/text-body";
import type { Context } from "@file/xml-components";
import { BaseXmlComponent } from "@file/xml-components";
import { convertPixelsToEmu } from "@office-open/core";

export interface LockedCanvasShapeOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly geometry?: string;
  readonly fill?: string;
  readonly textBody?: TextBodyOptions;
}

export interface LockedCanvasFrameOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly id?: number;
  readonly name?: string;
  readonly children?: readonly LockedCanvasShapeOptions[];
}

let nextLockedCanvasId = 2048;
let nextCanvasShapeId = 2;

/**
 * p:graphicFrame — Slide-level graphic frame wrapping a locked canvas.
 * Lazy: stores options, builds XML in toXml().
 */
export class LockedCanvasFrame extends BaseXmlComponent {
  private readonly options: LockedCanvasFrameOptions;

  public constructor(options: LockedCanvasFrameOptions = {}) {
    super("p:graphicFrame");
    this.options = options;
  }

  public override toXml(context: Context): string {
    const opts = this.options;
    const id = nextLockedCanvasId++;
    const name = opts.name ?? `Locked Canvas ${id}`;

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`,
    );

    // p:xfrm
    const x = convertPixelsToEmu(opts.x ?? 0);
    const y = convertPixelsToEmu(opts.y ?? 0);
    const cx = convertPixelsToEmu(opts.width ?? 0);
    const cy = convertPixelsToEmu(opts.height ?? 0);
    parts.push(`<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>`);

    // a:graphic > a:graphicData > lc:lockedCanvas
    const canvasChildren = buildCanvasChildren(opts.children, context);
    const grpSpPr = `<a:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/><a:chOff x="0" y="0"/><a:chExt cx="${cx}" cy="${cy}"/></a:xfrm></a:grpSpPr>`;
    parts.push(
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"><lc:lockedCanvas xmlns:lc="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"><a:nvGrpSpPr><a:cNvPr id="0" name=""/><a:cNvGrpSpPr/></a:nvGrpSpPr>${grpSpPr}${canvasChildren}</lc:lockedCanvas></a:graphicData></a:graphic>`,
    );

    return `<p:graphicFrame>${parts.join("")}</p:graphicFrame>`;
  }
}

function buildCanvasChildren(
  children: readonly LockedCanvasShapeOptions[] | undefined,
  context: Context,
): string {
  if (!children || children.length === 0) return "";

  const parts: string[] = [];
  for (const opts of children) {
    const id = nextCanvasShapeId++;
    const spParts: string[] = [];

    // a:nvSpPr
    spParts.push(`<a:nvSpPr><a:cNvPr id="${id}" name="Shape ${id}"/><a:cNvSpPr/></a:nvSpPr>`);

    // a:spPr
    const spPrParts: string[] = [];
    const x = convertPixelsToEmu(opts.x ?? 0);
    const y = convertPixelsToEmu(opts.y ?? 0);
    const cx = convertPixelsToEmu(opts.width ?? 0);
    const cy = convertPixelsToEmu(opts.height ?? 0);
    spPrParts.push(`<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`);
    spPrParts.push(`<a:prstGeom prst="${opts.geometry ?? "rect"}"><a:avLst/></a:prstGeom>`);
    if (opts.fill) {
      spPrParts.push(`<a:solidFill><a:srgbClr val="${opts.fill}"/></a:solidFill>`);
    }
    spParts.push(`<a:spPr>${spPrParts.join("")}</a:spPr>`);

    // a:txSp > a:txBody (DrawingML text shape)
    const txBodyXml = new TextBody(opts.textBody ?? {}, "a").toXml(context);
    spParts.push(`<a:txSp>${txBodyXml}<a:useSpRect/></a:txSp>`);

    parts.push(`<a:sp>${spParts.join("")}</a:sp>`);
  }

  return parts.join("");
}
