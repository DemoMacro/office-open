import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

import { buildFill } from "../drawingml/fill";
import type { FillOptions } from "../drawingml/fill";

export interface CellBorderOptions {
  readonly width?: number;
  readonly color?: string;
  readonly dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
}

function buildBorderLine(name: string, options: CellBorderOptions): string {
  const attrs: string[] = [];
  if (options.width !== undefined) attrs.push(`w="${options.width}"`);

  const children: string[] = [];
  if (options.color) {
    children.push(
      `<a:solidFill><a:srgbClr val="${options.color.replace("#", "")}"/></a:solidFill>`,
    );
  }
  if (options.dashStyle) {
    children.push(`<a:prstDash val="${options.dashStyle}"/>`);
  }

  const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  if (children.length === 0) return `<${name}${attrStr}/>`;
  return `<${name}${attrStr}>${children.join("")}</${name}>`;
}

/**
 * a:tcPr — Table cell properties (borders + fill).
 */
export class TableCellProperties extends BaseXmlComponent {
  private readonly options?: {
    readonly fill?: FillOptions;
    readonly borders?: {
      readonly top?: CellBorderOptions;
      readonly bottom?: CellBorderOptions;
      readonly left?: CellBorderOptions;
      readonly right?: CellBorderOptions;
    };
    readonly verticalAlign?: "t" | "ctr" | "b";
  };

  public constructor(options?: {
    readonly fill?: FillOptions;
    readonly borders?: {
      readonly top?: CellBorderOptions;
      readonly bottom?: CellBorderOptions;
      readonly left?: CellBorderOptions;
      readonly right?: CellBorderOptions;
    };
    readonly verticalAlign?: "t" | "ctr" | "b";
  }) {
    super("a:tcPr");
    this.options = options;
  }

  public override toXml(context: Context): string {
    const parts: string[] = [];
    const opts = this.options;

    if (opts?.verticalAlign) {
      parts.push(`anchor="${opts.verticalAlign}"`);
    }

    if (opts?.borders) {
      if (opts.borders.left) parts.push(buildBorderLine("a:lnL", opts.borders.left));
      if (opts.borders.right) parts.push(buildBorderLine("a:lnR", opts.borders.right));
      if (opts.borders.top) parts.push(buildBorderLine("a:lnT", opts.borders.top));
      if (opts.borders.bottom) parts.push(buildBorderLine("a:lnB", opts.borders.bottom));
    }

    if (opts?.fill !== undefined) {
      const fillComp = buildFill(opts.fill);
      parts.push(fillComp.toXml(context));
    }

    if (parts.length === 0) return "<a:tcPr/>";

    // Check if first part is an attribute string (no <)
    const firstIsAttr = !parts[0].startsWith("<");
    if (firstIsAttr) {
      const attrStr = parts[0];
      const children = parts.slice(1);
      if (children.length === 0) return `<a:tcPr ${attrStr}/>`;
      return `<a:tcPr ${attrStr}>${children.join("")}</a:tcPr>`;
    }

    return `<a:tcPr>${parts.join("")}</a:tcPr>`;
  }
}
