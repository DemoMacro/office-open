import {
  DEFAULT_OUTLINE_WIDTH,
  DEFAULT_SHADOW_ALPHA,
  DEFAULT_SHADOW_BLUR_RADIUS,
  DEFAULT_SHADOW_DIRECTION,
  DEFAULT_SHADOW_DISTANCE,
} from "@file/constants";
import type { FillOptions } from "@file/drawingml/fill";
import { buildFill } from "@file/drawingml/fill";
import { XmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";
import { createEffectList, createOutline } from "@office-open/core/drawingml";

let nextHyperlinkId = 1;

export const UnderlineStyle = {
  SINGLE: "sng",
  DOUBLE: "dbl",
  NONE: "none",
} as const;

export const StrikeStyle = {
  SINGLE: "sngStrike",
  DOUBLE: "dblStrike",
  NONE: "noStrike",
} as const;

export const TextCapitalization = {
  NONE: "none",
  ALL: "all",
  SMALL: "small",
} as const;

export interface HyperlinkOptions {
  readonly url: string;
  readonly tooltip?: string;
}

export interface RunPropertiesOptions {
  /** Font size in points. Serialized as OOXML `a:sz` (hundredths of a point). */
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: keyof typeof UnderlineStyle;
  readonly font?: string;
  readonly lang?: string;
  readonly fill?: FillOptions;
  readonly hyperlink?: HyperlinkOptions;
  readonly strike?: keyof typeof StrikeStyle;
  readonly baseline?: number;
  readonly spacing?: number;
  readonly capitalization?: keyof typeof TextCapitalization;
  readonly shadow?: boolean;
  readonly outline?: boolean;
  readonly rightToLeft?: boolean;
  readonly noProof?: boolean;
  readonly dirty?: boolean;
}

/**
 * Pure function: builds a:rPr XML object from options.
 * @param hyperlinkKey - pre-generated key for hyperlink placeholder
 * @param fillObject - pre-built fill IXmlableObject (from buildFill)
 */
export function buildRunProperties(
  options: RunPropertiesOptions,
  hyperlinkKey?: string,
  fillObject?: IXmlableObject,
  effectListObject?: IXmlableObject,
  outlineObject?: IXmlableObject,
): IXmlableObject | undefined {
  const children: IXmlableObject[] = [];

  const attrs: Record<string, string | number | boolean> = {};
  if (options.fontSize) attrs.sz = options.fontSize * 100;
  if (options.bold !== undefined) attrs.b = options.bold;
  if (options.italic !== undefined) attrs.i = options.italic;
  if (options.underline) attrs.u = UnderlineStyle[options.underline];
  if (options.lang) attrs.lang = options.lang;
  if (options.strike) attrs.strike = StrikeStyle[options.strike];
  if (options.baseline !== undefined) attrs.baseline = options.baseline;
  if (options.capitalization)
    attrs.cap = TextCapitalization[options.capitalization] ?? options.capitalization;
  if (options.spacing !== undefined) attrs.spc = options.spacing;
  if (options.noProof !== undefined) attrs.noProof = options.noProof;
  if (options.dirty !== undefined) attrs.dirty = options.dirty;
  if (Object.keys(attrs).length > 0) children.push({ _attr: attrs });

  // XSD order: ln → fill → effect → latin/ea → hlinkClick → rtl
  if (outlineObject) {
    children.push(outlineObject);
  }

  if (fillObject) {
    children.push(fillObject);
  }

  if (effectListObject) {
    children.push(effectListObject);
  }

  if (options.font) {
    children.push({ "a:latin": { _attr: { typeface: options.font } } });
    children.push({ "a:ea": { _attr: { typeface: options.font } } });
  }

  if (options.hyperlink && hyperlinkKey) {
    const hlinkAttrs: Record<string, string> = { "r:id": `{hlink:${hyperlinkKey}}` };
    if (options.hyperlink.tooltip) hlinkAttrs.tooltip = options.hyperlink.tooltip;
    children.push({ "a:hlinkClick": { _attr: hlinkAttrs } });
  }

  if (options.rightToLeft !== undefined) {
    children.push({ "a:rtl": { _attr: { val: options.rightToLeft ? 1 : 0 } } });
  }

  if (children.length === 0) return undefined;

  return {
    "a:rPr": children.length === 1 && "_attr" in children[0] ? children[0] : children,
  };
}

/**
 * a:rPr — Run properties (font, size, color, etc.).
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class RunProperties extends XmlComponent {
  private readonly options: RunPropertiesOptions;

  public static hasProperties(options: RunPropertiesOptions): boolean {
    return !!(
      options.fontSize ||
      options.bold !== undefined ||
      options.italic !== undefined ||
      options.underline ||
      options.font ||
      options.lang ||
      options.fill ||
      options.hyperlink ||
      options.strike ||
      options.baseline !== undefined ||
      options.spacing !== undefined ||
      options.capitalization ||
      options.shadow !== undefined ||
      options.outline !== undefined ||
      options.rightToLeft !== undefined ||
      options.noProof !== undefined ||
      options.dirty !== undefined
    );
  }

  public constructor(options: RunPropertiesOptions = {}) {
    super("a:rPr");
    this.options = options;
  }

  public prepForXml(context: Context): IXmlableObject | undefined {
    const opts = this.options;

    // Register hyperlinks (B-level: side effect on context)
    let hyperlinkKey: string | undefined;
    if (opts.hyperlink) {
      hyperlinkKey = `hlink_${nextHyperlinkId++}`;
      const file = context.fileData as {
        Hyperlinks?: { addHyperlink(key: string, url: string, tooltip?: string): void };
      };
      file?.Hyperlinks?.addHyperlink(hyperlinkKey, opts.hyperlink.url, opts.hyperlink.tooltip);
    }

    // Handle fill (needs context for prepForXml)
    let fillObj: IXmlableObject | undefined;
    if (opts.fill !== undefined) {
      fillObj = buildFill(opts.fill).prepForXml(context) ?? undefined;
    }

    // Handle outline
    let outlineObj: IXmlableObject | undefined;
    if (opts.outline) {
      outlineObj =
        createOutline({
          width: DEFAULT_OUTLINE_WIDTH,
          type: "solidFill",
          color: { value: "000000" },
        }).prepForXml(context) ?? undefined;
    }

    // Handle shadow
    let effectListObj: IXmlableObject | undefined;
    if (opts.shadow) {
      effectListObj =
        createEffectList({
          outerShadow: {
            blurRadius: DEFAULT_SHADOW_BLUR_RADIUS,
            distance: DEFAULT_SHADOW_DISTANCE,
            direction: DEFAULT_SHADOW_DIRECTION,
            color: { value: "000000", transforms: { alpha: DEFAULT_SHADOW_ALPHA } },
          },
        }).prepForXml(context) ?? undefined;
    }

    return buildRunProperties(opts, hyperlinkKey, fillObj, effectListObj, outlineObj);
  }
}
