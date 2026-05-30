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
import { xsdStrikeStyle, xsdTextCaps, xsdUnderlineStyle } from "@office-open/core";
import { createEffectList, createOutline } from "@office-open/core/drawingml";
import { xml } from "@office-open/xml";

let nextHyperlinkId = 1;

export const UnderlineStyle = {
  SINGLE: "single",
  DOUBLE: "double",
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
  readonly underline?: (typeof UnderlineStyle)[keyof typeof UnderlineStyle];
  readonly font?: string;
  readonly lang?: string;
  readonly fill?: FillOptions;
  readonly hyperlink?: HyperlinkOptions;
  readonly strike?: (typeof StrikeStyle)[keyof typeof StrikeStyle];
  readonly baseline?: number;
  readonly spacing?: number;
  readonly capitalization?: (typeof TextCapitalization)[keyof typeof TextCapitalization];
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
  if (options.underline) attrs.u = xsdUnderlineStyle.to(options.underline);
  if (options.lang) attrs.lang = options.lang;
  if (options.strike) attrs.strike = xsdStrikeStyle.to(options.strike);
  if (options.baseline !== undefined) attrs.baseline = options.baseline;
  if (options.capitalization) attrs.cap = xsdTextCaps.to(options.capitalization);
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

  public override toXml(context: Context): string {
    const opts = this.options;

    // Side-effect: register hyperlink
    let hyperlinkKey: string | undefined;
    if (opts.hyperlink) {
      hyperlinkKey = `hlink_${nextHyperlinkId++}`;
      const file = context.fileData as {
        hyperlinks?: { addHyperlink(key: string, url: string, tooltip?: string): void };
      };
      file?.hyperlinks?.addHyperlink(hyperlinkKey, opts.hyperlink.url, opts.hyperlink.tooltip);
    }

    // fill / outline / shadow still need prepForXml (child components have own deps)
    let fillObj: IXmlableObject | undefined;
    if (opts.fill !== undefined) {
      fillObj = buildFill(opts.fill).prepForXml(context) ?? undefined;
    }
    let outlineObj: IXmlableObject | undefined;
    if (opts.outline) {
      outlineObj =
        createOutline({
          width: DEFAULT_OUTLINE_WIDTH,
          type: "solidFill",
          color: { value: "000000" },
        }).prepForXml(context) ?? undefined;
    }
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

    const obj = buildRunProperties(opts, hyperlinkKey, fillObj, effectListObj, outlineObj);
    return obj ? xml(obj) : "";
  }
}
