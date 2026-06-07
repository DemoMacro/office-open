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
import type { Context } from "@file/xml-components";
import { xsdStrikeStyle, xsdTextCaps, xsdUnderlineStyle } from "@office-open/core";
import { createEffectList, createOutline } from "@office-open/core/drawingml";

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
  url: string;
  tooltip?: string;
  action?: string;
  highlightClick?: boolean;
  endSound?: boolean;
  invalidUrl?: boolean;
}

export interface RunPropertiesOptions {
  /** Font size in points. Serialized as OOXML `a:sz` (hundredths of a point). */
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: (typeof UnderlineStyle)[keyof typeof UnderlineStyle];
  font?: string;
  lang?: string;
  fill?: FillOptions;
  hyperlink?: HyperlinkOptions;
  strike?: (typeof StrikeStyle)[keyof typeof StrikeStyle];
  baseline?: number;
  spacing?: number;
  capitalization?: (typeof TextCapitalization)[keyof typeof TextCapitalization];
  shadow?: boolean;
  outline?: boolean;
  rightToLeft?: boolean;
  noProof?: boolean;
  dirty?: boolean;
  /** East Asian line break. When true, can break East Asian words. */
  kumimoji?: boolean;
  /** Alternate language for the text run. */
  alternateLanguage?: string;
  /** Normalize height. When true, normalize the height of text. */
  normalizeHeight?: boolean;
  /** Bookmark mark identifier. */
  bookmarkMark?: string;
  /** Smart tag ID. */
  smartTagId?: string;
}

/**
 * Builds a:rPr attribute string.
 */
function buildAttrString(options: RunPropertiesOptions): string {
  const attrs: string[] = [];
  if (options.fontSize) attrs.push(`sz="${options.fontSize * 100}"`);
  if (options.bold !== undefined) attrs.push(`b="${options.bold ? 1 : 0}"`);
  if (options.italic !== undefined) attrs.push(`i="${options.italic ? 1 : 0}"`);
  if (options.underline) attrs.push(`u="${xsdUnderlineStyle.to(options.underline)}"`);
  if (options.lang) attrs.push(`lang="${options.lang}"`);
  if (options.strike) attrs.push(`strike="${xsdStrikeStyle.to(options.strike)}"`);
  if (options.baseline !== undefined) attrs.push(`baseline="${options.baseline}"`);
  if (options.capitalization) attrs.push(`cap="${xsdTextCaps.to(options.capitalization)}"`);
  if (options.spacing !== undefined) attrs.push(`spc="${options.spacing}"`);
  if (options.noProof !== undefined) attrs.push(`noProof="${options.noProof ? 1 : 0}"`);
  if (options.dirty !== undefined) attrs.push(`dirty="${options.dirty ? 1 : 0}"`);
  if (options.kumimoji !== undefined) attrs.push(`kumimoji="${options.kumimoji ? 1 : 0}"`);
  if (options.alternateLanguage) attrs.push(`altLang="${options.alternateLanguage}"`);
  if (options.normalizeHeight !== undefined)
    attrs.push(`normalizeH="${options.normalizeHeight ? 1 : 0}"`);
  if (options.bookmarkMark) attrs.push(`bmk="${options.bookmarkMark}"`);
  if (options.smartTagId) attrs.push(`smtId="${options.smartTagId}"`);
  return attrs.join(" ");
}

/**
 * a:rPr — Run properties (font, size, color, etc.).
 */
export class RunProperties extends XmlComponent {
  private options: RunPropertiesOptions;

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

    const parts: string[] = [];
    const attrStr = buildAttrString(opts);

    // XSD order: ln → fill → effect → latin/ea → hlinkClick → rtl
    if (opts.outline) {
      parts.push(
        createOutline({
          width: DEFAULT_OUTLINE_WIDTH,
          type: "solidFill",
          color: { value: "000000" },
        }).toXml(context),
      );
    }
    if (opts.fill !== undefined) {
      parts.push(buildFill(opts.fill).toXml(context));
    }
    if (opts.shadow) {
      parts.push(
        createEffectList({
          outerShadow: {
            blurRadius: DEFAULT_SHADOW_BLUR_RADIUS,
            distance: DEFAULT_SHADOW_DISTANCE,
            direction: DEFAULT_SHADOW_DIRECTION,
            color: { value: "000000", transforms: { alpha: DEFAULT_SHADOW_ALPHA } },
          },
        }).toXml(context),
      );
    }

    if (opts.font) {
      parts.push(`<a:latin typeface="${opts.font}"/>`);
      parts.push(`<a:ea typeface="${opts.font}"/>`);
    }

    if (opts.hyperlink && hyperlinkKey) {
      const hl = opts.hyperlink;
      const hlAttrs: string[] = [`r:id="{hlink:${hyperlinkKey}}"`];
      if (hl.tooltip) hlAttrs.push(`tooltip="${hl.tooltip}"`);
      if (hl.action) hlAttrs.push(`action="${hl.action}"`);
      if (hl.highlightClick) hlAttrs.push('highlightClick="1"');
      if (hl.endSound) hlAttrs.push('endSnd="1"');
      if (hl.invalidUrl) hlAttrs.push('invalidUrl="1"');
      parts.push(`<a:hlinkClick ${hlAttrs.join(" ")}/>`);
    }

    if (opts.rightToLeft !== undefined) {
      parts.push(`<a:rtl val="${opts.rightToLeft ? 1 : 0}"/>`);
    }

    if (attrStr.length === 0 && parts.length === 0) return "";

    if (parts.length === 0) return `<a:rPr ${attrStr}/>`;
    return `<a:rPr ${attrStr}>${parts.join("")}</a:rPr>`;
  }
}
