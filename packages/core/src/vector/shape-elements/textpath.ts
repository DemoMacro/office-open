/**
 * v:textpath element — CT_TextPath (WordArt text on a path).
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd, CT_TextPath.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import {
  stringifyVmlAttributes,
  parseVmlAttributes,
  type VmlAttrSpec,
  type VmlTrueFalse,
} from "../attributes";

/**
 * v:textpath style vocabulary — font-centric CSS, distinct from the layout
 * vocabulary of `VmlShapeStyle`. Values stay strings: font sizes carry
 * units ("24pt"), families carry quotes, weights may be named or numeric.
 */
export interface VmlTextPathStyle {
  fontFamily?: string;
  fontSize?: string;
  fontWeight?: string;
  fontStyle?: string;
  fontVariant?: string;
  textTransform?: string;
  textDecoration?: string;
  letterSpacing?: string;
  /** v-text-align — WordArt alignment ("center", "letter-justify", …). */
  vTextAlign?: string;
  /** v-text-kern — kerning toggle ("t"/"f" literal in the CSS stream). */
  vTextKern?: string;
  /** v-text-spacing — character spacing. */
  vTextSpacing?: string;
  /** v-text-spacing-mode — spacing mode ("tightening", …). */
  vTextSpacingMode?: string;
}

const TEXT_PATH_STYLE_MAP: Record<keyof VmlTextPathStyle, string> = {
  fontFamily: "font-family",
  fontSize: "font-size",
  fontWeight: "font-weight",
  fontStyle: "font-style",
  fontVariant: "font-variant",
  textTransform: "text-transform",
  textDecoration: "text-decoration",
  letterSpacing: "letter-spacing",
  vTextAlign: "v-text-align",
  vTextKern: "v-text-kern",
  vTextSpacing: "v-text-spacing",
  vTextSpacingMode: "v-text-spacing-mode",
};

const KEY_TO_TEXT_PATH_FIELD = Object.fromEntries(
  Object.entries(TEXT_PATH_STYLE_MAP).map(([field, key]) => [key, field]),
) as Record<string, keyof VmlTextPathStyle>;

/** Serialize a VmlTextPathStyle to its CSS-stream form. */
function stringifyVmlTextPathStyle(style: VmlTextPathStyle): string {
  return Object.entries(style)
    .map(([key, value]) => `${TEXT_PATH_STYLE_MAP[key as keyof VmlTextPathStyle]}:${value}`)
    .join(";");
}

/** Parse a CSS stream into a VmlTextPathStyle; unmapped keys drop. */
function parseVmlTextPathStyle(styleStr: string): VmlTextPathStyle {
  const style: Partial<VmlTextPathStyle> = {};
  for (const part of styleStr.split(";")) {
    const [key, val] = part.split(":").map((s) => s.trim());
    if (!key || !val) continue;
    const field = KEY_TO_TEXT_PATH_FIELD[key];
    if (field) style[field] = val;
  }
  return style as VmlTextPathStyle;
}

/** v:textpath options (CT_TextPath). */
export interface VmlTextPathOptions {
  id?: string;
  style?: VmlTextPathStyle;
  on?: VmlTrueFalse;
  fitshape?: VmlTrueFalse;
  fitpath?: VmlTrueFalse;
  trim?: VmlTrueFalse;
  xscale?: VmlTrueFalse;
  string?: string;
}

const TEXT_PATH_ATTRS: readonly VmlAttrSpec[] = [
  { field: "id", attr: "id", kind: "string" },
  { field: "on", attr: "on", kind: "trueFalse" },
  { field: "fitshape", attr: "fitshape", kind: "trueFalse" },
  { field: "fitpath", attr: "fitpath", kind: "trueFalse" },
  { field: "trim", attr: "trim", kind: "trueFalse" },
  { field: "xscale", attr: "xscale", kind: "trueFalse" },
  { field: "string", attr: "string", kind: "string" },
];

/** Serialize v:textpath. */
export function stringifyVmlTextPath(opts: VmlTextPathOptions): string {
  const attrStr = stringifyVmlAttributes(
    opts as unknown as Record<string, unknown>,
    TEXT_PATH_ATTRS,
  );
  if (opts.style !== undefined) {
    return `<v:textpath${attrStr} style="${escapeXml(stringifyVmlTextPathStyle(opts.style))}"/>`;
  }
  return `<v:textpath${attrStr}/>`;
}

/** Parse a v:textpath element. */
export function parseVmlTextPath(el: XmlElement): VmlTextPathOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, TEXT_PATH_ATTRS, out);
  if (el.attributes?.style !== undefined) {
    out.style = parseVmlTextPathStyle(String(el.attributes.style));
  }
  return out as VmlTextPathOptions;
}
