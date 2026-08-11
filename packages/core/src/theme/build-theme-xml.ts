/**
 * Build theme XML string — structured driver over the CT_BaseStyles parts.
 *
 * Fresh output stays byte-identical to Office by reusing a pre-extracted
 * default fmtScheme fragment (O(1) fast path) unless the caller supplies a
 * structured formatScheme / objectDefaults / extraColorSchemes, in which case
 * the relevant descriptor serializes them.
 *
 * @module
 */
import type { WriteContext } from "../descriptor";
import { stringifyColorScheme } from "./color-scheme";
import { stringifyExtraColorSchemes } from "./extra-color-scheme";
import { stringifyFontScheme } from "./font-scheme";
import { stringifyObjectDefaults } from "./object-defaults";
import { stringifyFormatScheme } from "./style-matrix";
import type { ThemeOptions } from "./theme-options";

// Default Office fmtScheme — byte-identical to fresh output, keeps the common
// path O(1) without running fill/line/effect descriptors.
const DEFAULT_FORMAT_SCHEME_XML =
  '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/>' +
  '<a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr">' +
  '<a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>' +
  '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/>' +
  '</a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1">' +
  '<a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="102000"/><a:satMod val="103000"/><a:tint val="94000"/>' +
  '</a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="100000"/><a:satMod val="110000"/>' +
  '<a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/>' +
  '<a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/>' +
  '</a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>' +
  '<a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
  '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
  "</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/>" +
  '</a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" ' +
  'rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst>' +
  '</a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>' +
  '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="102000"/>' +
  '<a:satMod val="150000"/><a:tint val="93000"/><a:shade val="98000"/></a:schemeClr></a:gs><a:gs pos="50000">' +
  '<a:schemeClr val="phClr"><a:lumMod val="103000"/><a:satMod val="130000"/><a:tint val="98000"/><a:shade val="90000"/>' +
  '</a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:satMod val="120000"/><a:shade val="63000"/>' +
  '</a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>';

function requireCtx(ctx: WriteContext | undefined, field: string): WriteContext {
  if (!ctx) throw new Error(`buildThemeXml: ${field} requires a WriteContext`);
  return ctx;
}

export function buildThemeXml(options?: ThemeOptions, ctx?: WriteContext): string {
  const opts = options ?? {};
  const name = opts.name ?? "Office Theme";

  const clrScheme = stringifyColorScheme(opts.colorScheme, name);
  const fontScheme = stringifyFontScheme(opts.fontScheme, name);
  const fmtScheme = opts.formatScheme
    ? stringifyFormatScheme(opts.formatScheme, requireCtx(ctx, "formatScheme"))
    : DEFAULT_FORMAT_SCHEME_XML;
  const objectDefaults = opts.objectDefaults
    ? stringifyObjectDefaults(opts.objectDefaults, requireCtx(ctx, "objectDefaults"))
    : "<a:objectDefaults/>";
  const extraClrSchemeLst = stringifyExtraColorSchemes(opts.extraColorSchemes, name);

  return (
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${name}">` +
    `<a:themeElements>${clrScheme}${fontScheme}${fmtScheme}</a:themeElements>` +
    `${objectDefaults}${extraClrSchemeLst}</a:theme>`
  );
}
