import { parseOnOff, stringifyColorMapping } from "@office-open/core";
import type { ColorMappingOptions } from "@office-open/core";
import { attr, type Element } from "@office-open/xml";
import { SP_TREE_HEADER } from "@shared/constants";

/** Header/footer options for handout/notes master */
export interface HeaderFooterOptions {
  /** Show date/time */
  date?: boolean;
  /** Show header */
  header?: boolean;
  /** Show footer */
  footer?: boolean;
  /** Show slide number */
  slideNumber?: boolean;
}

/** Options for handout master parameterization */
export interface HandoutMasterOptions {
  /** Color mapping overrides */
  colorMapping?: Partial<ColorMappingOptions>;
  /** Header/footer settings */
  headerFooter?: HeaderFooterOptions;
}

export function buildHfAttrs(opts?: HeaderFooterOptions): string {
  if (!opts) return 'dt="0" hdr="0" ftr="0" sldNum="0"';
  return `dt="${opts.date ? 1 : 0}" hdr="${opts.header ? 1 : 0}" ftr="${opts.footer ? 1 : 0}" sldNum="${opts.slideNumber ? 1 : 0}"`;
}

/** Parse a p:hf element into HeaderFooterOptions (undefined when empty/absent). */
export function parseHeaderFooter(el: Element | undefined): HeaderFooterOptions | undefined {
  if (!el) return undefined;
  const headerFooter: HeaderFooterOptions = {};
  const dt = attr(el, "dt");
  if (dt !== undefined) headerFooter.date = parseOnOff(dt) ?? false;
  const hdr = attr(el, "hdr");
  if (hdr !== undefined) headerFooter.header = parseOnOff(hdr) ?? false;
  const ftr = attr(el, "ftr");
  if (ftr !== undefined) headerFooter.footer = parseOnOff(ftr) ?? false;
  const sldNum = attr(el, "sldNum");
  if (sldNum !== undefined) headerFooter.slideNumber = parseOnOff(sldNum) ?? false;
  return Object.keys(headerFooter).length > 0 ? headerFooter : undefined;
}

export function buildHandoutMasterXml(options?: HandoutMasterOptions): string {
  const colorMapping = stringifyColorMapping(options?.colorMapping, "p:clrMap");
  const hf = buildHfAttrs(options?.headerFooter);
  return (
    '<p:handoutMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
    `<p:spTree>${SP_TREE_HEADER}</p:spTree></p:cSld>` +
    colorMapping +
    `<p:hf ${hf}/>` +
    "</p:handoutMaster>"
  );
}
