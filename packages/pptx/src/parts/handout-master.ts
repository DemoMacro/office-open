import { parseOnOff, stringifyColorMapping } from "@office-open/core";
import type { ColorMappingOptions } from "@office-open/core";
import type { BlackWhiteMode } from "@office-open/core/drawing";
import type { ThemeOptions } from "@office-open/core/theme";
import { attr, type Element } from "@office-open/xml";
import type { BackgroundOptions } from "@parts/background";
import type { SlideChild } from "@parts/slide/slide-child";
import { stringifySpTreeHeader } from "@shared/constants";

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
  /** Background (p:bg); defaults to the Office bgRef idx="1001". */
  background?: BackgroundOptions;
  /** Custom spTree shapes (p:spTree children after the group header). */
  children?: SlideChild[];
  /** Root spTree black-and-white mode (p:grpSpPr @bwMode). */
  shapeTreeBlackWhiteMode?: BlackWhiteMode;
  /** Color mapping overrides */
  colorMapping?: Partial<ColorMappingOptions>;
  /** Header/footer settings */
  headerFooter?: HeaderFooterOptions;
  /** The handout master's own theme part; defaults to the Office theme. */
  theme?: ThemeOptions;
  /** Trailing p:extLst inner XML (e.g. p14:creationId) — verbatim round-trip. */
  ext?: string;
}

export function buildHfAttrs(opts?: HeaderFooterOptions): string {
  if (!opts) return "";
  const attrs: string[] = [];
  if (opts.date !== undefined) attrs.push(`dt="${opts.date ? 1 : 0}"`);
  if (opts.header !== undefined) attrs.push(`hdr="${opts.header ? 1 : 0}"`);
  if (opts.footer !== undefined) attrs.push(`ftr="${opts.footer ? 1 : 0}"`);
  if (opts.slideNumber !== undefined) attrs.push(`sldNum="${opts.slideNumber ? 1 : 0}"`);
  return attrs.join(" ");
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

/**
 * Build the handout master XML (fresh generation only).
 *
 * The handout master stays a fixed template: its spTree holds only print-layout
 * placeholders that PowerPoint regenerates on demand, real files carry no
 * authorable content worth round-tripping, and nothing in the parse pipeline
 * consumes it — structuring it would add surface with no consumer.
 */
export function buildHandoutMasterXml(options?: HandoutMasterOptions): string {
  const colorMapping = stringifyColorMapping(options?.colorMapping, "p:clrMap");
  const hfAttrs = buildHfAttrs(
    options?.headerFooter ?? { date: false, header: false, footer: false, slideNumber: false },
  );
  return (
    '<p:handoutMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
    `<p:spTree>${stringifySpTreeHeader(1, "", options?.shapeTreeBlackWhiteMode)}</p:spTree></p:cSld>` +
    colorMapping +
    `<p:hf${hfAttrs ? ` ${hfAttrs}` : ""}/>` +
    "</p:handoutMaster>"
  );
}
