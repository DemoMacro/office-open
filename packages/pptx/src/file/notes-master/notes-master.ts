import type { Context } from "@file/xml-components";
import { ImportedXmlComponent } from "@file/xml-components";

import type { ColorMapOptions, HeaderFooterOptions } from "../handout-master/handout-master";

export type { ColorMapOptions, HeaderFooterOptions };

/** Notes style level override */
export interface NotesLevelProperties {
  /** Font size in hundredths of a point (e.g., 1200 = 12pt) */
  readonly fontSize?: number;
  /** Left margin in EMU */
  readonly marginLeft?: number;
  /** Alignment ("l" | "ctr" | "r" | "just") */
  readonly alignment?: string;
}

/** Options for notes master parameterization */
export interface NotesMasterOptions {
  /** Color map overrides */
  readonly colorMap?: ColorMapOptions;
  /** Header/footer settings */
  readonly headerFooter?: HeaderFooterOptions;
  /** Notes style overrides (levels 1-9) */
  readonly notesStyle?: readonly NotesLevelProperties[];
}

const DEFAULT_COLOR_MAP =
  'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';

const NOTES_MASTER_XML = `<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001">
        <a:schemeClr val="bg1"/>
      </p:bgRef>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:hf dt="0" hdr="0" ftr="0" sldNum="0"/>
  <p:notesStyle>
    <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
      <a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
    </a:lvl1pPr>
    <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
      <a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
    </a:lvl2pPr>
    <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
      <a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
    </a:lvl3pPr>
    <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
      <a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
    </a:lvl4pPr>
    <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
      <a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
    </a:lvl5pPr>
    <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
      <a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
    </a:lvl6pPr>
    <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
      <a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
    </a:lvl7pPr>
    <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
      <a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
    </a:lvl8pPr>
    <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
      <a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
    </a:lvl9pPr>
  </p:notesStyle>
</p:notesMaster>`;

const DEFAULT_LEVEL_MARGINS = [
  0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600,
];

function buildColorMapAttrs(opts?: ColorMapOptions): string {
  if (!opts) return DEFAULT_COLOR_MAP;
  const defaults: Required<ColorMapOptions> = {
    bg1: "lt1",
    tx1: "dk1",
    bg2: "lt2",
    tx2: "dk2",
    accent1: "accent1",
    accent2: "accent2",
    accent3: "accent3",
    accent4: "accent4",
    accent5: "accent5",
    accent6: "accent6",
    hlink: "hlink",
    folHlink: "folHlink",
  };
  const merged = { ...defaults, ...opts };
  return Object.entries(merged)
    .map(([k, v]) => `${k}="${v}"`)
    .join(" ");
}

function buildHfAttrs(opts?: HeaderFooterOptions): string {
  if (!opts) return 'dt="0" hdr="0" ftr="0" sldNum="0"';
  return `dt="${opts.date ? 1 : 0}" hdr="${opts.header ? 1 : 0}" ftr="${opts.footer ? 1 : 0}" sldNum="${opts.slideNumber ? 1 : 0}"`;
}

function buildNotesStyleXml(levels?: readonly NotesLevelProperties[]): string {
  const parts: string[] = ["<p:notesStyle>"];
  for (let i = 0; i < 9; i++) {
    const level = levels?.[i];
    const marL = level?.marginLeft ?? DEFAULT_LEVEL_MARGINS[i];
    const algn = level?.alignment ?? "l";
    const sz = level?.fontSize ?? 1200;
    parts.push(
      `<a:lvl${i + 1}pPr marL="${marL}" algn="${algn}" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">` +
        `<a:defRPr sz="${sz}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>` +
        `<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`,
    );
  }
  parts.push("</p:notesStyle>");
  return parts.join("");
}

export class DefaultNotesMaster extends ImportedXmlComponent {
  private static instance = ImportedXmlComponent.fromXmlString(NOTES_MASTER_XML);
  private readonly options?: NotesMasterOptions;

  public constructor(options?: NotesMasterOptions) {
    super("p:notesMaster");
    this.options = options;
  }

  public override toXml(context: Context): string {
    // No options → use fast static XML
    if (!this.options) {
      return DefaultNotesMaster.instance.toXml(context);
    }

    // Build dynamic XML with custom options
    const colorMap = buildColorMapAttrs(this.options.colorMap);
    const hf = buildHfAttrs(this.options.headerFooter);
    const notesStyle = buildNotesStyleXml(this.options.notesStyle);

    return (
      '<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
      '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>' +
      `<p:clrMap ${colorMap}/>` +
      `<p:hf ${hf}/>` +
      notesStyle +
      "</p:notesMaster>"
    );
  }
}
