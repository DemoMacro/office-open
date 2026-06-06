import type { Context } from "@file/xml-components";
import { ImportedXmlComponent } from "@file/xml-components";

import { DEFAULT_COLOR_MAP, SP_TREE_HEADER } from "../constants";

/** Color map entry for handout/notes master */
export interface ColorMapOptions {
  readonly bg1?: string;
  readonly tx1?: string;
  readonly bg2?: string;
  readonly tx2?: string;
  readonly accent1?: string;
  readonly accent2?: string;
  readonly accent3?: string;
  readonly accent4?: string;
  readonly accent5?: string;
  readonly accent6?: string;
  readonly hlink?: string;
  readonly folHlink?: string;
}

/** Header/footer options for handout/notes master */
export interface HeaderFooterOptions {
  /** Show date/time */
  readonly date?: boolean;
  /** Show header */
  readonly header?: boolean;
  /** Show footer */
  readonly footer?: boolean;
  /** Show slide number */
  readonly slideNumber?: boolean;
}

/** Options for handout master parameterization */
export interface HandoutMasterOptions {
  /** Color map overrides */
  readonly colorMap?: ColorMapOptions;
  /** Header/footer settings */
  readonly headerFooter?: HeaderFooterOptions;
}

const HANDOUT_MASTER_XML = `<p:handoutMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
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
</p:handoutMaster>`;

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

export class DefaultHandoutMaster extends ImportedXmlComponent {
  private static instance = ImportedXmlComponent.fromXmlString(HANDOUT_MASTER_XML);
  private readonly options?: HandoutMasterOptions;

  public constructor(options?: HandoutMasterOptions) {
    super("p:handoutMaster");
    this.options = options;
  }

  public override toXml(context: Context): string {
    // No options → use fast static XML
    if (!this.options) {
      return DefaultHandoutMaster.instance.toXml(context);
    }

    // Build dynamic XML with custom colorMap and headerFooter
    const colorMap = buildColorMapAttrs(this.options.colorMap);
    const hf = buildHfAttrs(this.options.headerFooter);

    return (
      '<p:handoutMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
      `<p:spTree>${SP_TREE_HEADER}</p:spTree></p:cSld>` +
      `<p:clrMap ${colorMap}/>` +
      `<p:hf ${hf}/>` +
      "</p:handoutMaster>"
    );
  }
}
