import { DEFAULT_COLOR_MAP, SP_TREE_HEADER } from "./constants";

/** Color map entry for handout/notes master */
export interface ColorMapOptions {
  bg1?: string;
  tx1?: string;
  bg2?: string;
  tx2?: string;
  accent1?: string;
  accent2?: string;
  accent3?: string;
  accent4?: string;
  accent5?: string;
  accent6?: string;
  hlink?: string;
  folHlink?: string;
}

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
  /** Color map overrides */
  colorMap?: ColorMapOptions;
  /** Header/footer settings */
  headerFooter?: HeaderFooterOptions;
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

export function buildHandoutMasterXml(options?: HandoutMasterOptions): string {
  if (!options) return HANDOUT_MASTER_XML;
  const colorMap = buildColorMapAttrs(options.colorMap);
  const hf = buildHfAttrs(options.headerFooter);
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
