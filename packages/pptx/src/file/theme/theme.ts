import { ImportedXmlComponent } from "@file/xml-components";

export interface ColorSchemeOptions {
  dark1?: string;
  light1?: string;
  dark2?: string;
  light2?: string;
  accent1?: string;
  accent2?: string;
  accent3?: string;
  accent4?: string;
  accent5?: string;
  accent6?: string;
  hyperlink?: string;
  followedHyperlink?: string;
}

export interface FontSchemeOptions {
  majorFont?: string;
  minorFont?: string;
  majorFontAsian?: string;
  minorFontAsian?: string;
}

export interface ThemeOptions {
  name?: string;
  colors?: ColorSchemeOptions;
  fonts?: FontSchemeOptions;
}

const DEFAULT_COLORS: Required<ColorSchemeOptions> = {
  dark1: "000000",
  light1: "FFFFFF",
  dark2: "44546A",
  light2: "E7E6E6",
  accent1: "4472C4",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "5B9BD5",
  accent6: "70AD47",
  hyperlink: "0563C1",
  followedHyperlink: "954F72",
};

function buildThemeXml(options?: ThemeOptions): string {
  const name = options?.name ?? "Office Theme";
  const c = { ...DEFAULT_COLORS, ...options?.colors };
  const f = options?.fonts;

  const majorFont = f?.majorFont ?? "Calibri Light";
  const minorFont = f?.minorFont ?? "Calibri";
  const majorFontAsian = f?.majorFontAsian ?? "";
  const minorFontAsian = f?.minorFontAsian ?? "";

  // dk1/lt1 use sysClr in Office theme; srgbClr override when customized
  const dk1 =
    options?.colors?.dark1 !== undefined
      ? `<a:srgbClr val="${c.dark1}"/>`
      : `<a:sysClr val="windowText" lastClr="000000"/>`;
  const lt1 =
    options?.colors?.light1 !== undefined
      ? `<a:srgbClr val="${c.light1}"/>`
      : `<a:sysClr val="window" lastClr="FFFFFF"/>`;

  return `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${name}">
  <a:themeElements>
    <a:clrScheme name="${name}">
      <a:dk1>${dk1}</a:dk1>
      <a:lt1>${lt1}</a:lt1>
      <a:dk2><a:srgbClr val="${c.dark2}"/></a:dk2>
      <a:lt2><a:srgbClr val="${c.light2}"/></a:lt2>
      <a:accent1><a:srgbClr val="${c.accent1}"/></a:accent1>
      <a:accent2><a:srgbClr val="${c.accent2}"/></a:accent2>
      <a:accent3><a:srgbClr val="${c.accent3}"/></a:accent3>
      <a:accent4><a:srgbClr val="${c.accent4}"/></a:accent4>
      <a:accent5><a:srgbClr val="${c.accent5}"/></a:accent5>
      <a:accent6><a:srgbClr val="${c.accent6}"/></a:accent6>
      <a:hlink><a:srgbClr val="${c.hyperlink}"/></a:hlink>
      <a:folHlink><a:srgbClr val="${c.followedHyperlink}"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="${name}">
      <a:majorFont>
        <a:latin typeface="${majorFont}"/>
        <a:ea typeface="${majorFontAsian}"/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="${minorFont}"/>
        <a:ea typeface="${minorFontAsian}"/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`;
}

export class DefaultTheme extends ImportedXmlComponent {
  private static cache = new Map<string, ImportedXmlComponent>();
  private readonly cacheKey: string;

  public constructor(options?: ThemeOptions) {
    super("a:theme");
    this.cacheKey = options ? JSON.stringify(options) : "";
    if (!DefaultTheme.cache.has(this.cacheKey)) {
      DefaultTheme.cache.set(
        this.cacheKey,
        ImportedXmlComponent.fromXmlString(buildThemeXml(options)),
      );
    }
  }

  public prepForXml() {
    return DefaultTheme.cache.get(this.cacheKey)!.prepForXml({ stack: [] });
  }
}
