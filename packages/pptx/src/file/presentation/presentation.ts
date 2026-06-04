import { XmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

export interface PhotoAlbumOptions {
  readonly blackWhite?: boolean;
  readonly showCaptions?: boolean;
  readonly layout?: "fitToSlide" | "1Photo" | "2Photo" | "4Photo";
  readonly frame?:
    | "frameStyle1"
    | "frameStyle2"
    | "frameStyle3"
    | "frameStyle4"
    | "frameStyle5"
    | "frameStyle6"
    | "none";
}

export interface ModifyVerifierOptions {
  readonly algorithmName?: string;
  readonly hashValue?: string;
  readonly saltValue?: string;
  readonly spinValue?: number;
  readonly cryptProviderType?: string;
  readonly cryptAlgorithmClass?: string;
  readonly cryptAlgorithmType?: string;
  readonly cryptAlgorithmSid?: number;
  readonly spinCount?: number;
  readonly saltData?: string;
  readonly hashData?: string;
  readonly cryptProvider?: string;
  readonly algorithmIdExtension?: number;
  readonly algorithmIdExtensionSource?: string;
  readonly cryptProviderTypeExtension?: number;
  readonly cryptProviderTypeExtensionSource?: string;
}

export interface PresentationOptions {
  readonly slideWidth?: number;
  readonly slideHeight?: number;
  readonly slideIds: readonly number[];
  readonly masterCount: number;
  readonly notesMasterRId?: number;
  readonly handoutMasterRId?: number;
  readonly serverZoom?: string;
  readonly firstSlideNum?: number;
  readonly showSpecialPlsOnTitleSld?: boolean;
  readonly rtl?: boolean;
  readonly removePersonalInfoOnSave?: boolean;
  readonly compatMode?: boolean;
  readonly strictFirstAndLastChars?: boolean;
  readonly embedTrueTypeFonts?: boolean;
  readonly saveSubsetFonts?: boolean;
  readonly autoCompressPictures?: boolean;
  readonly bookmarkIdSeed?: number;
  readonly conformance?: "strict" | "transitional";
  readonly photoAlbum?: PhotoAlbumOptions;
  readonly modifyVerifier?: ModifyVerifierOptions;
}

const DEFAULT_TEXT_STYLE_XML = `<p:defaultTextStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:defPPr><a:defRPr/></a:defPPr>
  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl1pPr>
  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl2pPr>
  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl3pPr>
  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl4pPr>
  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl5pPr>
  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl6pPr>
  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl7pPr>
  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl8pPr>
  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl9pPr>
</p:defaultTextStyle>`;

/**
 * p:presentation — Root element of a PPTX file.
 * Lazy: stores options, builds XML in toXml().
 */
export class Presentation extends XmlComponent {
  private options: PresentationOptions;

  public constructor(options: PresentationOptions) {
    super("p:presentation");
    this.options = options;
  }

  public setNotesMasterRId(rId: number): void {
    this.options = { ...this.options, notesMasterRId: rId };
  }

  public setHandoutMasterRId(rId: number): void {
    this.options = { ...this.options, handoutMasterRId: rId };
  }

  public toXml(_context: Context): string {
    const opts = this.options;

    const cx = opts.slideWidth ?? 12192000;
    const cy = opts.slideHeight ?? 6858000;

    // Root attributes
    const rootAttrs: string[] = [];
    if (opts.serverZoom) rootAttrs.push(` serverZoom="${opts.serverZoom}"`);
    if (opts.firstSlideNum !== undefined) rootAttrs.push(` firstSlideNum="${opts.firstSlideNum}"`);
    if (opts.showSpecialPlsOnTitleSld === false) rootAttrs.push(' showSpecialPlsOnTitleSld="0"');
    if (opts.rtl) rootAttrs.push(' rtl="1"');
    if (opts.removePersonalInfoOnSave) rootAttrs.push(' removePersonalInfoOnSave="1"');
    if (opts.compatMode) rootAttrs.push(' compatMode="1"');
    if (opts.strictFirstAndLastChars === false) rootAttrs.push(' strictFirstAndLastChars="0"');
    if (opts.embedTrueTypeFonts) rootAttrs.push(' embedTrueTypeFonts="1"');
    if (opts.saveSubsetFonts) rootAttrs.push(' saveSubsetFonts="1"');
    if (opts.autoCompressPictures === false) rootAttrs.push(' autoCompressPictures="0"');
    if (opts.bookmarkIdSeed !== undefined)
      rootAttrs.push(` bookmarkIdSeed="${opts.bookmarkIdSeed}"`);
    if (opts.conformance) rootAttrs.push(` conformance="${opts.conformance}"`);

    let s = `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${rootAttrs.join("")}>`;
    s += "<p:sldMasterIdLst>";
    for (let mi = 0; mi < opts.masterCount; mi++) {
      s += `<p:sldMasterId id="${2147483648 + mi * 12}" r:id="rId${mi + 1}"/>`;
    }
    s += "</p:sldMasterIdLst>";
    if (opts.notesMasterRId) {
      s += `<p:notesMasterIdLst><p:notesMasterId r:id="rId${opts.notesMasterRId}"/></p:notesMasterIdLst>`;
    }
    if (opts.handoutMasterRId) {
      s += `<p:handoutMasterIdLst><p:handoutMasterId r:id="rId${opts.handoutMasterRId}"/></p:handoutMasterIdLst>`;
    }
    s += "<p:sldIdLst>";
    const sRidOff = opts.masterCount + 1;
    for (let i = 0; i < opts.slideIds.length; i++) {
      s += `<p:sldId id="${opts.slideIds[i]}" r:id="rId${sRidOff + i}"/>`;
    }
    s += "</p:sldIdLst>";
    s += `<p:sldSz cx="${cx}" cy="${cy}"/>`;
    s += '<p:notesSz cx="6858000" cy="9144000"/>';

    if (opts.photoAlbum) {
      const pa = opts.photoAlbum;
      const paAttrs: string[] = [];
      if (pa.blackWhite) paAttrs.push(' bw="1"');
      if (pa.showCaptions) paAttrs.push(' showCaptions="1"');
      if (pa.layout) paAttrs.push(` layout="${pa.layout}"`);
      if (pa.frame) paAttrs.push(` frame="${pa.frame}"`);
      s += `<p:photoAlbum${paAttrs.join("")}/>`;
    }

    s += DEFAULT_TEXT_STYLE_XML;

    if (opts.modifyVerifier) {
      const mv = opts.modifyVerifier;
      const mvAttrs: string[] = [];
      if (mv.algorithmName) mvAttrs.push(` algorithmName="${mv.algorithmName}"`);
      if (mv.hashValue) mvAttrs.push(` hashValue="${mv.hashValue}"`);
      if (mv.saltValue) mvAttrs.push(` saltValue="${mv.saltValue}"`);
      if (mv.spinValue !== undefined) mvAttrs.push(` spinValue="${mv.spinValue}"`);
      if (mv.cryptProviderType) mvAttrs.push(` cryptProviderType="${mv.cryptProviderType}"`);
      if (mv.cryptAlgorithmClass) mvAttrs.push(` cryptAlgorithmClass="${mv.cryptAlgorithmClass}"`);
      if (mv.cryptAlgorithmType) mvAttrs.push(` cryptAlgorithmType="${mv.cryptAlgorithmType}"`);
      if (mv.cryptAlgorithmSid !== undefined)
        mvAttrs.push(` cryptAlgorithmSid="${mv.cryptAlgorithmSid}"`);
      if (mv.spinCount !== undefined) mvAttrs.push(` spinCount="${mv.spinCount}"`);
      if (mv.saltData) mvAttrs.push(` saltData="${mv.saltData}"`);
      if (mv.hashData) mvAttrs.push(` hashData="${mv.hashData}"`);
      if (mv.cryptProvider) mvAttrs.push(` cryptProvider="${mv.cryptProvider}"`);
      if (mv.algorithmIdExtension !== undefined)
        mvAttrs.push(` algIdExt="${mv.algorithmIdExtension}"`);
      if (mv.algorithmIdExtensionSource)
        mvAttrs.push(` algIdExtSource="${mv.algorithmIdExtensionSource}"`);
      if (mv.cryptProviderTypeExtension !== undefined)
        mvAttrs.push(` cryptProviderTypeExt="${mv.cryptProviderTypeExtension}"`);
      if (mv.cryptProviderTypeExtensionSource)
        mvAttrs.push(` cryptProviderTypeExtSource="${mv.cryptProviderTypeExtensionSource}"`);
      s += `<p:modifyVerifier${mvAttrs.join("")}/>`;
    }
    s += "</p:presentation>";
    return s;
  }
}
