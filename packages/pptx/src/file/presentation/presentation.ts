import { XmlComponent } from "@file/xml-components";
import { ImportedXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

export interface IPresentationOptions {
  readonly slideWidth?: number;
  readonly slideHeight?: number;
  readonly slideIds: readonly number[];
  readonly masterCount: number;
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

const DEFAULT_TEXT_STYLE_OBJ: IXmlableObject = ImportedXmlComponent.fromXmlString(
  DEFAULT_TEXT_STYLE_XML,
).prepForXml({ stack: [] })!;

/**
 * p:presentation — Root element of a PPTX file.
 * Lazy: stores options, builds XML object directly in prepForXml.
 */
export class Presentation extends XmlComponent {
  private readonly options: IPresentationOptions;

  public constructor(options: IPresentationOptions) {
    super("p:presentation");
    this.options = options;
  }

  public prepForXml(_context: IContext): IXmlableObject | undefined {
    const opts = this.options;

    const children: IXmlableObject[] = [];

    // xmlns attributes
    children.push({
      _attr: {
        "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      },
    });

    // sldMasterIdLst (dynamic — supports multiple masters)
    const masterIds = Array.from({ length: opts.masterCount }, (_, i) => ({
      "p:sldMasterId": { _attr: { id: 2147483648 + i * 12, "r:id": `rId${i + 1}` } },
    }));
    children.push({ "p:sldMasterIdLst": masterIds });

    // sldIdLst (dynamic) — rIds follow after master rIds
    const slideRidOffset = opts.masterCount + 1;
    const sldIds = opts.slideIds.map((id, i) => ({
      "p:sldId": { _attr: { id, "r:id": `rId${slideRidOffset + i}` } },
    }));
    children.push({ "p:sldIdLst": sldIds });

    // sldSz
    const cx = opts.slideWidth ?? 12192000;
    const cy = opts.slideHeight ?? 6858000;
    children.push({
      "p:sldSz": {
        _attr: {
          cx,
          cy,
        },
      },
    });

    // notesSz
    children.push({
      "p:notesSz": { _attr: { cx: 6858000, cy: 9144000 } },
    });

    // defaultTextStyle (cached static singleton)
    children.push(DEFAULT_TEXT_STYLE_OBJ);

    return { "p:presentation": children };
  }

  public toXml(_context: IContext): string {
    const opts = this.options;

    const cx = opts.slideWidth ?? 12192000;
    const cy = opts.slideHeight ?? 6858000;

    let s = `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">`;
    s += "<p:sldMasterIdLst>";
    for (let mi = 0; mi < opts.masterCount; mi++) {
      s += `<p:sldMasterId id="${2147483648 + mi * 12}" r:id="rId${mi + 1}"/>`;
    }
    s += "</p:sldMasterIdLst>";
    s += "<p:sldIdLst>";
    const sRidOff = opts.masterCount + 1;
    for (let i = 0; i < opts.slideIds.length; i++) {
      s += `<p:sldId id="${opts.slideIds[i]}" r:id="rId${sRidOff + i}"/>`;
    }
    s += "</p:sldIdLst>";
    s += `<p:sldSz cx="${cx}" cy="${cy}"/>`;
    s += '<p:notesSz cx="6858000" cy="9144000"/>';
    s += DEFAULT_TEXT_STYLE_XML;
    s += "</p:presentation>";
    return s;
  }
}
