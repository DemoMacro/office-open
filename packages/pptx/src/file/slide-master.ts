import { Background, type BackgroundOptions } from "@file/background";
import { DEFAULT_COLOR_MAP, SP_TREE_HEADER } from "@file/constants";
import type { MasterChild } from "@file/file";
import type { SlideHeaderFooterOptions } from "@file/header-footer";
import { buildMasterChildrenXml } from "@file/slide/coerce";
import type { Context } from "@file/xml-components";
import { ImportedXmlComponent } from "@file/xml-components";
import { convertPositionToEmu } from "@office-open/core";

export interface MasterPlaceholderPosition {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface MasterPlaceholderOptions {
  readonly title?: boolean | MasterPlaceholderPosition;
  readonly body?: boolean | MasterPlaceholderPosition;
  readonly date?: boolean | MasterPlaceholderPosition;
  readonly footer?: boolean | MasterPlaceholderPosition;
  readonly slideNumber?: boolean | MasterPlaceholderPosition;
}

export interface SlideMasterOptions {
  readonly background?: BackgroundOptions;
  readonly children?: readonly MasterChild[];
  readonly placeholders?: MasterPlaceholderOptions;
}

// Reference positions (16:9 master, slideWidth = 12192000 EMU)
const SW_REF = 12192000;
const sx = (refX: number, sw: number) => Math.round((refX * sw) / SW_REF);

const REF_TITLE = { x: 838200, y: 365125, cx: 10515600, cy: 1325563 };
const REF_BODY = { x: 838200, y: 1825625, cx: 10515600, cy: 4351338 };
const REF_DATE = { x: 838200, y: 6356350, cx: 2743200, cy: 365125 };
const REF_FOOTER = { x: 4038600, y: 6356350, cx: 4114800, cy: 365125 };
const REF_SLDNUM = { x: 8610600, y: 6356350, cx: 2743200, cy: 365125 };

function resolvePos(
  opt: boolean | MasterPlaceholderPosition | undefined,
  ref: { x: number; y: number; cx: number; cy: number },
): { x: number; y: number; cx: number; cy: number } | null {
  if (opt === false) return null;
  if (opt === undefined || opt === true) return ref;
  return convertPositionToEmu(opt);
}

function phSp(
  id: number,
  name: string,
  phAttrs: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  bodyContent: string,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${phAttrs}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>${bodyContent}</p:txBody></p:sp>`;
}

const BODY_DEFAULT = `<a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p>`;

function footerBody(algn: string, fldType: string, fldId: string, fldText: string): string {
  const fld =
    fldType && fldId
      ? `<a:fld id="${fldId}" type="${fldType}"><a:rPr lang="en-US" smtClean="0"/><a:t>${fldText}</a:t></a:fld>`
      : "";
  return `<a:bodyPr/><a:lstStyle><a:lvl1pPr algn="${algn}"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p>${fld}<a:endParaRPr lang="en-US"/></a:p>`;
}

function buildBackgroundXml(bg?: BackgroundOptions): string {
  if (!bg) return `<p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>`;
  const bgObj = new Background(bg);
  const full = bgObj.toXml({ stack: [] } as Context);
  // Strip the outer <p:bg> tag — caller provides its own wrapper.
  const openEnd = full.indexOf(">");
  const closeStart = full.lastIndexOf("</p:bg>");
  return full.slice(openEnd + 1, closeStart);
}

function buildSlideMasterXml(
  layoutCount: number,
  _headerFooter?: SlideHeaderFooterOptions,
  masterOptions?: SlideMasterOptions,
  slideWidth: number = SW_REF,
  masterIndex: number = 0,
): string {
  const hfXml = `<p:hf dt="0" hdr="0" ftr="0" sldNum="0"/>`;
  const ph = masterOptions?.placeholders ?? {};

  // Layout IDs must be globally unique across all masters.
  // Master ID = 2147483648 + masterIndex * 12; layouts start at masterId + 1.
  const layoutIdBase = 2147483648 + masterIndex * 12 + 1;
  const layoutIdEntries: string[] = [];
  for (let i = 0; i < layoutCount; i++) {
    layoutIdEntries.push(`<p:sldLayoutId id="${layoutIdBase + i}" r:id="rId${i + 1}"/>`);
  }

  // Scaled reference positions (x/cx scaled, y/cy fixed)
  const sRef = (r: { x: number; y: number; cx: number; cy: number }) => ({
    x: sx(r.x, slideWidth),
    y: r.y,
    cx: sx(r.cx, slideWidth),
    cy: r.cy,
  });

  // Build placeholders
  const shapes: string[] = [];
  let nextId = 2;

  const titlePos = resolvePos(ph.title, sRef(REF_TITLE));
  if (titlePos) {
    shapes.push(
      phSp(
        nextId++,
        "Title Placeholder 1",
        'type="title"',
        titlePos.x,
        titlePos.y,
        titlePos.cx,
        titlePos.cy,
        BODY_DEFAULT,
      ),
    );
  }

  const bodyPos = resolvePos(ph.body, sRef(REF_BODY));
  if (bodyPos) {
    shapes.push(
      phSp(
        nextId++,
        "Text Placeholder 2",
        'type="body" idx="1"',
        bodyPos.x,
        bodyPos.y,
        bodyPos.cx,
        bodyPos.cy,
        BODY_DEFAULT,
      ),
    );
  }

  const datePos = resolvePos(ph.date, sRef(REF_DATE));
  if (datePos) {
    shapes.push(
      phSp(
        nextId++,
        "Date Placeholder 3",
        'type="dt" sz="half" idx="2"',
        datePos.x,
        datePos.y,
        datePos.cx,
        datePos.cy,
        footerBody("l", "datetimeFigureOut", "{5BCAD085-E8A6-8845-BD4E-CB4CCA059FC4}", "1/27/13"),
      ),
    );
  }

  const footerPos = resolvePos(ph.footer, sRef(REF_FOOTER));
  if (footerPos) {
    shapes.push(
      phSp(
        nextId++,
        "Footer Placeholder 4",
        'type="ftr" sz="quarter" idx="3"',
        footerPos.x,
        footerPos.y,
        footerPos.cx,
        footerPos.cy,
        footerBody("ctr", "", "", ""),
      ),
    );
  }

  const sldNumPos = resolvePos(ph.slideNumber, sRef(REF_SLDNUM));
  if (sldNumPos) {
    shapes.push(
      phSp(
        nextId++,
        "Slide Number Placeholder 5",
        'type="sldNum" sz="quarter" idx="4"',
        sldNumPos.x,
        sldNumPos.y,
        sldNumPos.cx,
        sldNumPos.cy,
        footerBody("r", "slidenum", "{C1FF6DA9-008F-8B48-92A6-B652298478BF}", "‹#›"),
      ),
    );
  }

  // Custom children shapes — shift ids to avoid conflicts with placeholder ids
  const childrenXml = buildMasterChildrenXml(masterOptions?.children);
  if (childrenXml) {
    const offset = nextId - 1;
    shapes.push(childrenXml.replace(/ id="(\d+)"/g, (_, n) => ` id="${parseInt(n) + offset}"`));
  }

  // Background
  const bgInnerXml = buildBackgroundXml(masterOptions?.background);

  return `<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg>${bgInnerXml}</p:bg><p:spTree>${SP_TREE_HEADER}${shapes.join("")}</p:spTree></p:cSld><p:clrMap ${DEFAULT_COLOR_MAP}/><p:sldLayoutIdLst>${layoutIdEntries.join("")}</p:sldLayoutIdLst>${hfXml}<p:txStyles><p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr marL="228600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="685800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:otherStyle></p:txStyles></p:sldMaster>`;
}

export class DefaultSlideMaster extends ImportedXmlComponent {
  private static cache = new Map<string, ImportedXmlComponent>();
  private readonly cacheKey: string;

  public constructor(
    layoutCount: number = 1,
    headerFooter?: SlideHeaderFooterOptions,
    masterOptions?: SlideMasterOptions,
    slideWidth: number = SW_REF,
    masterIndex: number = 0,
  ) {
    super("p:sldMaster");
    this.cacheKey = `${layoutCount}:${masterOptions ? JSON.stringify(masterOptions) : ""}:${slideWidth}:${masterIndex}`;
    if (!DefaultSlideMaster.cache.has(this.cacheKey)) {
      DefaultSlideMaster.cache.set(
        this.cacheKey,
        ImportedXmlComponent.fromXmlString(
          buildSlideMasterXml(layoutCount, headerFooter, masterOptions, slideWidth, masterIndex),
        ),
      );
    }
  }

  public override toXml(context: Context): string {
    return DefaultSlideMaster.cache.get(this.cacheKey)!.toXml(context);
  }
}
