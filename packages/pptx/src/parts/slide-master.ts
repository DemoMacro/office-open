import { convertToEmu } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { BackgroundOptions } from "@parts/background";
import type { TimingDescriptorOptions } from "@parts/descriptors/animation";
import type {
  ControlDescriptorOptions,
  TransitionDescriptorOptions,
} from "@parts/descriptors/slide";
import type { TextListStyleOptions } from "@parts/descriptors/text-list-style";
import { createPptxEffectList } from "@shared/drawingml/effects";
import { buildFill } from "@shared/drawingml/fill";
import type { MasterChild } from "@shared/file";

import type { ColorMapOptions, HeaderFooterOptions } from "./handout-master";

export interface MasterPlaceholderPosition {
  x: number | UniversalMeasure;
  y: number | UniversalMeasure;
  width: number | UniversalMeasure;
  height: number | UniversalMeasure;
}

export interface MasterPlaceholderOptions {
  title?: boolean | MasterPlaceholderPosition;
  body?: boolean | MasterPlaceholderPosition;
  date?: boolean | MasterPlaceholderPosition;
  footer?: boolean | MasterPlaceholderPosition;
  slideNumber?: boolean | MasterPlaceholderPosition;
}

export interface SlideMasterOptions {
  background?: BackgroundOptions;
  children?: MasterChild[];
  placeholders?: MasterPlaceholderOptions;
  /** Color map overrides (p:clrMap); defaults to the standard mapping. */
  colorMap?: ColorMapOptions;
  /** Header/footer visibility on the master (p:hf). */
  headerFooter?: HeaderFooterOptions;
  /**
   * Master text styles (p:txStyles — title/body/other, CT_SlideMasterTextStyles).
   * Defaults to the MS Office standard block when omitted. Round-tripped masters
   * carry their own so custom typography (fonts, indents, bullets) survives.
   */
  textStyles?: TextListStyleOptions;
  /** @preserve — keep this master even when no slide references it. */
  preserve?: boolean;
  /** p:transition — slide-transition defaults inherited by slides. */
  transition?: TransitionDescriptorOptions;
  /** p:timing — animation timeline. */
  timing?: TimingDescriptorOptions;
  /** cSld/custDataLst — relationship references to customer data parts. */
  customerData?: { rId: string }[];
  /** cSld/controls — embedded controls (ActiveX/legacy). */
  controls?: ControlDescriptorOptions[];
}

// ── Placeholder emit helpers (fresh-generate path) ──
//
// The master carries up to five standard placeholders (title/body/date/footer/
// slideNumber). Fresh generation scales their reference positions to the slide
// width; round-trip supplies explicit EMU positions parsed from spTree.

// Reference positions (16:9 master, slideWidth = 12192000 EMU)
export const SW_REF = 12192000;
const sx = (refX: number, sw: number) => Math.round((refX * sw) / SW_REF);

const REF_TITLE = { x: 838200, y: 365125, cx: 10515600, cy: 1325563 };
const REF_BODY = { x: 838200, y: 1825625, cx: 10515600, cy: 4351338 };
const REF_DATE = { x: 838200, y: 6356350, cx: 2743200, cy: 365125 };
const REF_FOOTER = { x: 4038600, y: 6356350, cx: 4114800, cy: 365125 };
const REF_SLDNUM = { x: 8610600, y: 6356350, cx: 2743200, cy: 365125 };

/** Resolve a placeholder option into a concrete EMU rect, or null when hidden. */
export function resolvePos(
  opt: boolean | MasterPlaceholderPosition | undefined,
  ref: { x: number; y: number; cx: number; cy: number },
  slideWidth: number,
): { x: number; y: number; cx: number; cy: number } | null {
  if (opt === false) return null;
  if (opt === undefined || opt === true) {
    return { x: sx(ref.x, slideWidth), y: ref.y, cx: sx(ref.cx, slideWidth), cy: ref.cy };
  }
  return {
    x: convertToEmu(opt.x),
    y: convertToEmu(opt.y),
    cx: convertToEmu(opt.width),
    cy: convertToEmu(opt.height),
  };
}

/** Emit one placeholder <p:sp> with the given geometry and txBody content. */
export function phSp(
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

export const BODY_DEFAULT = `<a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p>`;

export function footerBody(algn: string, fldType: string, fldId: string, fldText: string): string {
  const fld =
    fldType && fldId
      ? `<a:fld id="${fldId}" type="${fldType}"><a:rPr lang="en-US" smtClean="0"/><a:t>${fldText}</a:t></a:fld>`
      : "";
  return `<a:bodyPr/><a:lstStyle><a:lvl1pPr algn="${algn}"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p>${fld}<a:endParaRPr lang="en-US"/></a:p>`;
}

export interface PlaceholderEmitResult {
  /** Concatenated placeholder <p:sp> XML (empty when all hidden). */
  xml: string;
  /** Next free cNvPr id after the emitted placeholders (children start here). */
  nextId: number;
}

/**
 * Emit the master's standard placeholders in XSD order (title → body → date →
 * footer → slideNumber). Returns the concatenated XML and the next free id.
 */
export function buildPlaceholderShapes(
  placeholders: MasterPlaceholderOptions | undefined,
  slideWidth: number,
): PlaceholderEmitResult {
  const ph = placeholders ?? {};
  const shapes: string[] = [];
  let nextId = 2;

  const titlePos = resolvePos(ph.title, REF_TITLE, slideWidth);
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

  const bodyPos = resolvePos(ph.body, REF_BODY, slideWidth);
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

  const datePos = resolvePos(ph.date, REF_DATE, slideWidth);
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

  const footerPos = resolvePos(ph.footer, REF_FOOTER, slideWidth);
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

  const sldNumPos = resolvePos(ph.slideNumber, REF_SLDNUM, slideWidth);
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

  return { xml: shapes.join(""), nextId };
}

// ── Background ──

/** Emit p:bg. Undefined background emits the MS Office default bgRef idx="1001". */
export function buildBackgroundXml(bg?: BackgroundOptions): string {
  if (!bg) return '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>';
  const bgAttrs: string[] = [];
  if (bg.blackWhiteMode) bgAttrs.push(` p:bwMode="${bg.blackWhiteMode}"`);
  const bgPrAttrs: string[] = [];
  if (bg.shadeToTitle) bgPrAttrs.push(' shadeToTitle="1"');
  const fillXml = buildFill(bg.fill ?? { type: "none" });
  let effectsXml = "";
  if (bg.effects) {
    const el = createPptxEffectList(bg.effects);
    if (el) effectsXml = el;
  }
  return `<p:bg${bgAttrs.join("")}><p:bgPr${bgPrAttrs.join("")}>${fillXml}${effectsXml}</p:bgPr></p:bg>`;
}
