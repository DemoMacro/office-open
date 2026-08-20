import { convertToEmu } from "@office-open/core";
import type { ColorMappingOptions, UniversalMeasure } from "@office-open/core";
import type { WriteContext } from "@office-open/core/descriptor";
import {
  pickNonVisualDrawingProperties,
  shapePropertiesDesc,
  stringifyNonVisualDrawingProperties,
  textBodyDesc,
} from "@office-open/core/drawing";
import type { ShapePropertiesOptions } from "@office-open/core/drawing";
import type { BackgroundOptions } from "@parts/background";
import { DEFAULT_BACKGROUND_REFERENCE } from "@parts/background";
import { backgroundDesc } from "@parts/descriptors/background";
import { stringifyShapeStyle } from "@parts/descriptors/shape";
import type { TextStylesOptions } from "@parts/descriptors/text-list-style";
import type { ControlOptions } from "@parts/slide/slide";
import type { SlideChild } from "@parts/slide/slide-child";
import type { AnimationsOptions } from "@shared/animation/timing";
import type { PlaceholderDefinition } from "@shared/placeholder";
import type { TransitionOptions } from "@shared/transition";

import type { PptxWriteContext } from "../context";
import type { HeaderFooterOptions } from "./handout-master";

export interface MasterPlaceholderPosition {
  x: number | UniversalMeasure;
  y: number | UniversalMeasure;
  width: number | UniversalMeasure;
  height: number | UniversalMeasure;
}

export interface MasterPlaceholderOptions {
  title?: PlaceholderDefinition | false;
  body?: PlaceholderDefinition | false;
  date?: PlaceholderDefinition | false;
  footer?: PlaceholderDefinition | false;
  slideNumber?: PlaceholderDefinition | false;
}

export interface SlideMasterOptions {
  background?: BackgroundOptions;
  children?: SlideChild[];
  placeholders?: MasterPlaceholderOptions;
  /** Color mapping overrides (p:clrMap); defaults to the standard mapping. */
  colorMapping?: Partial<ColorMappingOptions>;
  /** Header/footer visibility on the master (p:hf). */
  headerFooter?: HeaderFooterOptions;
  /**
   * Master text styles (p:txStyles — title/body/other, CT_SlideMasterTextStyles).
   * Defaults to the MS Office standard block when omitted. Round-tripped masters
   * carry their own so custom typography (fonts, indents, bullets) survives.
   */
  textStyles?: TextStylesOptions;
  /** `@preserve` — keep this master even when no slide references it. */
  preserve?: boolean;
  /** p:transition — slide-transition defaults inherited by slides. */
  transition?: TransitionOptions;
  /** p:timing — animation timeline. */
  animations?: AnimationsOptions;
  /** cSld/custDataLst — relationship references to customer data parts. */
  customerData?: { rId: string }[];
  /** cSld/controls — embedded controls (ActiveX/legacy). */
  controls?: ControlOptions[];
  /**
   * Raw inner XML of the p:extLst inside p:cSld (CT_CommonSlideData tail —
   * where p14:creationId lives) — verbatim round-trip.
   */
  cSldExt?: string;
  /** Raw extLst inner XML — verbatim round-trip for unmodeled extensions. */
  ext?: string;
}

// ── Placeholder emit helpers (fresh-generate + round-trip path) ──
//
// The master carries up to five standard placeholders (title/body/date/footer/
// slideNumber). Fresh generation scales their reference positions to the slide
// width; round-trip supplies explicit EMU positions parsed from spTree. Both
// paths now flow through shapePropertiesDesc/textBodyDesc so custom facets
// (geometry/fill/outline/effects/textBody/style) round-trip losslessly while
// the default layout stays byte-equivalent with MS Office's master output.

// Reference positions (16:9 master, slideWidth = 12192000 EMU)
export const SW_REF = 12192000;
const sx = (refX: number, sw: number) => Math.round((refX * sw) / SW_REF);

const REF_TITLE = { x: 838200, y: 365125, cx: 10515600, cy: 1325563 };
const REF_BODY = { x: 838200, y: 1825625, cx: 10515600, cy: 4351338 };
const REF_DATE = { x: 838200, y: 6356350, cx: 2743200, cy: 365125 };
const REF_FOOTER = { x: 4038600, y: 6356350, cx: 4114800, cy: 365125 };
const REF_SLDNUM = { x: 8610600, y: 6356350, cx: 2743200, cy: 365125 };

/**
 * Resolve a placeholder option into a concrete EMU position, or null when
 * hidden. Legacy helper kept for callers that only need the rect.
 */
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

/**
 * Resolve a placeholder option into a full definition (EMU position + carried
 * facets), or null when hidden. Fresh input (undefined/true) takes the
 * reference position; an explicit definition preserves its facets for emit.
 */
function resolveDef(
  opt: PlaceholderDefinition | false | undefined,
  ref: { x: number; y: number; cx: number; cy: number },
  slideWidth: number,
): PlaceholderDefinition | null {
  if (opt === false) return null;
  const def: Partial<PlaceholderDefinition> = {};
  if (opt === undefined) {
    def.x = sx(ref.x, slideWidth);
    def.y = ref.y;
    def.width = sx(ref.cx, slideWidth);
    def.height = ref.cy;
  } else {
    def.x = convertToEmu(opt.x);
    def.y = convertToEmu(opt.y);
    def.width = convertToEmu(opt.width);
    def.height = convertToEmu(opt.height);
    copyFacets(opt, def);
    Object.assign(def, pickNonVisualDrawingProperties(opt));
  }
  return def as PlaceholderDefinition;
}

function copyFacets(src: PlaceholderDefinition, dst: Partial<PlaceholderDefinition>): void {
  if (src.geometry !== undefined) dst.geometry = src.geometry;
  if (src.customGeometry !== undefined) dst.customGeometry = src.customGeometry;
  if (src.fill !== undefined) dst.fill = src.fill;
  if (src.outline !== undefined) dst.outline = src.outline;
  if (src.effects !== undefined) dst.effects = src.effects;
  if (src.scene3d !== undefined) dst.scene3d = src.scene3d;
  if (src.shape3d !== undefined) dst.shape3d = src.shape3d;
  if (src.textBody !== undefined) dst.textBody = src.textBody;
  if (src.style !== undefined) dst.style = src.style;
}

/** Emit one placeholder <p:sp> from a full definition (position + facets). */
function phSp(
  id: number,
  name: string,
  phAttrs: string,
  def: PlaceholderDefinition,
  defaultBody: string,
  ctx: WriteContext,
): string {
  // p:spPr — xfrm + geometry (defaults to rect, the placeholder standard) +
  // any inherited fill/outline/effects/3D facets.
  const spPrContent = shapePropertiesDesc.stringify(
    {
      x: def.x,
      y: def.y,
      width: def.width,
      height: def.height,
      geometry: def.customGeometry ? undefined : (def.geometry ?? "rect"),
      customGeometry: def.customGeometry,
      fill: def.fill,
      outline: def.outline,
      effects: def.effects,
      scene3d: def.scene3d,
      shape3d: def.shape3d,
    } as ShapePropertiesOptions,
    ctx,
  );
  const spPr = spPrContent ? `<p:spPr>${spPrContent}</p:spPr>` : "<p:spPr/>";

  const styleXml = def.style ? stringifyShapeStyle(def.style, ctx) : "";
  const bodyContent = def.textBody ? textBodyDesc.stringify(def.textBody, ctx) : defaultBody;

  // cNvPr via the shared serializer — carries description/title/hidden and the
  // a16:creationId extension when the definition (round-trip) has them; fresh
  // definitions emit the same plain `<p:cNvPr id name/>` as before.
  const cNvPr = stringifyNonVisualDrawingProperties("p:cNvPr", id, def, name);

  return `<p:sp><p:nvSpPr>${cNvPr}<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${phAttrs}/></p:nvPr></p:nvSpPr>${spPr}${styleXml}<p:txBody>${bodyContent}</p:txBody></p:sp>`;
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
  ctx: PptxWriteContext,
): PlaceholderEmitResult {
  const ph = placeholders ?? {};
  const slideWidth = ctx.slideWidth;
  const shapes: string[] = [];
  let nextId = 2;

  const titleDef = resolveDef(ph.title, REF_TITLE, slideWidth);
  if (titleDef) {
    shapes.push(phSp(nextId++, "Title Placeholder 1", 'type="title"', titleDef, BODY_DEFAULT, ctx));
  }

  const bodyDef = resolveDef(ph.body, REF_BODY, slideWidth);
  if (bodyDef) {
    shapes.push(
      phSp(nextId++, "Text Placeholder 2", 'type="body" idx="1"', bodyDef, BODY_DEFAULT, ctx),
    );
  }

  const dateDef = resolveDef(ph.date, REF_DATE, slideWidth);
  if (dateDef) {
    shapes.push(
      phSp(
        nextId++,
        "Date Placeholder 3",
        'type="dt" sz="half" idx="2"',
        dateDef,
        footerBody("l", "datetimeFigureOut", "{5BCAD085-E8A6-8845-BD4E-CB4CCA059FC4}", "1/27/13"),
        ctx,
      ),
    );
  }

  const footerDef = resolveDef(ph.footer, REF_FOOTER, slideWidth);
  if (footerDef) {
    shapes.push(
      phSp(
        nextId++,
        "Footer Placeholder 4",
        'type="ftr" sz="quarter" idx="3"',
        footerDef,
        footerBody("ctr", "", "", ""),
        ctx,
      ),
    );
  }

  const sldNumDef = resolveDef(ph.slideNumber, REF_SLDNUM, slideWidth);
  if (sldNumDef) {
    shapes.push(
      phSp(
        nextId++,
        "Slide Number Placeholder 5",
        'type="sldNum" sz="quarter" idx="4"',
        sldNumDef,
        footerBody("r", "slidenum", "{C1FF6DA9-008F-8B48-92A6-B652298478BF}", "‹#›"),
        ctx,
      ),
    );
  }

  return { xml: shapes.join(""), nextId };
}

// ── Background ──

/** Emit p:bg. Undefined background emits the MS Office default bgRef idx="1001". */
export function buildBackgroundXml(bg?: BackgroundOptions): string {
  return (
    backgroundDesc.stringify(
      bg ?? { reference: DEFAULT_BACKGROUND_REFERENCE },
      undefined as never,
    ) ?? ""
  );
}
