import { convertToEmu } from "@office-open/core";
import type { ColorMappingOptions, UniversalMeasure } from "@office-open/core";
import type { WriteContext } from "@office-open/core/descriptor";
import {
  createShapeLocking,
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
import { stringifyCustDataLst } from "@parts/slide/c-sld";
import type { ControlOptions } from "@parts/slide/slide";
import type { SlideChild } from "@parts/slide/slide-child";
import type { AnimationsOptions } from "@shared/animation/timing";
import type { SlideCustomerDataReferenceOptions } from "@shared/customer-data";
import type { PlaceholderDefinition, PlaceholderMapOptions } from "@shared/placeholder";
import type { TransitionOptions } from "@shared/transition";

import type { PptxWriteContext } from "../context";
import type { HeaderFooterOptions } from "./handout-master";

export interface MasterPlaceholderPosition {
  x: number | UniversalMeasure;
  y: number | UniversalMeasure;
  width: number | UniversalMeasure;
  height: number | UniversalMeasure;
}

export type MasterPlaceholderOptions = PlaceholderMapOptions;

export type StandardPlaceholderKey = "title" | "body" | "date" | "footer" | "slideNumber";

export type MasterChild =
  | SlideChild
  | { placeholder: StandardPlaceholderKey }
  | { placeholderShape: number };

/** Ordered standard placeholder source shape; duplicates are significant. */
export interface MasterPlaceholderShape {
  key: StandardPlaceholderKey;
  definition: PlaceholderDefinition | false;
  /** p:nvPr `@userDrawn` on the source placeholder shape (round-trip). */
  userDrawn?: boolean;
}

export interface SlideMasterOptions {
  /** Root spTree cNvPr id — round-trip only. */
  shapeTreeId?: number;
  /** Root spTree cNvPr name — round-trip only. */
  shapeTreeName?: string;
  /** Root spTree cNvPr creation id — round-trip only. */
  shapeTreeCreationId?: string;
  /** Root spTree cNvPr extension-list inner XML — round-trip only. */
  shapeTreeExt?: string;
  /** Root spTree grpSpPr black-and-white mode (round-trip). */
  shapeTreeBlackWhiteMode?: string;
  /** Omit the master shape-tree root transform when the source omitted it. */
  groupTransformless?: boolean;
  background?: BackgroundOptions;
  /** Whether the source master carried p:bg; round-trip only. */
  backgroundPresent?: boolean;
  children?: MasterChild[];
  placeholders?: MasterPlaceholderOptions;
  /** Ordered source placeholders; preserves repeated slots such as two bodies. */
  placeholderShapes?: MasterPlaceholderShape[];
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
  /** Whether the source master carried p:txStyles; round-trip only. */
  textStylesPresent?: boolean;
  /** `@preserve` — keep this master even when no slide references it. */
  preserve?: boolean;
  /**
   * p:transition — slide-transition defaults inherited by slides. A string is
   * the verbatim markup-compatibility form used by newer duration extensions.
   */
  transition?: TransitionOptions | string;
  /** p:timing — animation timeline. */
  animations?: AnimationsOptions;
  /** cSld/custDataLst — relationship references to customer data parts. */
  customerData?: SlideCustomerDataReferenceOptions[];
  /** cSld/controls — embedded controls (ActiveX/legacy). */
  controls?: ControlOptions[];
  /**
   * Raw inner XML of the p:extLst inside p:cSld (CT_CommonSlideData tail —
   * where p14:creationId lives) — verbatim round-trip (captured from a
   * parsed source; do not hand-author).
   */
  cSldExt?: string;
  /** Raw extLst inner XML — verbatim round-trip for unmodeled extensions (captured from a parsed source; do not hand-author). */
  ext?: string;
  /** CT_ExtensionListModify/@mod for the trailing master extension list. */
  extensionListModified?: boolean;
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
  if (src.id !== undefined) dst.id = src.id;
  if (src.cNvPrExtension !== undefined) dst.cNvPrExtension = src.cNvPrExtension;
  if (src.customerData !== undefined) dst.customerData = src.customerData;
  if (src.blackWhiteMode !== undefined) dst.blackWhiteMode = src.blackWhiteMode;
  if (src.placeholderSize !== undefined) dst.placeholderSize = src.placeholderSize;
  if (src.placeholderIndex !== undefined) dst.placeholderIndex = src.placeholderIndex;
  if (src.placeholderOrientation !== undefined)
    dst.placeholderOrientation = src.placeholderOrientation;
  if (src.hasCustomPrompt !== undefined) dst.hasCustomPrompt = src.hasCustomPrompt;
  if (src.userDrawn !== undefined) dst.userDrawn = src.userDrawn;
  if (src.locking !== undefined) dst.locking = src.locking;
  if (src.lockingPresent !== undefined) dst.lockingPresent = src.lockingPresent;
  if (src.textBox !== undefined) dst.textBox = src.textBox;
  if (src.flipHorizontal !== undefined) dst.flipHorizontal = src.flipHorizontal;
  if (src.flipVertical !== undefined) dst.flipVertical = src.flipVertical;
  if (src.rotation !== undefined) dst.rotation = src.rotation;
  if (src.geometry !== undefined) dst.geometry = src.geometry;
  if (src.customGeometry !== undefined) dst.customGeometry = src.customGeometry;
  if (src.fill !== undefined) dst.fill = src.fill;
  if (src.outline !== undefined) dst.outline = src.outline;
  if (src.effects !== undefined) dst.effects = src.effects;
  if (src.effectDag !== undefined) dst.effectDag = src.effectDag;
  if (src.scene3d !== undefined) dst.scene3d = src.scene3d;
  if (src.shape3d !== undefined) dst.shape3d = src.shape3d;
  if (src.extensions !== undefined) dst.extensions = src.extensions;
  if (src.ext !== undefined) dst.ext = src.ext;
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
  const sourceId = def.id ?? id;
  // p:spPr — xfrm + geometry (defaults to rect, the placeholder standard) +
  // any inherited fill/outline/effects/3D facets.
  const spPrContent = shapePropertiesDesc.stringify(
    {
      x: def.x,
      y: def.y,
      width: def.width,
      height: def.height,
      flipHorizontal: def.flipHorizontal,
      flipVertical: def.flipVertical,
      rotation: def.rotation,
      geometry: def.customGeometry || def.geometry === null ? undefined : (def.geometry ?? "rect"),
      customGeometry: def.customGeometry,
      fill: def.fill,
      outline: def.outline,
      effects: def.effects,
      effectDag: def.effectDag,
      scene3d: def.scene3d,
      shape3d: def.shape3d,
      extensions: def.extensions,
      ext: def.ext,
    } as ShapePropertiesOptions,
    ctx,
  );
  const bwMode = def.blackWhiteMode ? ` bwMode="${def.blackWhiteMode}"` : "";
  const spPr = spPrContent ? `<p:spPr${bwMode}>${spPrContent}</p:spPr>` : `<p:spPr${bwMode}/>`;

  const styleXml = def.style ? stringifyShapeStyle(def.style, ctx) : "";
  const bodyContent = def.textBody ? textBodyDesc.stringify(def.textBody, ctx) : defaultBody;

  // cNvPr via the shared serializer — carries description/title/hidden and the
  // a16:creationId extension when the definition (round-trip) has them; fresh
  // definitions emit the same plain `<p:cNvPr id name/>` as before.
  const sourceName = def.name ?? name;
  const cNvPr = stringifyNonVisualDrawingProperties(
    "p:cNvPr",
    sourceId,
    {
      name: def.name,
      description: def.description,
      title: def.title,
      hidden: def.hidden,
      creationId: def.creationId,
      creationIdExtension: def.creationIdExtension,
      ...(def.cNvPrExtension !== undefined ? { ext: def.cNvPrExtension } : {}),
    },
    sourceName,
  );

  const locking =
    def.locking !== undefined
      ? createShapeLocking(def.locking)
      : def.lockingPresent === false
        ? ""
        : '<a:spLocks noGrp="1"/>';
  const txBox = def.textBox === undefined ? "" : ` txBox="${def.textBox ? "1" : "0"}"`;
  const customerDataXml = stringifyCustDataLst(def.customerData);
  const nvPrAttrs = def.userDrawn ? ` userDrawn="1"` : "";
  return `<p:sp><p:nvSpPr>${cNvPr}<p:cNvSpPr${txBox}>${locking}</p:cNvSpPr><p:nvPr${nvPrAttrs}><p:ph ${phAttrs}/>${customerDataXml}</p:nvPr></p:nvSpPr>${spPr}${styleXml}<p:txBody>${bodyContent}</p:txBody></p:sp>`;
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
  /** Placeholder XML by logical slot, for masters that preserve source order. */
  xmlByPlaceholder: Partial<Record<StandardPlaceholderKey, string>>;
  /** XML in the ordered source-shape form, indexed by `placeholderShapes`. */
  xmlShapes: string[];
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
  const shapes: Partial<Record<StandardPlaceholderKey, string>> = {};
  let nextId = 2;

  const titleDef = resolveDef(ph.title, REF_TITLE, slideWidth);
  if (titleDef) {
    shapes.title = phSp(
      nextId++,
      "Title Placeholder 1",
      `type="title"${placeholderSizeAttr(titleDef.placeholderSize)}${placeholderIndexAttr(
        titleDef,
      )}${placeholderOrientationAttr(titleDef)}${placeholderPromptAttr(titleDef)}`,
      titleDef,
      BODY_DEFAULT,
      ctx,
    );
  }

  const bodyDef = resolveDef(ph.body, REF_BODY, slideWidth);
  if (bodyDef) {
    shapes.body = phSp(
      nextId++,
      "Text Placeholder 2",
      `type="body"${placeholderSizeAttr(bodyDef.placeholderSize)}${placeholderIndexAttr(
        bodyDef,
        1,
      )}${placeholderOrientationAttr(bodyDef)}${placeholderPromptAttr(bodyDef)}`,
      bodyDef,
      BODY_DEFAULT,
      ctx,
    );
  }

  const dateDef = resolveDef(ph.date, REF_DATE, slideWidth);
  if (dateDef) {
    shapes.date = phSp(
      nextId++,
      "Date Placeholder 3",
      `type="dt"${placeholderSizeAttr(
        dateDef.placeholderSize === null ? null : (dateDef.placeholderSize ?? "half"),
      )}${placeholderIndexAttr(dateDef, 2)}${placeholderOrientationAttr(
        dateDef,
      )}${placeholderPromptAttr(dateDef)}`,
      dateDef,
      footerBody("l", "datetimeFigureOut", "{5BCAD085-E8A6-8845-BD4E-CB4CCA059FC4}", "1/27/13"),
      ctx,
    );
  }

  const footerDef = resolveDef(ph.footer, REF_FOOTER, slideWidth);
  if (footerDef) {
    shapes.footer = phSp(
      nextId++,
      "Footer Placeholder 4",
      `type="ftr"${placeholderSizeAttr(
        footerDef.placeholderSize === null ? null : (footerDef.placeholderSize ?? "quarter"),
      )}${placeholderIndexAttr(footerDef, 3)}${placeholderOrientationAttr(
        footerDef,
      )}${placeholderPromptAttr(footerDef)}`,
      footerDef,
      footerBody("ctr", "", "", ""),
      ctx,
    );
  }

  const sldNumDef = resolveDef(ph.slideNumber, REF_SLDNUM, slideWidth);
  if (sldNumDef) {
    shapes.slideNumber = phSp(
      nextId++,
      "Slide Number Placeholder 5",
      `type="sldNum"${placeholderSizeAttr(
        sldNumDef.placeholderSize === null ? null : (sldNumDef.placeholderSize ?? "quarter"),
      )}${placeholderIndexAttr(sldNumDef, 4)}${placeholderOrientationAttr(
        sldNumDef,
      )}${placeholderPromptAttr(sldNumDef)}`,
      sldNumDef,
      footerBody("r", "slidenum", "{C1FF6DA9-008F-8B48-92A6-B652298478BF}", "‹#›"),
      ctx,
    );
  }

  return {
    xml: Object.values(shapes).join(""),
    xmlByPlaceholder: shapes,
    xmlShapes: [],
    nextId,
  };
}

/**
 * Emit ordered source placeholders. Unlike the slot map this model permits
 * repeated standard keys (two body placeholders are valid OOXML).
 */
export function buildOrderedPlaceholderShapes(
  shapes: MasterPlaceholderShape[] | undefined,
  ctx: PptxWriteContext,
): PlaceholderEmitResult {
  const refs: Record<StandardPlaceholderKey, { x: number; y: number; cx: number; cy: number }> = {
    title: REF_TITLE,
    body: REF_BODY,
    date: REF_DATE,
    footer: REF_FOOTER,
    slideNumber: REF_SLDNUM,
  };
  const defaults: Record<
    StandardPlaceholderKey,
    { name: string; body: string; attrs: (def: PlaceholderDefinition) => string }
  > = {
    title: {
      name: "Title Placeholder 1",
      body: BODY_DEFAULT,
      attrs: (def) =>
        `type="title"${placeholderSizeAttr(def.placeholderSize)}${placeholderIndexAttr(def)}${placeholderOrientationAttr(def)}${placeholderPromptAttr(def)}`,
    },
    body: {
      name: "Text Placeholder 2",
      body: BODY_DEFAULT,
      attrs: (def) =>
        `type="body"${placeholderSizeAttr(def.placeholderSize)}${placeholderIndexAttr(def, 1)}${placeholderOrientationAttr(def)}${placeholderPromptAttr(def)}`,
    },
    date: {
      name: "Date Placeholder 3",
      body: footerBody(
        "l",
        "datetimeFigureOut",
        "{5BCAD085-E8A6-8845-BD4E-CB4CCA059FC4}",
        "1/27/13",
      ),
      attrs: (def) =>
        `type="dt"${placeholderSizeAttr(def.placeholderSize === null ? null : (def.placeholderSize ?? "half"))}${placeholderIndexAttr(def, 2)}${placeholderOrientationAttr(def)}${placeholderPromptAttr(def)}`,
    },
    footer: {
      name: "Footer Placeholder 4",
      body: footerBody("ctr", "", "", ""),
      attrs: (def) =>
        `type="ftr"${placeholderSizeAttr(def.placeholderSize === null ? null : (def.placeholderSize ?? "quarter"))}${placeholderIndexAttr(def, 3)}${placeholderOrientationAttr(def)}${placeholderPromptAttr(def)}`,
    },
    slideNumber: {
      name: "Slide Number Placeholder 5",
      body: footerBody("r", "slidenum", "{C1FF6DA9-008F-8B48-92A6-B652298478BF}", "‹#›"),
      attrs: (def) =>
        `type="sldNum"${placeholderSizeAttr(def.placeholderSize === null ? null : (def.placeholderSize ?? "quarter"))}${placeholderIndexAttr(def, 4)}${placeholderOrientationAttr(def)}${placeholderPromptAttr(def)}`,
    },
  };

  let nextId = 2;
  for (const shape of shapes ?? []) {
    if (shape.definition !== false && shape.definition.id !== undefined) {
      nextId = Math.max(nextId, shape.definition.id + 1);
    }
  }
  const xmlShapes = (shapes ?? []).map((shape) => {
    if (shape.definition === false) return "";
    const spec = defaults[shape.key];
    const def = resolveDef(shape.definition, refs[shape.key], ctx.slideWidth);
    if (!def) return "";
    return phSp(def.id ?? nextId++, spec.name, spec.attrs(def), def, spec.body, ctx);
  });
  return {
    xml: xmlShapes.join(""),
    xmlByPlaceholder: {},
    xmlShapes,
    nextId,
  };
}

function placeholderSizeAttr(size: PlaceholderDefinition["placeholderSize"]): string {
  return size === undefined || size === null ? "" : ` sz="${size}"`;
}

function placeholderIndexAttr(def: PlaceholderDefinition, fallback?: number): string {
  if (def.placeholderIndex === null) return "";
  const index = def.placeholderIndex ?? fallback;
  return index === undefined ? "" : ` idx="${index}"`;
}

function placeholderOrientationAttr(def: PlaceholderDefinition): string {
  if (def.placeholderOrientation === undefined || def.placeholderOrientation === null) return "";
  const orientation =
    def.placeholderOrientation === "horizontal"
      ? "horz"
      : def.placeholderOrientation === "vertical"
        ? "vert"
        : undefined;
  if (!orientation)
    throw new Error(`Invalid placeholder orientation: ${def.placeholderOrientation}`);
  return ` orient="${orientation}"`;
}

function placeholderPromptAttr(def: PlaceholderDefinition): string {
  return def.hasCustomPrompt === undefined || def.hasCustomPrompt === null
    ? ""
    : ` hasCustomPrompt="${def.hasCustomPrompt ? 1 : 0}"`;
}

// ── Background ──

/** Emit p:bg. Undefined background emits the MS Office default bgRef idx="1001". */
export function buildBackgroundXml(
  bg: BackgroundOptions | undefined,
  ctx: PptxWriteContext,
): string {
  return backgroundDesc.stringify(bg ?? { reference: DEFAULT_BACKGROUND_REFERENCE }, ctx) ?? "";
}
