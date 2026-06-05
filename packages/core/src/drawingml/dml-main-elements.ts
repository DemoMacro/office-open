/**
 * DrawingML Main elements (a: namespace).
 *
 * Factory functions for the structural/container elements defined in
 * dml-main.xsd.  Each returns a BuilderElement whose XML tag matches
 * the corresponding OOXML element name.
 *
 * Reference: ISO/IEC 29500-1, Part 4 - DrawingML
 *
 * @module
 */
import { BuilderElement } from "../xml-components";
import type { XmlComponent, IXmlableObject } from "../xml-components";

// ── Shapes ──────────────────────────────────────────────────────────

/** Creates an a:sp element (Shape). */
export const createSp = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:sp" });

/** Creates an a:cxnSp element (Connector Shape). */
export const createCxnSp = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:cxnSp" });

/** Creates an a:grpSp element (Group Shape). */
export const createGrpSp = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:grpSp" });

/** Creates an a:pic element (Picture). */
export const createPic = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:pic" });

/** Creates an a:graphicFrame element (Graphic Frame). */
export const createGraphicFrame = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:graphicFrame" });

/** Creates an a:txSp element (Text Shape). */
export const createTxSp = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:txSp" });

// ── Non-Visual Properties ───────────────────────────────────────────

/** Creates an a:cNvSpPr element (Non-Visual Drawing Shape Properties). */
export const createCNvSpPr = (): XmlComponent => new BuilderElement({ name: "a:cNvSpPr" });

/** Creates an a:cNvCxnSpPr element (Non-Visual Connector Properties). */
export const createCNvCxnSpPr = (): XmlComponent => new BuilderElement({ name: "a:cNvCxnSpPr" });

/** Creates an a:cNvGrpSpPr element (Non-Visual Group Shape Properties). */
export const createCNvGrpSpPr = (): XmlComponent => new BuilderElement({ name: "a:cNvGrpSpPr" });

/** Creates an a:cNvGraphicFramePr element (Non-Visual Graphic Frame Properties). */
export const createCNvGraphicFramePr = (): XmlComponent =>
  new BuilderElement({ name: "a:cNvGraphicFramePr" });

/** Creates an a:nvSpPr element (Shape Non-Visual Properties). */
export const createNvSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:nvSpPr" });

/** Creates an a:nvCxnSpPr element (Connector Non-Visual Properties). */
export const createNvCxnSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:nvCxnSpPr" });

/** Creates an a:nvGrpSpPr element (Group Shape Non-Visual Properties). */
export const createNvGrpSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:nvGrpSpPr" });

/** Creates an a:nvGraphicFramePr element (Graphic Frame Non-Visual Properties). */
export const createNvGraphicFramePr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:nvGraphicFramePr" });

/** Creates an a:nvPicPr element (Picture Non-Visual Properties). */
export const createNvPicPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:nvPicPr" });

// ── Group Shape / Locking ───────────────────────────────────────────

/** Creates an a:grpSpPr element (Group Shape Properties). */
export const createGrpSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:grpSpPr" });

/** Creates an a:cpLocks element (Content Part Locking). */
export const createCpLocks = (): XmlComponent => new BuilderElement({ name: "a:cpLocks" });

/** Creates an a:cxnSpLocks element (Connector Locking). */
export const createCxnSpLocks = (): XmlComponent => new BuilderElement({ name: "a:cxnSpLocks" });

/** Creates an a:useSpRect element (Use Shape Rectangle). */
export const createUseSpRect = (): XmlComponent => new BuilderElement({ name: "a:useSpRect" });

// ── Table ────────────────────────────────────────────────────────────

/** Creates an a:tableStyle element (Table Style). */
export const createTableStyleElement = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:tableStyle" });

/** Creates an a:cell3D element (Cell 3-D Properties). */
export const createCell3D = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:cell3D" });

/** Creates an a:insideH element (Inside Horizontal Border). */
export const createInsideH = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:insideH" });

/** Creates an a:insideV element (Inside Vertical Border). */
export const createInsideV = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:insideV" });

/** Creates an a:top element (Top Border). */
export const createTop = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:top" });

/** Creates an a:bottom element (Bottom Border). */
export const createBottom = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:bottom" });

/** Creates an a:left element (Left Border). */
export const createLeft = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:left" });

/** Creates an a:right element (Right Border). */
export const createRight = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:right" });

/** Creates an a:tl2br element (Top-Left to Bottom-Right Border). */
export const createTl2br = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:tl2br" });

/** Creates an a:tr2bl element (Top-Right to Bottom-Left Border). */
export const createTr2bl = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:tr2bl" });

/** Creates an a:lnTlToBr element (Top-Left to Bottom-Right Line). */
export const createLnTlToBr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:lnTlToBr" });

/** Creates an a:lnBlToTr element (Bottom-Left to Top-Right Line). */
export const createLnBlToTr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:lnBlToTr" });

/** Creates an a:round element (Round Line Join). */
export const createRound = (): XmlComponent => new BuilderElement({ name: "a:round" });

// ── Connector ────────────────────────────────────────────────────────

/** Creates an a:start element (Connection Start Time). */
export const createStart = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:start" });

/** Creates an a:end element (Audio CD End Time). */
export const createEnd = (options?: AudioCdTimeOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: number }> = {};
  if (options?.track !== undefined) attrs.track = { key: "track", value: options.track };
  if (options?.time !== undefined) attrs.time = { key: "time", value: options.time };
  return new BuilderElement({
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
    name: "a:end",
  });
};

/** Creates an a:stCxn element (Start Connection). */
export const createStCxn = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:stCxn" });

/** Creates an a:endCxn element (End Connection). */
export const createEndCxn = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:endCxn" });

// ── Theme ────────────────────────────────────────────────────────────

/** Creates an a:clrMap element (Color Map). */
export const createClrMap = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:clrMap" });

/** Creates an a:overrideClrMapping element (Override Color Mapping). */
export const createOverrideClrMapping = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:overrideClrMapping" });

/** Creates an a:extraClrScheme element (Extra Color Scheme). */
export const createExtraClrScheme = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:extraClrScheme" });

/** Creates an a:custClr element (Custom Color). */
export const createCustClr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:custClr" });

/** Creates an a:custClrLst element (Custom Color List). */
export const createCustClrLst = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:custClrLst" });

/** Creates an a:themeManager element (Theme Manager). */
export const createThemeManager = (): XmlComponent =>
  new BuilderElement({ name: "a:themeManager" });

/** Creates an a:themeOverride element (Theme Override). */
export const createThemeOverride = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:themeOverride" });

/** Creates an a:font element (Supplemental Font / Font Collection). */
export const createFontElement = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:font" });

/** Creates an a:spDef element (Shape Default). */
export const createSpDef = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:spDef" });

/** Creates an a:lnDef element (Line Default). */
export const createLnDef = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:lnDef" });

/** Creates an a:txDef element (Text Default). */
export const createTxDef = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:txDef" });

// ── Text ─────────────────────────────────────────────────────────────

/** Creates an a:br element (Text Line Break). */
export const createBr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:br" });

/** Creates an a:tab element (Text Tab Stop). */
export const createTab = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:tab" });

/** Creates an a:tabLst element (Text Tab Stop List). */
export const createTabLst = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:tabLst" });

/** Creates an a:sym element (Symbol Font). */
export const createSym = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:sym" });

/** Creates an a:highlight element (Text Highlight). */
export const createHighlight = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:highlight" });

/** Creates an a:buBlip element (Bullet Blip). */
export const createBuBlip = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:buBlip" });

/** Creates an a:buClrTx element (Bullet Color - Follow Text). */
export const createBuClrTx = (): XmlComponent => new BuilderElement({ name: "a:buClrTx" });

/** Creates an a:buFontTx element (Bullet Font - Follow Text). */
export const createBuFontTx = (): XmlComponent => new BuilderElement({ name: "a:buFontTx" });

/** Creates an a:buSzPts element (Bullet Size - Points). */
export const createBuSzPts = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:buSzPts" });

/** Creates an a:buSzTx element (Bullet Size - Follow Text). */
export const createBuSzTx = (): XmlComponent => new BuilderElement({ name: "a:buSzTx" });

// ── Underline ────────────────────────────────────────────────────────

/** Creates an a:uFill element (Underline Fill). */
export const createUFill = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:uFill" });

/** Creates an a:uFillTx element (Underline Fill - Follow Text). */
export const createUFillTx = (): XmlComponent => new BuilderElement({ name: "a:uFillTx" });

/** Creates an a:uLn element (Underline Line Properties). */
export const createULn = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:uLn" });

/** Creates an a:uLnTx element (Underline Line - Follow Text). */
export const createULnTx = (): XmlComponent => new BuilderElement({ name: "a:uLnTx" });

// ── Media ────────────────────────────────────────────────────────────

/** Creates an a:audioCd element (Audio CD). */
export const createAudioCd = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:audioCd" });

/** Creates an a:quickTimeFile element (QuickTime File). */
export const createQuickTimeFile = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:quickTimeFile" });

/** Creates an a:snd element (Sound / Embedded WAV Audio). */
export const createSnd = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:snd" });

/** Creates an a:wavAudioFile element (WAV Audio File). */
export const createWavAudioFile = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:wavAudioFile" });

// ── Animation Build ──────────────────────────────────────────────────

export interface BuildChartOptions {
  readonly build?: string;
  readonly animateBackground?: boolean;
}

/** Creates an a:bldChart element (Build Chart). */
export const createBldChart = (options?: BuildChartOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.build !== undefined) attrs.build = { key: "bld", value: options.build };
  if (options?.animateBackground !== undefined)
    attrs.animateBackground = { key: "animBg", value: options.animateBackground ? "1" : "0" };
  return new BuilderElement({
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
    name: "a:bldChart",
  });
};

export interface BuildDiagramOptions {
  readonly build?: string;
  readonly reverse?: boolean;
}

/** Creates an a:bldDgm element (Build Diagram). */
export const createBldDgm = (options?: BuildDiagramOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.build !== undefined) attrs.build = { key: "bld", value: options.build };
  if (options?.reverse !== undefined)
    attrs.reverse = { key: "rev", value: options.reverse ? "1" : "0" };
  return new BuilderElement({
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
    name: "a:bldDgm",
  });
};

// ── Other ────────────────────────────────────────────────────────────

/** Creates an a:bevel element (Line Join Bevel). */
export const createBevelElement = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:bevel" });

export interface ChartAnimationOptions {
  readonly seriesIndex?: number;
  readonly categoryIndex?: number;
  readonly buildStep?: string;
}

/** Creates an a:chart element (Animation Chart Element). */
export const createChart = (options?: ChartAnimationOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string | number }> = {};
  if (options?.seriesIndex !== undefined)
    attrs.seriesIndex = { key: "seriesIdx", value: options.seriesIndex };
  if (options?.categoryIndex !== undefined)
    attrs.categoryIndex = { key: "categoryIdx", value: options.categoryIndex };
  if (options?.buildStep !== undefined)
    attrs.buildStep = { key: "bldStep", value: options.buildStep };
  return new BuilderElement({
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
    name: "a:chart",
  });
};

export interface DiagramAnimationOptions {
  readonly id?: string;
  readonly buildStep?: string;
}

/** Creates an a:dgm element (Animation Diagram Element). */
export const createDgm = (options?: DiagramAnimationOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options?.id !== undefined) attrs.id = { key: "id", value: options.id };
  if (options?.buildStep !== undefined)
    attrs.buildStep = { key: "bldStep", value: options.buildStep };
  return new BuilderElement({
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
    name: "a:dgm",
  });
};

/** Creates an a:headers element (Headers). */
export const createHeaders = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:headers" });

/** Creates an a:header element (Header). */
export const createHeader = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:header" });

/** Creates an a:hlinkMouseOver element (Hyperlink Mouse Over). */
export const createHlinkMouseOver = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:hlinkMouseOver" });

export interface AudioCdTimeOptions {
  readonly track?: number;
  readonly time?: number;
}

/** Creates an a:st element (Audio CD Start Time). */
export const createSt = (options?: AudioCdTimeOptions): XmlComponent => {
  const attrs: Record<string, { key: string; value: number }> = {};
  if (options?.track !== undefined) attrs.track = { key: "track", value: options.track };
  if (options?.time !== undefined) attrs.time = { key: "time", value: options.time };
  return new BuilderElement({
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
    name: "a:st",
  });
};

/** Creates an a:style element (Shape Style). */
export const createStyleElement = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "a:style" });
