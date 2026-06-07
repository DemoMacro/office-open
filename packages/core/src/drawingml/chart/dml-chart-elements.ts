/**
 * DrawingML Chart elements (c: namespace) — simple BuilderElement factories.
 *
 * Covers all 120 missing elements from dml-chart.xsd.
 * Each factory returns a XmlComponent that can be used as a child in chart XML trees.
 *
 * Reference: ISO/IEC 29500-4, dml-chart.xsd
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";

// ── CT_Boolean (val attr, default true) ──

export const createApplyToEnd = (): XmlComponent => new BuilderElement({ name: "c:applyToEnd" });
export const createApplyToFront = (): XmlComponent =>
  new BuilderElement({ name: "c:applyToFront" });
export const createApplyToSides = (): XmlComponent =>
  new BuilderElement({ name: "c:applyToSides" });
export const createAutoUpdate = (): XmlComponent => new BuilderElement({ name: "c:autoUpdate" });
export const createBubble3D = (): XmlComponent => new BuilderElement({ name: "c:bubble3D" });
export const createInvertIfNegative = (): XmlComponent =>
  new BuilderElement({ name: "c:invertIfNegative" });
export const createPlotVisOnly = (): XmlComponent => new BuilderElement({ name: "c:plotVisOnly" });
export const createShowDLblsOverMax = (): XmlComponent =>
  new BuilderElement({ name: "c:showDLblsOverMax" });
export const createShowNegBubbles = (): XmlComponent =>
  new BuilderElement({ name: "c:showNegBubbles" });
export const createShowLeaderLines = (): XmlComponent =>
  new BuilderElement({ name: "c:showLeaderLines" });
export const createSmooth = (): XmlComponent => new BuilderElement({ name: "c:smooth" });

// ── CT_UnsignedInt (val attr, required) ──

export const createExplosion = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:explosion", attributes: { val: { key: "c:val", value: val } } });

export const createFmtId = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:fmtId", attributes: { val: { key: "c:val", value: val } } });

export const createSize = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:size", attributes: { val: { key: "c:val", value: val } } });

// ── CT_Double (val attr, required) ──

export const createBubbleScale = (val: number | string): XmlComponent =>
  new BuilderElement({ name: "c:bubbleScale", attributes: { val: { key: "c:val", value: val } } });

export const createCrossesAt = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:crossesAt", attributes: { val: { key: "c:val", value: val } } });

export const createCustUnit = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:custUnit", attributes: { val: { key: "c:val", value: val } } });

export const createGapDepth = (val: number | string): XmlComponent =>
  new BuilderElement({ name: "c:gapDepth", attributes: { val: { key: "c:val", value: val } } });

export const createGapWidth = (val: number | string): XmlComponent =>
  new BuilderElement({ name: "c:gapWidth", attributes: { val: { key: "c:val", value: val } } });

export const createH = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:h", attributes: { val: { key: "c:val", value: val } } });

export const createHoleSize = (val: number | string): XmlComponent =>
  new BuilderElement({ name: "c:holeSize", attributes: { val: { key: "c:val", value: val } } });

export const createLogBase = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:logBase", attributes: { val: { key: "c:val", value: val } } });

export const createMajorUnit = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:majorUnit", attributes: { val: { key: "c:val", value: val } } });

export const createMax = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:max", attributes: { val: { key: "c:val", value: val } } });

export const createMin = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:min", attributes: { val: { key: "c:val", value: val } } });

export const createMinorUnit = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:minorUnit", attributes: { val: { key: "c:val", value: val } } });

export const createOverlap = (val: number | string): XmlComponent =>
  new BuilderElement({ name: "c:overlap", attributes: { val: { key: "c:val", value: val } } });

export const createSecondPieSize = (val: number | string): XmlComponent =>
  new BuilderElement({
    name: "c:secondPieSize",
    attributes: { val: { key: "c:val", value: val } },
  });

export const createSplitPos = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:splitPos", attributes: { val: { key: "c:val", value: val } } });

export const createW = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:w", attributes: { val: { key: "c:val", value: val } } });

export const createX = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:x", attributes: { val: { key: "c:val", value: val } } });

export const createY = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:y", attributes: { val: { key: "c:val", value: val } } });

// ── String enum val elements ──

export const createBaseTimeUnit = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:baseTimeUnit", attributes: { val: { key: "c:val", value: val } } });

export const createBuiltInUnit = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:builtInUnit", attributes: { val: { key: "c:val", value: val } } });

export const createCrossBetween = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:crossBetween", attributes: { val: { key: "c:val", value: val } } });

export const createDispBlanksAs = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:dispBlanksAs", attributes: { val: { key: "c:val", value: val } } });

export const createHMode = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:hMode", attributes: { val: { key: "c:val", value: val } } });

export const createLayoutTarget = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:layoutTarget", attributes: { val: { key: "c:val", value: val } } });

export const createLblAlgn = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:lblAlgn", attributes: { val: { key: "c:val", value: val } } });

export const createMajorTickMark = (val: string): XmlComponent =>
  new BuilderElement({
    name: "c:majorTickMark",
    attributes: { val: { key: "c:val", value: val } },
  });

export const createMajorTimeUnit = (val: string): XmlComponent =>
  new BuilderElement({
    name: "c:majorTimeUnit",
    attributes: { val: { key: "c:val", value: val } },
  });

export const createMinorTickMark = (val: string): XmlComponent =>
  new BuilderElement({
    name: "c:minorTickMark",
    attributes: { val: { key: "c:val", value: val } },
  });

export const createMinorTimeUnit = (val: string): XmlComponent =>
  new BuilderElement({
    name: "c:minorTimeUnit",
    attributes: { val: { key: "c:val", value: val } },
  });

export const createOfPieType = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:ofPieType", attributes: { val: { key: "c:val", value: val } } });

export const createPictureFormat = (val: string): XmlComponent =>
  new BuilderElement({
    name: "c:pictureFormat",
    attributes: { val: { key: "c:val", value: val } },
  });

export const createPictureStackUnit = (val: number): XmlComponent =>
  new BuilderElement({
    name: "c:pictureStackUnit",
    attributes: { val: { key: "c:val", value: val } },
  });

export const createShape = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:shape", attributes: { val: { key: "c:val", value: val } } });

export const createSizeRepresents = (val: string): XmlComponent =>
  new BuilderElement({
    name: "c:sizeRepresents",
    attributes: { val: { key: "c:val", value: val } },
  });

export const createSplitType = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:splitType", attributes: { val: { key: "c:val", value: val } } });

export const createSymbol = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:symbol", attributes: { val: { key: "c:val", value: val } } });

export const createTickLblPos = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:tickLblPos", attributes: { val: { key: "c:val", value: val } } });

export const createTickMarkSkip = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:tickMarkSkip", attributes: { val: { key: "c:val", value: val } } });

export const createThickness = (val: number | string): XmlComponent =>
  new BuilderElement({ name: "c:thickness", attributes: { val: { key: "c:val", value: val } } });

export const createWMode = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:wMode", attributes: { val: { key: "c:val", value: val } } });

export const createXMode = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:xMode", attributes: { val: { key: "c:val", value: val } } });

export const createYMode = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:yMode", attributes: { val: { key: "c:val", value: val } } });

// ── CT_NumFmt (formatCode + sourceLinked attrs) ──

export interface NumFmtOptions {
  formatCode: string;
  sourceLinked?: boolean;
}

export const createNumFmt = (options: NumFmtOptions): XmlComponent =>
  new BuilderElement({
    name: "c:numFmt",
    attributes: {
      formatCode: { key: "formatCode", value: options.formatCode },
      ...(options.sourceLinked !== undefined
        ? { sourceLinked: { key: "sourceLinked", value: options.sourceLinked } }
        : {}),
    },
  });

// ── CT_Skip (val attr, unsignedInt, required) ──

// (tickMarkSkip is above)

// ── String container elements ──

export const createName = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:name", children: [val] });

// ── Container / complex-type elements (children passed in) ──

export const createBackWall = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:backWall", children });

export const createBandFmt = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:bandFmt", children });

export const createBandFmts = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:bandFmts", children });

export const createChartObject = (): XmlComponent => new BuilderElement({ name: "c:chartObject" });

export const createClrMapOvr = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:clrMapOvr", children });

export const createCustSplit = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:custSplit", children });

export const createDLbl = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:dLbl", children });

export const createDPt = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:dPt", children });

export const createDTable = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:dTable", children });

export const createData = (): XmlComponent => new BuilderElement({ name: "c:data" });

export const createDispUnits = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:dispUnits", children });

export const createDispUnitsLbl = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:dispUnitsLbl", children });

export const createDownBars = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:downBars", children });

export const createEvenFooter = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:evenFooter", children: [val] });

export const createEvenHeader = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:evenHeader", children: [val] });

export const createExt = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:ext", children });

export const createExtLst = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:extLst", children });

export const createExternalData = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:externalData", children });

export const createFirstFooter = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:firstFooter", children: [val] });

export const createFirstHeader = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:firstHeader", children: [val] });

export const createFloor = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:floor", children });

export const createFormatting = (): XmlComponent => new BuilderElement({ name: "c:formatting" });

export interface HeaderFooterOptions {
  alignWithMargins?: boolean;
  differentOddEven?: boolean;
  differentFirst?: boolean;
}

export const createHeaderFooter = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
  options?: HeaderFooterOptions,
): XmlComponent =>
  new BuilderElement({
    name: "c:headerFooter",
    children,
    attributes: options
      ? {
          ...(options.alignWithMargins !== undefined
            ? { alignWithMargins: { key: "alignWithMargins", value: options.alignWithMargins } }
            : {}),
          ...(options.differentOddEven !== undefined
            ? { differentOddEven: { key: "differentOddEven", value: options.differentOddEven } }
            : {}),
          ...(options.differentFirst !== undefined
            ? { differentFirst: { key: "differentFirst", value: options.differentFirst } }
            : {}),
        }
      : undefined,
  });

export const createHiLowLines = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:hiLowLines", children });

export const createLeaderLines = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:leaderLines", children });

export const createLegacyDrawingHF = (): XmlComponent =>
  new BuilderElement({ name: "c:legacyDrawingHF" });

export const createLegendEntry = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:legendEntry", children });

export const createLvl = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:lvl", children });

export const createMajorGridlines = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:majorGridlines", children });

export const createManualLayout = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:manualLayout", children });

export const createMarker = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:marker", children });

export const createMinorGridlines = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:minorGridlines", children });

export const createMinus = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:minus", children });

export const createMultiLvlStrCache = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:multiLvlStrCache", children });

export const createMultiLvlStrRef = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:multiLvlStrRef", children });

export const createNumLit = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:numLit", children });

export const createOddFooter = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:oddFooter", children: [val] });

export const createOddHeader = (val: string): XmlComponent =>
  new BuilderElement({ name: "c:oddHeader", children: [val] });

export const createOfPieChart = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:ofPieChart", children });

export interface PageMarginsOptions {
  l: number;
  r: number;
  t: number;
  b: number;
  header: number;
  footer: number;
}

export const createPageMargins = (options: PageMarginsOptions): XmlComponent =>
  new BuilderElement({
    name: "c:pageMargins",
    attributes: {
      l: { key: "l", value: options.l },
      r: { key: "r", value: options.r },
      t: { key: "t", value: options.t },
      b: { key: "b", value: options.b },
      header: { key: "header", value: options.header },
      footer: { key: "footer", value: options.footer },
    },
  });

export interface PageSetupOptions {
  paperSize?: number;
  paperHeight?: string;
  paperWidth?: string;
  firstPageNumber?: number;
  orientation?: string;
  blackAndWhite?: boolean;
  draft?: boolean;
  useFirstPageNumber?: boolean;
  horizontalDpi?: number;
  verticalDpi?: number;
  copies?: number;
}

export const createPageSetup = (options?: PageSetupOptions): XmlComponent =>
  new BuilderElement({
    name: "c:pageSetup",
    attributes: options
      ? {
          ...(options.paperSize !== undefined
            ? { paperSize: { key: "paperSize", value: options.paperSize } }
            : {}),
          ...(options.paperHeight
            ? { paperHeight: { key: "paperHeight", value: options.paperHeight } }
            : {}),
          ...(options.paperWidth
            ? { paperWidth: { key: "paperWidth", value: options.paperWidth } }
            : {}),
          ...(options.firstPageNumber !== undefined
            ? { firstPageNumber: { key: "firstPageNumber", value: options.firstPageNumber } }
            : {}),
          ...(options.orientation
            ? { orientation: { key: "orientation", value: options.orientation } }
            : {}),
          ...(options.blackAndWhite !== undefined
            ? { blackAndWhite: { key: "blackAndWhite", value: options.blackAndWhite } }
            : {}),
          ...(options.draft !== undefined ? { draft: { key: "draft", value: options.draft } } : {}),
          ...(options.useFirstPageNumber !== undefined
            ? {
                useFirstPageNumber: {
                  key: "useFirstPageNumber",
                  value: options.useFirstPageNumber,
                },
              }
            : {}),
          ...(options.horizontalDpi !== undefined
            ? { horizontalDpi: { key: "horizontalDpi", value: options.horizontalDpi } }
            : {}),
          ...(options.verticalDpi !== undefined
            ? { verticalDpi: { key: "verticalDpi", value: options.verticalDpi } }
            : {}),
          ...(options.copies !== undefined
            ? { copies: { key: "copies", value: options.copies } }
            : {}),
        }
      : undefined,
  });

export const createPictureOptions = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:pictureOptions", children });

export const createPivotFmt = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:pivotFmt", children });

export const createPivotFmts = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:pivotFmts", children });

export const createPivotSource = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:pivotSource", children });

export const createPlus = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:plus", children });

export const createPrintSettings = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:printSettings", children });

export const createProtection = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:protection", children });

export const createSecondPiePt = (val: number): XmlComponent =>
  new BuilderElement({ name: "c:secondPiePt", attributes: { val: { key: "c:val", value: val } } });

export const createSelection = (): XmlComponent => new BuilderElement({ name: "c:selection" });

export const createSerLines = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:serLines", children });

export const createShowHorzBorder = (): XmlComponent =>
  new BuilderElement({ name: "c:showHorzBorder" });
export const createShowKeys = (): XmlComponent => new BuilderElement({ name: "c:showKeys" });
export const createShowOutline = (): XmlComponent => new BuilderElement({ name: "c:showOutline" });
export const createShowVertBorder = (): XmlComponent =>
  new BuilderElement({ name: "c:showVertBorder" });

export const createSideWall = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:sideWall", children });

export const createStrLit = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:strLit", children });

export const createSurface3DChart = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:surface3DChart", children });

export const createTrendlineLbl = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:trendlineLbl", children });

export const createUpBars = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:upBars", children });

export const createUpDownBars = (
  children?: readonly (import("../../xml-components").IXmlableObject | XmlComponent)[],
): XmlComponent => new BuilderElement({ name: "c:upDownBars", children });

export const createUserInterface = (): XmlComponent =>
  new BuilderElement({ name: "c:userInterface" });

export const createUserShapes = (): XmlComponent => new BuilderElement({ name: "c:userShapes" });
