/**
 * Bidirectional mappings between user-friendly values and XSD abbreviated values.
 *
 * When XSD uses full English words (e.g. "center", "start"), values are used directly — no mapping needed.
 * When XSD uses abbreviations (e.g. "ctr", "l", "rnd"), this module maps them to full words.
 *
 * Usage in generation (Options → XML):  xsdAlign.to("center") → "ctr"
 * Usage in parsing    (XML → Options):  xsdAlign.from("ctr")  → "center"
 */

/** Invert a Record<K, V> into Record<V, K>. */
export function invertMap<K extends string, V extends string>(map: Record<K, V>): Record<V, K> {
  const result = {} as Record<V, K>;
  for (const key of Object.keys(map) as K[]) {
    result[map[key]] = key;
  }
  return result;
}

/** Create a bidirectional mapping helper from a single forward map. */
function bidi<K extends string, V extends string>(forward: Record<K, V>) {
  const reverse = invertMap(forward);
  return {
    /** User-friendly value → XSD value */
    to: (key: string): string => (forward as Record<string, string>)[key] ?? key,
    /** XSD value → user-friendly value */
    from: (xsd: string): string => (reverse as Record<string, string>)[xsd] ?? xsd,
    /** The forward map (user → XSD) */
    forward,
    /** The reverse map (XSD → user) */
    reverse,
  };
}

// ---------------------------------------------------------------------------
// DrawingML — Direction / Position (ST_RectAlignment)
// Used by: RectAlignment, TileAlignment, ReflectionAlignment
// ---------------------------------------------------------------------------

export const xsdRectAlignment = bidi({
  topLeft: "tl",
  top: "t",
  topRight: "tr",
  left: "l",
  center: "ctr",
  right: "r",
  bottomLeft: "bl",
  bottom: "b",
  bottomRight: "br",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Text alignment (ST_TextAlignType)
// Used by: TextAlignment (pptx)
// ---------------------------------------------------------------------------

export const xsdTextAlign = bidi({
  left: "l",
  center: "ctr",
  right: "r",
  justify: "just",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Text anchoring (ST_TextAnchoringType)
// Used by: VerticalAlignment (pptx)
// ---------------------------------------------------------------------------

export const xsdTextAnchor = bidi({
  top: "t",
  center: "ctr",
  bottom: "b",
  justify: "just",
  distribute: "dist",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Line cap (ST_LineCap)
// ---------------------------------------------------------------------------

export const xsdLineCap = bidi({
  round: "rnd",
  square: "sq",
  flat: "flat",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Compound line (ST_CompoundLine)
// ---------------------------------------------------------------------------

export const xsdCompoundLine = bidi({
  single: "sng",
  double: "dbl",
  thickThin: "thickThin",
  thinThick: "thinThick",
  triple: "tri",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Pen alignment (ST_PenAlignment)
// ---------------------------------------------------------------------------

export const xsdPenAlignment = bidi({
  center: "ctr",
  inside: "in",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Line end size (ST_LineEndWidth, ST_LineEndLength)
// ---------------------------------------------------------------------------

export const xsdLineEndSize = bidi({
  small: "sm",
  medium: "med",
  large: "lg",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Blend mode (ST_BlendMode)
// ---------------------------------------------------------------------------

export const xsdBlendMode = bidi({
  over: "over",
  multiply: "mult",
  screen: "screen",
  darken: "darken",
  lighten: "lighten",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Path fill mode (ST_PathFillMode)
// ---------------------------------------------------------------------------

export const xsdPathFillMode = bidi({
  none: "none",
  normal: "norm",
  lighten: "lighten",
  lightenLess: "lightenLess",
  darken: "darken",
  darkenLess: "darkenLess",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Effect container type
// ---------------------------------------------------------------------------

export const xsdEffectContainer = bidi({
  sibling: "sib",
  tree: "tree",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Preset shadow (ST_PresetShadowVal)
// ---------------------------------------------------------------------------

export const xsdPresetShadow = bidi({
  shadow1: "shdw1",
  shadow2: "shdw2",
  shadow3: "shdw3",
  shadow4: "shdw4",
  shadow5: "shdw5",
  shadow6: "shdw6",
  shadow7: "shdw7",
  shadow8: "shdw8",
  shadow9: "shdw9",
  shadow10: "shdw10",
  shadow11: "shdw11",
  shadow12: "shdw12",
  shadow13: "shdw13",
  shadow14: "shdw14",
  shadow15: "shdw15",
  shadow16: "shdw16",
  shadow17: "shdw17",
  shadow18: "shdw18",
  shadow19: "shdw19",
  shadow20: "shdw20",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Preset material type (ST_PresetMaterialType)
// Only the abbreviated ones need mapping; full-word values pass through.
// ---------------------------------------------------------------------------

export const xsdLightRigDirection = bidi({
  topLeft: "tl",
  top: "t",
  topRight: "tr",
  left: "l",
  right: "r",
  bottomLeft: "bl",
  bottom: "b",
  bottomRight: "br",
} as const);

export const xsdMaterialType = bidi({
  legacyMatte: "legacyMatte",
  legacyPlastic: "legacyPlastic",
  legacyMetal: "legacyMetal",
  legacyWireframe: "legacyWireframe",
  matte: "matte",
  plastic: "plastic",
  metal: "metal",
  warmMatte: "warmMatte",
  translucentPowder: "translucentPowder",
  powder: "powder",
  darkEdge: "dkEdge",
  softEdge: "softEdge",
  clear: "clear",
  flat: "flat",
  softMetal: "softmetal",
} as const);

// ---------------------------------------------------------------------------
// DrawingML — Preset pattern fill (ST_PresetPatternVal)
// Only the abbreviated ones are mapped; values already using full words pass through.
// ---------------------------------------------------------------------------

export const xsdPattern = bidi({
  percent5: "pct5",
  percent10: "pct10",
  percent20: "pct20",
  percent25: "pct25",
  percent30: "pct30",
  percent40: "pct40",
  percent50: "pct50",
  percent60: "pct60",
  percent70: "pct70",
  percent75: "pct75",
  percent80: "pct80",
  percent90: "pct90",
  horizontal: "horz",
  vertical: "vert",
  lightHorizontal: "ltHorz",
  lightVertical: "ltVert",
  darkHorizontal: "dkHorz",
  darkVertical: "dkVert",
  narrowHorizontal: "narHorz",
  narrowVertical: "narVert",
  dashedHorizontal: "dashHorz",
  dashedVertical: "dashVert",
  cross: "cross",
  downDiagonal: "dnDiag",
  upDiagonal: "upDiag",
  lightDownDiagonal: "ltDnDiag",
  lightUpDiagonal: "ltUpDiag",
  darkDownDiagonal: "dkDnDiag",
  darkUpDiagonal: "dkUpDiag",
  wideDownDiagonal: "wdDnDiag",
  wideUpDiagonal: "wdUpDiag",
  dashedDownDiagonal: "dashDnDiag",
  dashedUpDiagonal: "dashUpDiag",
  diagonalCross: "diagCross",
  smallChecker: "smCheck",
  largeChecker: "lgCheck",
  smallGrid: "smGrid",
  largeGrid: "lgGrid",
  dotGrid: "dotGrid",
  smallConfetti: "smConfetti",
  largeConfetti: "lgConfetti",
  horizontalBrick: "horzBrick",
  diagonalBrick: "diagBrick",
  solidDiamond: "solidDmnd",
  openDiamond: "openDmnd",
  dottedDiamond: "dotDmnd",
  plaid: "plaid",
  sphere: "sphere",
  weave: "weave",
  divot: "divot",
  shingle: "shingle",
  wave: "wave",
  trellis: "trellis",
  zigZag: "zigZag",
} as const);

// ---------------------------------------------------------------------------
// DOCX — Vertical merge revision (ST_VerticalMergeRestart)
// ---------------------------------------------------------------------------

export const xsdVerticalMergeRev = bidi({
  continue: "cont",
  restart: "rest",
} as const);

// ---------------------------------------------------------------------------
// PPTX — Underline style (ST_TextUnderlineType, abbreviated subset)
// ---------------------------------------------------------------------------

export const xsdUnderlineStyle = bidi({
  single: "sng",
  double: "dbl",
  none: "none",
} as const);

// ---------------------------------------------------------------------------
// PPTX — Strike style (ST_TextStrikeType)
// ---------------------------------------------------------------------------

export const xsdStrikeStyle = bidi({
  singleStrike: "sngStrike",
  doubleStrike: "dblStrike",
  noStrike: "noStrike",
} as const);

// ---------------------------------------------------------------------------
// PPTX — Text capitalization (ST_TextCapsType)
// ---------------------------------------------------------------------------

export const xsdTextCaps = bidi({
  none: "none",
  all: "all",
  small: "small",
} as const);

// ---------------------------------------------------------------------------
// Chart — axis / legend / label placement (c:axPos, c:legendPos, c:dLblPos)
// Only the abbreviated tokens are mapped; identical values pass through.
// ---------------------------------------------------------------------------

export const xsdAxisPosition = bidi({ bottom: "b", left: "l", right: "r", top: "t" } as const);

export const xsdLegendPosition = bidi({
  bottom: "b",
  topRight: "tr",
  left: "l",
  right: "r",
  top: "t",
} as const);

export const xsdDataLabelPosition = bidi({
  bottom: "b",
  center: "ctr",
  insideBase: "inBase",
  insideEnd: "inEnd",
  left: "l",
  outsideEnd: "outEnd",
  right: "r",
  top: "t",
} as const);

export const xsdAxisCrossBetween = bidi({ middleOfCategory: "midCat" } as const);

export const xsdAxisCrosses = bidi({ zero: "autoZero" } as const);

export const xsdComments = bidi({
  none: "commNone",
  indicator: "commIndicator",
  comment: "commIndAndComment",
} as const);

export const xsdErrorValueType = bidi({
  custom: "cust",
  fixedValue: "fixedVal",
  standardDeviation: "stdDev",
  standardError: "stdErr",
} as const);

export const xsdSizeRepresents = bidi({ width: "w" } as const);

export const xsdSplitType = bidi({ custom: "cust", position: "pos", value: "val" } as const);

export const xsdAxisLabelAlignment = bidi({ center: "ctr", left: "l", right: "r" } as const);

export const xsdAxisOrientation = bidi({ ascending: "minMax", descending: "maxMin" } as const);

export const xsdTrendlineType = bidi({
  exponential: "exp",
  logarithmic: "log",
  movingAverage: "movingAvg",
  polynomial: "poly",
} as const);

// ---------------------------------------------------------------------------
// Chart / text — tab alignment, font alignment, compound line
// ---------------------------------------------------------------------------

export const xsdTextTabAlignment = bidi({
  left: "l",
  center: "ctr",
  right: "r",
  decimal: "dec",
} as const);

export const xsdFontAlignment = bidi({ top: "t", center: "ctr", bottom: "b" } as const);

// ---------------------------------------------------------------------------
// Diagram — animation levels, hierarchy branch, hue direction (dgm:)
// ---------------------------------------------------------------------------

export const xsdAnimLevel = bidi({ level: "lvl", center: "ctr" } as const);

export const xsdHierBranch = bidi({
  left: "l",
  right: "r",
  hanging: "hang",
  standard: "std",
  initial: "init",
} as const);

export const xsdHueDirection = bidi({ clockwise: "cw", counterClockwise: "ccw" } as const);

export const xsdDiagramDirection = bidi({ normal: "norm", reversed: "rev" } as const);

export const xsdResizeHandles = bidi({ exact: "exact", relative: "rel" } as const);

export const xsdOnOffStyle = bidi({ default: "def" } as const);

// ---------------------------------------------------------------------------
// PPTX — animation classes, calc modes, slide layouts, text direction
// ---------------------------------------------------------------------------

export const xsdAnimClass = bidi({
  entrance: "entr",
  emphasis: "emph",
  mediaCall: "mediacall",
} as const);

export const xsdAnimCalcMode = bidi({ linear: "lin", formula: "fmla" } as const);

export const xsdAnimValueType = bidi({ string: "str", number: "num", color: "clr" } as const);

export const xsdTextBuild = bidi({ paragraph: "p", custom: "cust" } as const);

// Diagram build (p:bldDgm @bld, ST_TLDiagramBuildType) — rotation directions
// and "cust" are abbreviated; the rest pass through verbatim.
export const xsdDiagramBuild = bidi({
  clockwise: "cw",
  clockwiseIn: "cwIn",
  clockwiseOut: "cwOut",
  counterclockwise: "ccw",
  counterclockwiseIn: "ccwIn",
  counterclockwiseOut: "ccwOut",
  custom: "cust",
} as const);

// Chart build (p:bldOleChart @bld and a:bldChart @bld — identical token sets,
// ST_TLOleChartBuildType / ST_AnimationChartBuildType).
export const xsdChartBuild = bidi({
  seriesElement: "seriesEl",
  categoryElement: "categoryEl",
} as const);

// Diagram sub-build (a:bldDgm @bld, ST_AnimationDgmBuildType).
export const xsdAnimationDgmBuild = bidi({
  levelOne: "lvlOne",
  levelAtOnce: "lvlAtOnce",
} as const);

export const xsdIterateType = bidi({ element: "el", word: "wd", letter: "lt" } as const);

export const xsdSlideLayoutType = bidi({
  text: "tx",
  twoColumnText: "twoColTx",
  object: "obj",
  sectionHeader: "secHead",
  table: "tbl",
  clipArtAndText: "clipArtAndTx",
  pictureText: "picTx",
  twoObjects: "twoObj",
  twoTextAndTwoObjects: "twoTxTwoObj",
  objectAndText: "objTx",
  verticalText: "vertTx",
  verticalTitleAndText: "vertTitleAndTx",
} as const);

export const xsdTextVerticalType = bidi({
  horizontal: "horz",
  vertical: "vert",
  vertical270: "vert270",
  wordArtVertical: "wordArtVert",
  eastAsianVertical: "eaVert",
  mongolianVertical: "mongolianVert",
  wordArtVerticalRightToLeft: "wordArtVertRtl",
} as const);

export const xsdDashStyle = bidi({
  longDash: "lgDash",
  systemDot: "sysDot",
  systemDash: "sysDash",
} as const);

export const xsdOrient = bidi({ horizontal: "horz", vertical: "vert" } as const);

// ST_PlaceholderType (p:ph/@type) — identity pairs title/body/subTitle/chart/
// clipArt/media are omitted.
export const xsdPlaceholderType = bidi({
  centerTitle: "ctrTitle",
  date: "dt",
  slideNumber: "sldNum",
  footer: "ftr",
  header: "hdr",
  object: "obj",
  table: "tbl",
  diagram: "dgm",
  slideImage: "sldImg",
  picture: "pic",
} as const);

// ---------------------------------------------------------------------------
// DOCX / XLSX — math run style, table width unit, totals-row aggregates
// ---------------------------------------------------------------------------

export const xsdMathStyle = bidi({ plain: "p", bold: "b", italic: "i", boldItalic: "bi" } as const);

export const xsdTableWidthType = bidi({ twips: "dxa", percent: "pct" } as const);

// ---------------------------------------------------------------------------
// DOCX — paragraph/table justification (ST_Jc*) and shading pattern (ST_Shd)
// ---------------------------------------------------------------------------

export const xsdJcAlignment = bidi({ numericTab: "numTab" } as const);

export const xsdConsolidateFunction = bidi({
  countNumbers: "countNums",
  standardDeviation: "stdDev",
  standardDeviationPopulation: "stdDevp",
  variance: "var",
  variancePopulation: "varp",
} as const);

export const xsdShadingPattern = bidi({
  horizontalStripe: "horzStripe",
  verticalStripe: "vertStripe",
  reverseDiagonalStripe: "reverseDiagStripe",
  diagonalStripe: "diagStripe",
  horizontalCross: "horzCross",
  diagonalCross: "diagCross",
  thinHorizontalStripe: "thinHorzStripe",
  thinVerticalStripe: "thinVertStripe",
  thinReverseDiagonalStripe: "thinReverseDiagStripe",
  thinDiagonalStripe: "thinDiagStripe",
  thinHorizontalCross: "thinHorzCross",
  thinDiagonalCross: "thinDiagCross",
  percent5: "pct5",
  percent10: "pct10",
  percent12: "pct12",
  percent15: "pct15",
  percent20: "pct20",
  percent25: "pct25",
  percent30: "pct30",
  percent35: "pct35",
  percent37: "pct37",
  percent40: "pct40",
  percent45: "pct45",
  percent50: "pct50",
  percent55: "pct55",
  percent60: "pct60",
  percent62: "pct62",
  percent65: "pct65",
  percent70: "pct70",
  percent75: "pct75",
  percent80: "pct80",
  percent85: "pct85",
  percent87: "pct87",
  percent90: "pct90",
  percent95: "pct95",
} as const);

export const xsdTotalsRowFunction = bidi({
  countNumbers: "countNums",
  standardDeviation: "stdDev",
  variance: "var",
} as const);
