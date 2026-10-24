/**
 * Settings descriptor — produces word/settings.xml.
 *
 * Stringifies pure JSON SettingsOptions into XML without creating
 * Settings/Compatibility/DocumentProtection/etc. class instances.
 * Zero XmlComponent allocation — pure string concatenation.
 *
 * @module
 */

import { derivePasswordHash, parseOnOff } from "@office-open/core";
import type { Percentage, PositiveUniversalMeasure } from "@office-open/core";
import { twipsMeasureValue } from "@office-open/core";
import {
  parseVmlShapeDefaults,
  parseVmlShapeLayout,
  stringifyVmlShapeDefaults,
  stringifyVmlShapeLayout,
} from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { RELATIONSHIP_TYPES } from "@office-open/core/opc";
import { attr, attrMeasure, escapeXml, findChild, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { DocxParseError } from "../../errors";
import type { DocumentOptions } from "../core-properties";
import {
  documentNamespaceAttributesInDialect,
  documentNamespaceDialect,
} from "../document/document-attributes";
import type { OnOffLexicalValue } from "../paragraph/run/properties";
import type {
  SettingsOptions,
  DocumentProtectionOptions,
  WriteProtectionOptions,
  MailMergeOptions,
  OdsoOptions,
  OdsoFieldMapDataOptions,
  MailMergeRecipientsOptions,
  RecipientDataOptions,
  SettingsOnOffValue,
  CompatibilityOptions,
  CompatSettingOptions,
  CaptionsOptions,
  CaptionOptions,
  AutoCaptionOptions,
  MathPropertiesOptions,
  RsidsOptions,
  ReadModeInkLockDownOptions,
  RevisionViewOptions,
  TmCommentsPropertiesOptions,
  TmLastPositionOptions,
  TmReviewPropertiesOptions,
  DocumentFootnotePropertiesOptions,
  DocumentEndnotePropertiesOptions,
  ShapeDefaultsOptions,
} from "./settings";

const SETTINGS_VIEWS = ["none", "print", "outline", "masterPages", "normal", "web"] as const;

const SETTINGS_CHARACTER_SPACING = [
  "compressPunctuation",
  "compressPunctuationAndJapaneseKana",
  "doNotCompress",
] as const;

const POSITIVE_MEASURE_PATTERN = /^\d+(?:\.\d+)?(?:mm|cm|in|pt|pc|pi)$/;

function isSettingsView(value: string): value is (typeof SETTINGS_VIEWS)[number] {
  return (SETTINGS_VIEWS as readonly string[]).includes(value);
}

function isCharacterSpacingControl(
  value: string,
): value is (typeof SETTINGS_CHARACTER_SPACING)[number] {
  return (SETTINGS_CHARACTER_SPACING as readonly string[]).includes(value);
}

function assignSetting<K extends keyof SettingsOptions>(
  options: SettingsOptions,
  key: K,
  value: SettingsOptions[K],
): void {
  if (value !== undefined) options[key] = value;
}

function parsePositiveMeasure(value: string | number): number | PositiveUniversalMeasure {
  if (typeof value === "number") {
    if (value >= 0) return value;
  } else if (POSITIVE_MEASURE_PATTERN.test(value)) {
    return value as PositiveUniversalMeasure;
  }
  throw new DocxParseError(
    `Invalid positive measurement "${value}".`,
    "word/settings.xml",
    "w:defaultTabStop",
    "invalid-measure",
  );
}

/** Derive the namespace-prefixed val attribute from the element tag. */
function valAttr(tag: string): string {
  const ns = tag.split(":")[0];
  return `${ns}:val`;
}

function onOff(tag: string, val: SettingsOnOffValue | undefined): string {
  if (val === undefined) return "";
  if (typeof val === "boolean") return val ? `<${tag}/>` : `<${tag} ${valAttr(tag)}="0"/>`;
  return `<${tag} ${valAttr(tag)}="${val}"/>`;
}

function numVal(tag: string, val: number | undefined): string {
  return val !== undefined ? `<${tag} ${valAttr(tag)}="${val}"/>` : "";
}

function measureVal(tag: string, val: number | PositiveUniversalMeasure | undefined): string {
  if (val === undefined) return "";
  const value = typeof val === "string" ? val : twipsMeasureValue(val);
  return `<${tag} ${valAttr(tag)}="${value}"/>`;
}

function strVal(tag: string, val: string | undefined): string {
  return val !== undefined ? `<${tag} ${valAttr(tag)}="${escapeXml(val)}"/>` : "";
}

function docIdVal(tag: string, val: string | undefined, text: string | undefined): string {
  if (val === undefined) return "";
  const inline = text === undefined ? "" : escapeXml(text);
  return `<${tag} ${valAttr(tag)}="${escapeXml(val)}">${inline}</${tag}>`;
}

function stringifyTmCommentsProperties(options: TmCommentsPropertiesOptions): string {
  return `<w:tmCommentsPr>${numVal("w:tmCommentsPlace", options.place)}${numVal(
    "w:tmCommentsWidth",
    options.width,
  )}${numVal("w:tmCommentsColor", options.color)}</w:tmCommentsPr>`;
}

function stringifyTmReviewProperties(options: TmReviewPropertiesOptions): string {
  const children: Array<[string, number | undefined]> = [
    ["w:tmReviewEnabled", options.enabled],
    ["w:tmReviewShow", options.show],
    ["w:tmReviewPrint", options.print],
    ["w:tmRevisionNum", options.revisionNumber],
    ["w:tmReviewMarkIns", options.insertionMark],
    ["w:tmReviewColorIns", options.insertionColor],
    ["w:tmReviewMarkDel", options.deletionMark],
    ["w:tmReviewColorDel", options.deletionColor],
    ["w:tmReviewMarkFmt", options.formattingMark],
    ["w:tmReviewColorFmt", options.formattingColor],
    ["w:tmReviewMarkLn", options.lineMark],
    ["w:tmReviewColorLn", options.lineColor],
    ["w:tmReviewToolTip", options.tooltip],
  ];
  return `<w:tmReviewPr>${children
    .map(([tag, value]) => numVal(tag, value))
    .join("")}</w:tmReviewPr>`;
}

function stringifyTmPosition(
  tag: string,
  options: { paragraphIndex?: number; index?: number } | undefined,
): string {
  if (!options) return "";
  return `<${tag}>${numVal("w:tmLastPosPgfIdx", options.paragraphIndex)}${numVal(
    "w:tmLastPosIdx",
    options.index,
  )}</${tag}>`;
}

function stringifyTmLastPosition(options: TmLastPositionOptions): string {
  const tableRectangle =
    options.tableRectangle === undefined
      ? ""
      : attrEl("w:tmLastPosTblRect", {
          "w:left": options.tableRectangle.left,
          "w:top": options.tableRectangle.top,
          "w:right": options.tableRectangle.right,
          "w:bottom": options.tableRectangle.bottom,
        });
  return `<w:tmLastPos>${numVal("w:tmLastPosPage", options.page)}${numVal(
    "w:tmLastPosSelect",
    options.selection,
  )}${numVal("w:tmLastPosFrameIdx", options.frameIndex)}${stringifyTmPosition(
    "w:tmLastPosCaret",
    options.caret,
  )}${stringifyTmPosition("w:tmLastPosAnchor", options.anchor)}${tableRectangle}</w:tmLastPos>`;
}

function parseTmCommentsProperties(el: Element): TmCommentsPropertiesOptions {
  return {
    place: readNum(findChild(el, "w:tmCommentsPlace"), "w:val"),
    width: readNum(findChild(el, "w:tmCommentsWidth"), "w:val"),
    color: readNum(findChild(el, "w:tmCommentsColor"), "w:val"),
  };
}

function parseTmReviewProperties(el: Element): TmReviewPropertiesOptions {
  const value = (tag: string): number | undefined => readNum(findChild(el, tag), "w:val");
  return {
    enabled: value("w:tmReviewEnabled"),
    show: value("w:tmReviewShow"),
    print: value("w:tmReviewPrint"),
    revisionNumber: value("w:tmRevisionNum"),
    insertionMark: value("w:tmReviewMarkIns"),
    insertionColor: value("w:tmReviewColorIns"),
    deletionMark: value("w:tmReviewMarkDel"),
    deletionColor: value("w:tmReviewColorDel"),
    formattingMark: value("w:tmReviewMarkFmt"),
    formattingColor: value("w:tmReviewColorFmt"),
    lineMark: value("w:tmReviewMarkLn"),
    lineColor: value("w:tmReviewColorLn"),
    tooltip: value("w:tmReviewToolTip"),
  };
}

function parseTmPosition(el: Element | undefined): { paragraphIndex?: number; index?: number } {
  return {
    paragraphIndex: readNum(findChild(el, "w:tmLastPosPgfIdx"), "w:val"),
    index: readNum(findChild(el, "w:tmLastPosIdx"), "w:val"),
  };
}

function parseTmLastPosition(el: Element): TmLastPositionOptions {
  const rect = findChild(el, "w:tmLastPosTblRect");
  const position = (tag: string) => parseTmPosition(findChild(el, tag));
  return {
    page: readNum(findChild(el, "w:tmLastPosPage"), "w:val"),
    selection: readNum(findChild(el, "w:tmLastPosSelect"), "w:val"),
    frameIndex: readNum(findChild(el, "w:tmLastPosFrameIdx"), "w:val"),
    caret: position("w:tmLastPosCaret"),
    anchor: position("w:tmLastPosAnchor"),
    tableRectangle: rect
      ? {
          left: readNum(rect, "w:left"),
          top: readNum(rect, "w:top"),
          right: readNum(rect, "w:right"),
          bottom: readNum(rect, "w:bottom"),
        }
      : undefined,
  };
}

/** Build attribute string from key-value pairs, skipping undefined. */
function attrStr(attrs: Record<string, string | number | boolean | undefined>): string {
  const parts: string[] = [];
  for (const [k, v] of Object.entries(attrs)) {
    if (v !== undefined) parts.push(`${k}="${escapeXml(String(v))}"`);
  }
  return parts.join(" ");
}

/** Self-closing element with attributes only. */
function attrEl(tag: string, attrs: Record<string, string | number | boolean | undefined>): string {
  const a = attrStr(attrs);
  return a ? `<${tag} ${a}/>` : `<${tag}/>`;
}

function compatSetting(name: string, val: string | number, uri?: string): string {
  const u = uri ?? "http://schemas.microsoft.com/office/word";
  return `<w:compatSetting w:name="${escapeXml(name)}" w:uri="${u}" w:val="${val}"/>`;
}

// ── Parse helpers ──

/** Read a CT_OnOff child as boolean (presence true unless val is explicitly false). */
function readOnOffValue(el: Element | undefined): SettingsOnOffValue | undefined {
  if (!el || !el.name) return undefined;
  const v = attr(el, valAttr(el.name));
  return parseOnOff(v) === undefined ? true : (v as SettingsOnOffValue);
}

function readOnOff(el: Element | undefined): boolean | undefined {
  const value = readOnOffValue(el);
  return typeof value === "string" ? parseOnOff(value) : value;
}

/** Read an attribute as a number, or undefined if absent/unparseable. */
function readNum(el: Element | undefined, name: string): number | undefined {
  if (!el) return undefined;
  const v = attr(el, name);
  if (v === undefined || v === "") return undefined;
  const n = parseInt(v, 10);
  return Number.isNaN(n) ? undefined : n;
}

function readPositiveMeasure(
  el: Element | undefined,
): number | PositiveUniversalMeasure | undefined {
  const value = attrMeasure(el, "m:val");
  return typeof value === "number" || typeof value === "string"
    ? (value as number | PositiveUniversalMeasure)
    : undefined;
}

/** Read an attribute as a string, or undefined if absent. */
function readStr(el: Element | undefined, name: string): string | undefined {
  if (!el) return undefined;
  const v = attr(el, name);
  return v === undefined || v === "" ? undefined : v;
}

/** Read an optional string while preserving an explicitly empty `w:val=""`. */
function readOptionalStr(el: Element | undefined, name: string): string | undefined {
  return el ? attr(el, name) : undefined;
}

/** Read an attribute constrained to an enum; undefined if absent or not allowed. */
function readEnum<T extends string>(
  el: Element | undefined,
  name: string,
  allowed: readonly T[],
): T | undefined {
  const v = readStr(el, name);
  return v !== undefined && (allowed as readonly string[]).includes(v) ? (v as T) : undefined;
}

/** Shared password/crypto attributes (AG_TransitionalPassword) → options keys. */
const PASSWORD_ATTR_MAP: [string, string, boolean][] = [
  ["hashValue", "w:hashValue", false],
  ["saltValue", "w:saltValue", false],
  ["hash", "w:hash", false],
  ["salt", "w:salt", false],
  ["spinCount", "w:spinCount", true],
  ["algorithmName", "w:algorithmName", false],
  ["cryptoAlgorithmClass", "w:cryptAlgorithmClass", false],
  ["cryptoAlgorithmSid", "w:cryptAlgorithmSid", true],
  ["cryptoAlgorithmType", "w:cryptAlgorithmType", false],
  ["cryptoProvider", "w:cryptProvider", false],
  ["cryptoProviderType", "w:cryptProviderType", false],
  ["cryptoProviderTypeExtension", "w:cryptProviderTypeExt", true],
  ["cryptoProviderTypeExtensionSource", "w:cryptProviderTypeExtSource", false],
  ["algorithmExtensionId", "w:algIdExt", true],
  ["algorithmExtensionSource", "w:algIdExtSource", false],
  ["cryptoSpinCount", "w:cryptSpinCount", true],
];

/** Read shared password/crypto attributes into an options object. */
function readPasswordAttrs(el: Element): Record<string, string | number> {
  const out: Record<string, string | number> = {};
  for (const [key, xmlAttr, isNum] of PASSWORD_ATTR_MAP) {
    const v = attr(el, xmlAttr);
    if (v === undefined || v === "") continue;
    out[key] = isNum ? parseInt(v, 10) : v;
  }
  return out;
}

/**
 * CT_OnOff compat flag elements → CompatibilityOptions keys. The XML tag often
 * differs from the option key (e.g. wordPerfectJustification → w:wpJustification).
 * Order mirrors stringifyCompatibility so round-trip preserves element order.
 */
const COMPAT_FLAG_MAP: [string, string][] = [
  ["useSingleBorderforContiguousCells", "w:useSingleBorderforContiguousCells"],
  ["wordPerfectJustification", "w:wpJustification"],
  ["noTabStopForHangingIndent", "w:noTabHangInd"],
  ["noLeading", "w:noLeading"],
  ["spaceForUnderline", "w:spaceForUL"],
  ["noColumnBalance", "w:noColumnBalance"],
  ["balanceSingleByteDoubleByteWidth", "w:balanceSingleByteDoubleByteWidth"],
  ["noExtraLineSpacing", "w:noExtraLineSpacing"],
  ["doNotLeaveBackslashAlone", "w:doNotLeaveBackslashAlone"],
  ["underlineTrailingSpaces", "w:ulTrailSpace"],
  ["doNotExpandShiftReturn", "w:doNotExpandShiftReturn"],
  ["spacingInWholePoints", "w:spacingInWholePoints"],
  ["lineWrapLikeWord6", "w:lineWrapLikeWord6"],
  ["printBodyTextBeforeHeader", "w:printBodyTextBeforeHeader"],
  ["printColorsBlack", "w:printColBlack"],
  ["spaceWidth", "w:wpSpaceWidth"],
  ["showBreaksInFrames", "w:showBreaksInFrames"],
  ["subFontBySize", "w:subFontBySize"],
  ["suppressBottomSpacing", "w:suppressBottomSpacing"],
  ["suppressTopSpacing", "w:suppressTopSpacing"],
  ["suppressSpacingAtTopOfPage", "w:suppressSpacingAtTopOfPage"],
  ["suppressTopSpacingWP", "w:suppressTopSpacingWP"],
  ["suppressSpBfAfterPgBrk", "w:suppressSpBfAfterPgBrk"],
  ["swapBordersFacingPages", "w:swapBordersFacingPages"],
  ["convertMailMergeEsc", "w:convMailMergeEsc"],
  ["truncateFontHeightsLikeWP6", "w:truncateFontHeightsLikeWP6"],
  ["macWordSmallCaps", "w:mwSmallCaps"],
  ["usePrinterMetrics", "w:usePrinterMetrics"],
  ["doNotSuppressParagraphBorders", "w:doNotSuppressParagraphBorders"],
  ["wrapTrailSpaces", "w:wrapTrailSpaces"],
  ["footnoteLayoutLikeWW8", "w:footnoteLayoutLikeWW8"],
  ["shapeLayoutLikeWW8", "w:shapeLayoutLikeWW8"],
  ["alignTablesRowByRow", "w:alignTablesRowByRow"],
  ["forgetLastTabAlignment", "w:forgetLastTabAlignment"],
  ["adjustLineHeightInTable", "w:adjustLineHeightInTable"],
  ["autoSpaceLikeWord95", "w:autoSpaceLikeWord95"],
  ["noSpaceRaiseLower", "w:noSpaceRaiseLower"],
  ["doNotUseHTMLParagraphAutoSpacing", "w:doNotUseHTMLParagraphAutoSpacing"],
  ["layoutRawTableWidth", "w:layoutRawTableWidth"],
  ["layoutTableRowsApart", "w:layoutTableRowsApart"],
  ["useWord97LineBreakRules", "w:useWord97LineBreakRules"],
  ["doNotBreakWrappedTables", "w:doNotBreakWrappedTables"],
  ["doNotSnapToGridInCell", "w:doNotSnapToGridInCell"],
  ["selectFieldWithFirstOrLastCharacter", "w:selectFldWithFirstOrLastChar"],
  ["applyBreakingRules", "w:applyBreakingRules"],
  ["doNotWrapTextWithPunctuation", "w:doNotWrapTextWithPunct"],
  ["doNotUseEastAsianBreakRules", "w:doNotUseEastAsianBreakRules"],
  ["useWord2002TableStyleRules", "w:useWord2002TableStyleRules"],
  ["growAutofit", "w:growAutofit"],
  ["useFELayout", "w:useFELayout"],
  ["useNormalStyleForList", "w:useNormalStyleForList"],
  ["doNotUseIndentAsNumberingTabStop", "w:doNotUseIndentAsNumberingTabStop"],
  ["useAlternateEastAsianLineBreakRules", "w:useAltKinsokuLineBreakRules"],
  ["allowSpaceOfSameStyleInTable", "w:allowSpaceOfSameStyleInTable"],
  ["doNotSuppressIndentation", "w:doNotSuppressIndentation"],
  ["doNotAutofitConstrainedTables", "w:doNotAutofitConstrainedTables"],
  ["autofitToFirstFixedWidthCell", "w:autofitToFirstFixedWidthCell"],
  ["underlineTabInNumberingList", "w:underlineTabInNumList"],
  ["displayHangulFixedWidth", "w:displayHangulFixedWidth"],
  ["splitPgBreakAndParaMark", "w:splitPgBreakAndParaMark"],
  ["doNotVerticallyAlignCellWithSp", "w:doNotVertAlignCellWithSp"],
  ["doNotBreakConstrainedForcedTable", "w:doNotBreakConstrainedForcedTable"],
  ["ignoreVerticalAlignmentInTextboxes", "w:doNotVertAlignInTxbx"],
  ["useAnsiKerningPairs", "w:useAnsiKerningPairs"],
  ["cachedColumnBalance", "w:cachedColBalance"],
];

/** compatSetting names that map to dedicated sugar fields (not into compatSettings[]). */
const COMPAT_SETTING_SUGAR: Record<
  string,
  | "version"
  | "overrideTableStyleFontSizeAndJustification"
  | "enableOpenTypeFeatures"
  | "doNotFlipMirrorIndents"
> = {
  compatibilityMode: "version",
  overrideTableStyleFontSizeAndJustification: "overrideTableStyleFontSizeAndJustification",
  enableOpenTypeFeatures: "enableOpenTypeFeatures",
  doNotFlipMirrorIndents: "doNotFlipMirrorIndents",
};

/** Parse w:footnotePr / w:endnotePr (CT_FtnDocProps / CT_EdnDocProps). */
function parseFtnEdnPr(
  el: Element,
  refTag: "w:footnote" | "w:endnote",
): Record<string, unknown> | undefined {
  const o: Record<string, unknown> = {};
  const pos = readStr(findChild(el, "w:pos"), "w:val");
  if (pos) o.pos = pos;
  const numFmtEl = findChild(el, "w:numFmt");
  if (numFmtEl) {
    const v = readStr(numFmtEl, "w:val");
    if (v) o.numFmt = v;
    const format = readStr(numFmtEl, "w:format");
    if (format) o.format = format;
  }
  const numStart = readNum(findChild(el, "w:numStart"), "w:val");
  if (numStart !== undefined) o.numStart = numStart;
  const numRestart = readStr(findChild(el, "w:numRestart"), "w:val");
  if (numRestart) o.numRestart = numRestart;
  // Special-note references (separator -1 / continuation 0) — the element can
  // carry these alone, so their presence alone keeps the properties object.
  const refs = (el.elements ?? [])
    .filter((c) => c.type === "element" && c.name === refTag)
    .map((c) => readNum(c, "w:id"))
    .filter((id): id is number => id !== undefined);
  if (refs.length > 0) o[refTag === "w:footnote" ? "footnotes" : "endnotes"] = refs;
  return Object.keys(o).length > 0 ? o : undefined;
}

/** Apply one w:compatSetting entry to a compatibility options object: named
 *  sugar fields first, remaining entries accumulate into compatSettings[]. */
type CompatibilityAccumulator = { [K in keyof CompatibilityOptions]: unknown };

function applyCompatSetting(target: CompatibilityAccumulator, child: Element): void {
  const name = attr(child, "w:name");
  const val = attr(child, "w:val");
  if (name === undefined || val === undefined) return;
  const extras = (target.compatSettings as CompatSettingOptions[] | undefined) ?? [];
  extras.push({ name, val, uri: attr(child, "w:uri") });
  target.compatSettings = extras;
  const sugar = COMPAT_SETTING_SUGAR[name];
  if (sugar === "version") {
    const n = parseInt(val, 10);
    if (!Number.isNaN(n)) target.version = n;
  } else if (sugar !== undefined && target[sugar] === undefined) {
    target[sugar] = parseOnOff(val) ?? true;
  }
}

/** Parse w:compat (CT_Compat): on/off flag elements + w:compatSetting entries. */
function parseCompatibility(el: Element): CompatibilityOptions | undefined {
  const o: Record<string, unknown> = {};
  const flagMap: Record<string, string> = Object.fromEntries(
    COMPAT_FLAG_MAP.map(([key, tag]) => [tag, key]),
  );
  for (const child of el.elements ?? []) {
    if (child.type !== "element" || !child.name) continue;
    const key = flagMap[child.name];
    if (key !== undefined) {
      o[key] = readOnOff(child);
      continue;
    }
    if (child.name === "w:compatSetting") {
      applyCompatSetting(o, child);
    } else if (child.name === "w14:enableOpenTypeKerning") {
      o.enableOpenTypeKerning = parseOnOff(attr(child, "w14:val")) ?? true;
    }
  }
  // Presence-based: a bare `<w:compat/>` (no flags) round-trips as an empty
  // object — returning undefined here would drop the element and let the
  // fresh-document defaults get injected over it.
  return o as CompatibilityOptions;
}

/** Parse m:mathPr (CT_MathPr). */
function parseMathPr(el: Element): MathPropertiesOptions | undefined {
  const o: Partial<MathPropertiesOptions> = {};
  const mathFont = readStr(findChild(el, "m:mathFont"), "m:val");
  if (mathFont) o.mathFont = mathFont;
  const binaryOperatorBreak = readEnum(findChild(el, "m:brkBin"), "m:val", [
    "before",
    "after",
    "repeat",
  ] as const);
  if (binaryOperatorBreak) o.binaryOperatorBreak = binaryOperatorBreak;
  const binaryOperatorBreakSubtraction = readEnum(findChild(el, "m:brkBinSub"), "m:val", [
    "--",
    "-+",
    "+-",
  ] as const);
  if (binaryOperatorBreakSubtraction)
    o.binaryOperatorBreakSubtraction = binaryOperatorBreakSubtraction;
  const smallFractions = readOnOff(findChild(el, "m:smallFrac"));
  if (smallFractions !== undefined) o.smallFractions = smallFractions;
  const displayDefaults = readOnOff(findChild(el, "m:dispDef"));
  if (displayDefaults !== undefined) o.displayDefaults = displayDefaults;
  const leftMargin = readPositiveMeasure(findChild(el, "m:lMargin"));
  if (leftMargin !== undefined) o.leftMargin = leftMargin;
  const rightMargin = readPositiveMeasure(findChild(el, "m:rMargin"));
  if (rightMargin !== undefined) o.rightMargin = rightMargin;
  const defaultJustification = readEnum(findChild(el, "m:defJc"), "m:val", [
    "left",
    "right",
    "center",
    "centerGroup",
  ] as const);
  if (defaultJustification) o.defaultJustification = defaultJustification;
  const preSpacing = readPositiveMeasure(findChild(el, "m:preSp"));
  if (preSpacing !== undefined) o.preSpacing = preSpacing;
  const postSpacing = readPositiveMeasure(findChild(el, "m:postSp"));
  if (postSpacing !== undefined) o.postSpacing = postSpacing;
  const interSpacing = readPositiveMeasure(findChild(el, "m:interSp"));
  if (interSpacing !== undefined) o.interSpacing = interSpacing;
  const intraSpacing = readPositiveMeasure(findChild(el, "m:intraSp"));
  if (intraSpacing !== undefined) o.intraSpacing = intraSpacing;
  const wrapIndent = readPositiveMeasure(findChild(el, "m:wrapIndent"));
  if (wrapIndent !== undefined) o.wrapIndent = wrapIndent;
  const wrapRight = readOnOff(findChild(el, "m:wrapRight"));
  if (wrapRight !== undefined) o.wrapRight = wrapRight;
  const integralLimitLocation = readEnum(findChild(el, "m:intLim"), "m:val", [
    "subSup",
    "undOvr",
  ] as const);
  if (integralLimitLocation) o.integralLimitLocation = integralLimitLocation;
  const naryLimitLocation = readEnum(findChild(el, "m:naryLim"), "m:val", [
    "subSup",
    "undOvr",
  ] as const);
  if (naryLimitLocation) o.naryLimitLocation = naryLimitLocation;
  return Object.keys(o).length > 0 ? (o as MathPropertiesOptions) : undefined;
}

/** Parse w:captions (CT_Captions). */
function parseCaptions(el: Element): CaptionsOptions | undefined {
  const captions: CaptionOptions[] = [];
  const autoCaptions: AutoCaptionOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === "w:caption") {
      const c: CaptionOptions = { name: attr(child, "w:name") ?? "" };
      const pos = attr(child, "w:pos");
      if (pos) c.pos = pos as CaptionOptions["pos"];
      const chapNum = attr(child, "w:chapNum");
      if (chapNum !== undefined) c.chapterNumber = parseOnOff(chapNum) ?? false;
      const heading = attr(child, "w:heading");
      if (heading !== undefined) c.heading = parseInt(heading, 10);
      const noLabel = attr(child, "w:noLabel");
      if (noLabel !== undefined) c.noLabel = parseOnOff(noLabel) ?? false;
      const numFmt = attr(child, "w:numFmt");
      if (numFmt) c.numFmt = numFmt as CaptionOptions["numFmt"];
      const sep = attr(child, "w:sep");
      if (sep) c.sep = sep as CaptionOptions["sep"];
      captions.push(c);
    } else if (child.name === "w:autoCaptions") {
      for (const ac of child.elements ?? []) {
        if (ac.name !== "w:autoCaption") continue;
        const name = attr(ac, "w:name");
        const caption = attr(ac, "w:caption");
        if (name && caption) autoCaptions.push({ name, caption });
      }
    }
  }
  if (captions.length === 0) return undefined;
  const o: CaptionsOptions = { captions };
  if (autoCaptions.length > 0) o.autoCaptions = autoCaptions;
  return o;
}

/** Parse w:odso (CT_Odso). */
function parseOdso(el: Element): OdsoOptions | undefined {
  const o: OdsoOptions = {};
  const udl = readStr(findChild(el, "w:udl"), "w:val");
  if (udl) o.udl = udl;
  const table = readStr(findChild(el, "w:table"), "w:val");
  if (table) o.table = table;
  const srcEl = findChild(el, "w:src");
  if (srcEl) {
    const rid = attr(srcEl, "r:id");
    if (rid) o.src = rid;
  }
  const colDelim = readNum(findChild(el, "w:colDelim"), "w:val");
  if (colDelim !== undefined) o.colDelim = colDelim;
  const type = readStr(findChild(el, "w:type"), "w:val");
  if (type) o.type = type as OdsoOptions["type"];
  const fHdr = readOnOff(findChild(el, "w:fHdr"));
  if (fHdr !== undefined) o.fHdr = fHdr;
  const fieldMapData: OdsoFieldMapDataOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "w:fieldMapData") continue;
    const fm: OdsoFieldMapDataOptions = {};
    const t = readStr(findChild(child, "w:type"), "w:val");
    if (t) fm.type = t as OdsoFieldMapDataOptions["type"];
    const n = readStr(findChild(child, "w:name"), "w:val");
    if (n) fm.name = n;
    const mn = readStr(findChild(child, "w:mappedName"), "w:val");
    if (mn) fm.mappedName = mn;
    const col = readNum(findChild(child, "w:column"), "w:val");
    if (col !== undefined) fm.column = col;
    const lid = readStr(findChild(child, "w:lid"), "w:val");
    if (lid) fm.lid = lid;
    const dyn = readOnOff(findChild(child, "w:dynamicAddress"));
    if (dyn !== undefined) fm.dynamicAddress = dyn;
    if (Object.keys(fm).length > 0) fieldMapData.push(fm);
  }
  if (fieldMapData.length > 0) o.fieldMapData = fieldMapData;
  const recipientData: string[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "w:recipientData") continue;
    const rid = attr(child, "r:id");
    if (rid) recipientData.push(rid);
  }
  if (recipientData.length > 0) o.recipientData = recipientData;
  return Object.keys(o).length > 0 ? o : undefined;
}

/** Parse w:mailMerge (CT_MailMerge). */
function parseMailMerge(el: Element): MailMergeOptions | undefined {
  const o: Partial<MailMergeOptions> = {};
  const mdt = readStr(findChild(el, "w:mainDocumentType"), "w:val");
  if (mdt) o.mainDocumentType = mdt as MailMergeOptions["mainDocumentType"];
  const dataType = readStr(findChild(el, "w:dataType"), "w:val");
  if (dataType) o.dataType = dataType as MailMergeOptions["dataType"];
  const dest = readStr(findChild(el, "w:destination"), "w:val");
  if (dest) o.destination = dest as MailMergeOptions["destination"];
  const connectString = readOptionalStr(findChild(el, "w:connectString"), "w:val");
  if (connectString !== undefined) o.connectString = connectString;
  const query = readStr(findChild(el, "w:query"), "w:val");
  if (query) o.query = query;
  const dsEl = findChild(el, "w:dataSource");
  if (dsEl) {
    const rid = attr(dsEl, "r:id");
    if (rid) o.dataSource = rid;
  }
  const hsEl = findChild(el, "w:headerSource");
  if (hsEl) {
    const rid = attr(hsEl, "r:id");
    if (rid) o.headerSource = rid;
  }
  const linkToQuery = readOnOff(findChild(el, "w:linkToQuery"));
  if (linkToQuery !== undefined) o.linkToQuery = linkToQuery;
  const doNotSuppress = readOnOff(findChild(el, "w:doNotSuppressBlankLines"));
  if (doNotSuppress !== undefined) o.doNotSuppressBlankLines = doNotSuppress;
  const addressFieldName = readStr(findChild(el, "w:addressFieldName"), "w:val");
  if (addressFieldName) o.addressFieldName = addressFieldName;
  const mailSubject = readStr(findChild(el, "w:mailSubject"), "w:val");
  if (mailSubject) o.mailSubject = mailSubject;
  const mailAsAttachment = readOnOff(findChild(el, "w:mailAsAttachment"));
  if (mailAsAttachment !== undefined) o.mailAsAttachment = mailAsAttachment;
  const viewMergedData = readOnOff(findChild(el, "w:viewMergedData"));
  if (viewMergedData !== undefined) o.viewMergedData = viewMergedData;
  const activeRecord = readNum(findChild(el, "w:activeRecord"), "w:val");
  if (activeRecord !== undefined) o.activeRecord = activeRecord;
  const checkErrors = readNum(findChild(el, "w:checkErrors"), "w:val");
  if (checkErrors !== undefined) o.checkErrors = checkErrors;
  const odsoEl = findChild(el, "w:odso");
  if (odsoEl) {
    const odso = parseOdso(odsoEl);
    if (odso) o.odso = odso;
  }
  if (Object.keys(o).length === 0) return undefined;
  return o as MailMergeOptions;
}

// ── Password derivation ──

interface DerivedPassword {
  hashValue: string;
  saltValue: string;
  spinCount: number;
  algorithmName: string;
}

function maybeDerive(
  password: string | undefined,
  hashValue: string | undefined,
): DerivedPassword | undefined {
  return password !== undefined && hashValue === undefined
    ? derivePasswordHash(password)
    : undefined;
}

// ── Complex sub-elements ──

/** Valid w:documentProtection/`@w:edit` values (ST_DocProtect). */
const DOC_PROTECT_EDITS = ["none", "readOnly", "comments", "trackedChanges", "forms"] as const;

function stringifyDocProtect(opts: DocumentProtectionOptions): string {
  const derived = maybeDerive(opts.password, opts.hashValue);
  const enforcement =
    opts.enforcementRaw ??
    (opts.enforcement === null ? undefined : opts.enforcement === false ? "0" : "1");
  const attrs: Record<string, string | number> = {};
  if (enforcement !== undefined) attrs["w:enforcement"] = enforcement;
  if (opts.edit !== undefined) attrs["w:edit"] = opts.edit;
  if (opts.formatting !== undefined) attrs["w:formatting"] = opts.formatting ? "1" : "0";
  if (opts.algorithmName ?? derived?.algorithmName)
    attrs["w:algorithmName"] = opts.algorithmName ?? derived!.algorithmName;
  if (opts.hashValue ?? derived?.hashValue)
    attrs["w:hashValue"] = opts.hashValue ?? derived!.hashValue;
  if (opts.saltValue ?? derived?.saltValue)
    attrs["w:saltValue"] = opts.saltValue ?? derived!.saltValue;
  if (opts.cryptoAlgorithmSidRaw !== undefined)
    attrs["w:cryptAlgorithmSid"] = opts.cryptoAlgorithmSidRaw;
  if (opts.spinCount ?? derived?.spinCount)
    attrs["w:spinCount"] = opts.spinCount ?? derived!.spinCount;
  if (opts.cryptoAlgorithmClass !== undefined)
    attrs["w:cryptAlgorithmClass"] = opts.cryptoAlgorithmClass;
  else if (opts.cryptoAlgorithmSid !== undefined)
    attrs["w:cryptAlgorithmSid"] = opts.cryptoAlgorithmSid;
  if (opts.cryptoAlgorithmType !== undefined)
    attrs["w:cryptAlgorithmType"] = opts.cryptoAlgorithmType;
  if (opts.cryptoAlgorithmSid !== undefined && opts.cryptoAlgorithmSidRaw === undefined)
    attrs["w:cryptAlgorithmSid"] = opts.cryptoAlgorithmSid;
  if (opts.cryptoProvider !== undefined) attrs["w:cryptProvider"] = opts.cryptoProvider;
  if (opts.cryptoProviderType !== undefined) attrs["w:cryptProviderType"] = opts.cryptoProviderType;
  if (opts.hashRaw !== undefined) attrs["w:hash"] = opts.hashRaw;
  else if (opts.hash !== undefined) attrs["w:hash"] = opts.hash;
  if (opts.saltRaw !== undefined) attrs["w:salt"] = opts.saltRaw;
  else if (opts.salt !== undefined) attrs["w:salt"] = opts.salt;
  if (opts.cryptoProviderTypeExtension !== undefined)
    attrs["w:cryptProviderTypeExt"] = opts.cryptoProviderTypeExtension;
  if (opts.cryptoProviderTypeExtensionSource !== undefined)
    attrs["w:cryptProviderTypeExtSource"] = opts.cryptoProviderTypeExtensionSource;
  if (opts.algorithmExtensionId !== undefined) attrs["w:algIdExt"] = opts.algorithmExtensionId;
  if (opts.algorithmExtensionSource !== undefined)
    attrs["w:algIdExtSource"] = opts.algorithmExtensionSource;
  if (opts.cryptoSpinCount !== undefined) attrs["w:cryptSpinCount"] = opts.cryptoSpinCount;
  return attrEl("w:documentProtection", attrs);
}

function stringifyWriteProtect(opts: WriteProtectionOptions): string {
  const derived = maybeDerive(opts.password, opts.hashValue);
  const attrs: Record<string, string | number> = {};
  if (opts.recommended !== undefined) attrs["w:recommended"] = opts.recommended ? "1" : "0";
  if (opts.hashValue ?? derived?.hashValue)
    attrs["w:hashValue"] = opts.hashValue ?? derived!.hashValue;
  if (opts.saltValue ?? derived?.saltValue)
    attrs["w:saltValue"] = opts.saltValue ?? derived!.saltValue;
  if (opts.hash !== undefined) attrs["w:hash"] = opts.hash;
  if (opts.salt !== undefined) attrs["w:salt"] = opts.salt;
  if (opts.spinCount ?? derived?.spinCount)
    attrs["w:spinCount"] = opts.spinCount ?? derived!.spinCount;
  if (opts.algorithmName ?? derived?.algorithmName)
    attrs["w:algorithmName"] = opts.algorithmName ?? derived!.algorithmName;
  if (opts.cryptoAlgorithmClass !== undefined)
    attrs["w:cryptAlgorithmClass"] = opts.cryptoAlgorithmClass;
  if (opts.cryptoAlgorithmSid !== undefined) attrs["w:cryptAlgorithmSid"] = opts.cryptoAlgorithmSid;
  if (opts.cryptoAlgorithmType !== undefined)
    attrs["w:cryptAlgorithmType"] = opts.cryptoAlgorithmType;
  if (opts.cryptoProvider !== undefined) attrs["w:cryptProvider"] = opts.cryptoProvider;
  if (opts.cryptoProviderType !== undefined) attrs["w:cryptProviderType"] = opts.cryptoProviderType;
  if (opts.algorithmExtensionId !== undefined) attrs["w:algIdExt"] = opts.algorithmExtensionId;
  if (opts.algorithmExtensionSource !== undefined)
    attrs["w:algIdExtSource"] = opts.algorithmExtensionSource;
  if (opts.cryptoProviderTypeExtension !== undefined)
    attrs["w:cryptProviderTypeExt"] = opts.cryptoProviderTypeExtension;
  if (opts.cryptoProviderTypeExtensionSource !== undefined)
    attrs["w:cryptProviderTypeExtSource"] = opts.cryptoProviderTypeExtensionSource;
  if (opts.cryptoSpinCount !== undefined) attrs["w:cryptSpinCount"] = opts.cryptoSpinCount;
  return attrEl("w:writeProtection", attrs);
}

function stringifyRevisionView(opts: RevisionViewOptions): string {
  const attrs: Record<string, string> = {};
  if (opts.markup !== undefined) attrs["w:markup"] = opts.markup ? "true" : "false";
  if (opts.comments !== undefined) attrs["w:comments"] = opts.comments ? "true" : "false";
  if (opts.insDel !== undefined) attrs["w:insDel"] = opts.insDel ? "true" : "false";
  if (opts.formatting !== undefined) attrs["w:formatting"] = opts.formatting ? "true" : "false";
  if (opts.inkAnnotations !== undefined)
    attrs["w:inkAnnotations"] = opts.inkAnnotations ? "true" : "false";
  return attrEl("w:revisionView", attrs);
}

function stringifyMailMerge(opts: MailMergeOptions): string {
  const p: string[] = [];
  p.push(strVal("w:mainDocumentType", opts.mainDocumentType));
  p.push(onOff("w:linkToQuery", opts.linkToQuery));
  p.push(strVal("w:dataType", opts.dataType));
  p.push(strVal("w:connectString", opts.connectString));
  p.push(strVal("w:query", opts.query));
  if (opts.dataSource !== undefined) p.push(attrEl("w:dataSource", { "r:id": opts.dataSource }));
  if (opts.headerSource !== undefined)
    p.push(attrEl("w:headerSource", { "r:id": opts.headerSource }));
  p.push(onOff("w:doNotSuppressBlankLines", opts.doNotSuppressBlankLines));
  p.push(strVal("w:destination", opts.destination));
  p.push(strVal("w:addressFieldName", opts.addressFieldName));
  p.push(strVal("w:mailSubject", opts.mailSubject));
  p.push(onOff("w:mailAsAttachment", opts.mailAsAttachment));
  p.push(onOff("w:viewMergedData", opts.viewMergedData));
  p.push(numVal("w:activeRecord", opts.activeRecord));
  p.push(numVal("w:checkErrors", opts.checkErrors));
  if (opts.odso !== undefined) p.push(stringifyOdso(opts.odso));
  return `<w:mailMerge>${p.join("")}</w:mailMerge>`;
}

function stringifyOdso(opts: OdsoOptions): string {
  const p: string[] = [];
  p.push(strVal("w:udl", opts.udl));
  p.push(strVal("w:table", opts.table));
  if (opts.src !== undefined) p.push(attrEl("w:src", { "r:id": opts.src }));
  p.push(numVal("w:colDelim", opts.colDelim));
  p.push(strVal("w:type", opts.type));
  p.push(onOff("w:fHdr", opts.fHdr));
  if (opts.fieldMapData !== undefined) {
    for (const fm of opts.fieldMapData) p.push(stringifyOdsoFieldMap(fm));
  }
  if (opts.recipientData !== undefined) {
    for (const rd of opts.recipientData) p.push(attrEl("w:recipientData", { "r:id": rd }));
  }
  return `<w:odso>${p.join("")}</w:odso>`;
}

function stringifyOdsoFieldMap(opts: OdsoFieldMapDataOptions): string {
  const p: string[] = [];
  p.push(strVal("w:type", opts.type));
  p.push(strVal("w:name", opts.name));
  p.push(strVal("w:mappedName", opts.mappedName));
  p.push(numVal("w:column", opts.column));
  p.push(strVal("w:lid", opts.lid));
  p.push(onOff("w:dynamicAddress", opts.dynamicAddress));
  return `<w:fieldMapData>${p.join("")}</w:fieldMapData>`;
}

function stringifyFootnotePr(opts: DocumentFootnotePropertiesOptions): string {
  const p: string[] = [];
  if (opts.pos !== undefined) p.push(attrEl("w:pos", { "w:val": opts.pos }));
  if (opts.numFmt !== undefined || opts.format !== undefined) {
    const a: Record<string, string> = {};
    if (opts.numFmt !== undefined) a["w:val"] = opts.numFmt;
    if (opts.format !== undefined) a["w:format"] = opts.format;
    p.push(attrEl("w:numFmt", a));
  }
  p.push(numVal("w:numStart", opts.numStart));
  p.push(strVal("w:numRestart", opts.numRestart));
  for (const id of opts.footnotes ?? []) p.push(attrEl("w:footnote", { "w:id": id }));
  return `<w:footnotePr>${p.join("")}</w:footnotePr>`;
}

function stringifyEndnotePr(opts: DocumentEndnotePropertiesOptions): string {
  const p: string[] = [];
  if (opts.pos !== undefined) p.push(attrEl("w:pos", { "w:val": opts.pos }));
  if (opts.numFmt !== undefined || opts.format !== undefined) {
    const a: Record<string, string> = {};
    if (opts.numFmt !== undefined) a["w:val"] = opts.numFmt;
    if (opts.format !== undefined) a["w:format"] = opts.format;
    p.push(attrEl("w:numFmt", a));
  }
  p.push(numVal("w:numStart", opts.numStart));
  p.push(strVal("w:numRestart", opts.numRestart));
  for (const id of opts.endnotes ?? []) p.push(attrEl("w:endnote", { "w:id": id }));
  return `<w:endnotePr>${p.join("")}</w:endnotePr>`;
}

function stringifyRsids(opts: RsidsOptions): string {
  const p: string[] = [];
  if (opts.rsidRoot !== undefined) p.push(attrEl("w:rsidRoot", { "w:val": opts.rsidRoot }));
  if (opts.rsids !== undefined) {
    for (const rsid of opts.rsids) {
      // Word writes a bare <w:rsid/> when the val attribute is absent —
      // preserve that shape instead of emitting an empty w:val.
      p.push(rsid ? attrEl("w:rsid", { "w:val": rsid }) : "<w:rsid/>");
    }
  }
  return `<w:rsids>${p.join("")}</w:rsids>`;
}

function stringifyReadModeInkLockDown(opts: ReadModeInkLockDownOptions): string {
  return attrEl("w:readModeInkLockDown", {
    "w:actualPg": opts.actualPg === false ? "0" : "1",
    "w:w": opts.w,
    "w:h": opts.h,
    "w:fontSz": opts.fontSz,
  });
}

function stringifyCaptions(opts: CaptionsOptions): string {
  const p: string[] = [];
  for (const cap of opts.captions) {
    const attrs: Record<string, string | number> = { "w:name": cap.name };
    if (cap.pos !== undefined) attrs["w:pos"] = cap.pos;
    if (cap.chapterNumber !== undefined) attrs["w:chapNum"] = cap.chapterNumber ? "1" : "0";
    if (cap.heading !== undefined) attrs["w:heading"] = cap.heading;
    if (cap.noLabel !== undefined) attrs["w:noLabel"] = cap.noLabel ? "1" : "0";
    if (cap.numFmt !== undefined) attrs["w:numFmt"] = cap.numFmt;
    if (cap.sep !== undefined) attrs["w:sep"] = cap.sep;
    p.push(attrEl("w:caption", attrs));
  }
  if (opts.autoCaptions?.length) {
    const acXml = opts.autoCaptions
      .map((ac) => attrEl("w:autoCaption", { "w:name": ac.name, "w:caption": ac.caption }))
      .join("");
    p.push(`<w:autoCaptions>${acXml}</w:autoCaptions>`);
  }
  return `<w:captions>${p.join("")}</w:captions>`;
}

function stringifyMathPr(opts: MathPropertiesOptions): string {
  const p: string[] = [];
  if (opts.mathFont !== undefined) p.push(attrEl("m:mathFont", { "m:val": opts.mathFont }));
  if (opts.binaryOperatorBreak !== undefined)
    p.push(attrEl("m:brkBin", { "m:val": opts.binaryOperatorBreak }));
  if (opts.binaryOperatorBreakSubtraction !== undefined)
    p.push(attrEl("m:brkBinSub", { "m:val": opts.binaryOperatorBreakSubtraction }));
  p.push(onOff("m:smallFrac", opts.smallFractions));
  p.push(onOff("m:dispDef", opts.displayDefaults));
  p.push(measureVal("m:lMargin", opts.leftMargin));
  p.push(measureVal("m:rMargin", opts.rightMargin));
  if (opts.defaultJustification !== undefined)
    p.push(attrEl("m:defJc", { "m:val": opts.defaultJustification }));
  p.push(measureVal("m:preSp", opts.preSpacing));
  p.push(measureVal("m:postSp", opts.postSpacing));
  p.push(measureVal("m:interSp", opts.interSpacing));
  p.push(measureVal("m:intraSp", opts.intraSpacing));
  p.push(measureVal("m:wrapIndent", opts.wrapIndent));
  // CT_MathPr wraps with wrapIndent | wrapRight (a choice) — wrapIndent wins
  if (opts.wrapIndent === undefined) p.push(onOff("m:wrapRight", opts.wrapRight));
  if (opts.integralLimitLocation !== undefined)
    p.push(attrEl("m:intLim", { "m:val": opts.integralLimitLocation }));
  if (opts.naryLimitLocation !== undefined)
    p.push(attrEl("m:naryLim", { "m:val": opts.naryLimitLocation }));
  return `<m:mathPr>${p.join("")}</m:mathPr>`;
}

function stringifyColorSchemeMapping(
  opts: NonNullable<SettingsOptions["colorSchemeMapping"]>,
): string {
  const attrs: Record<string, string> = {};
  if (opts.bg1 !== undefined) attrs["w:bg1"] = opts.bg1;
  if (opts.t1 !== undefined) attrs["w:t1"] = opts.t1;
  if (opts.bg2 !== undefined) attrs["w:bg2"] = opts.bg2;
  if (opts.t2 !== undefined) attrs["w:t2"] = opts.t2;
  if (opts.accent1 !== undefined) attrs["w:accent1"] = opts.accent1;
  if (opts.accent2 !== undefined) attrs["w:accent2"] = opts.accent2;
  if (opts.accent3 !== undefined) attrs["w:accent3"] = opts.accent3;
  if (opts.accent4 !== undefined) attrs["w:accent4"] = opts.accent4;
  if (opts.accent5 !== undefined) attrs["w:accent5"] = opts.accent5;
  if (opts.accent6 !== undefined) attrs["w:accent6"] = opts.accent6;
  if (opts.hyperlink !== undefined) attrs["w:hyperlink"] = opts.hyperlink;
  if (opts.followedHyperlink !== undefined) attrs["w:followedHyperlink"] = opts.followedHyperlink;
  return attrEl("w:clrSchemeMapping", attrs);
}

// ── Compatibility ──

/** MS Office default compatSettings for a fresh Word 2013+ document. */
const FRESH_COMPATIBILITY: CompatibilityOptions = {
  version: 15,
  overrideTableStyleFontSizeAndJustification: true,
  enableOpenTypeFeatures: true,
  doNotFlipMirrorIndents: true,
};

function stringifyCompatibility(opts: CompatibilityOptions): string {
  const p: string[] = [];
  // Individual compat on/off elements (XSD order — same table drives parse)
  for (const [key, tag] of COMPAT_FLAG_MAP) {
    const value = opts[key as keyof CompatibilityOptions];
    if (value !== undefined) p.push(onOff(tag, value as boolean));
  }
  // Word 2010 extension flag — after the CT_Compat on/off elements
  if (opts.enableOpenTypeKerning) p.push("<w14:enableOpenTypeKerning/>");
  // Sugar defaults are needed only when parsed/fresh input did not retain the
  // original compatSetting entries. Parsed input always carries every entry in
  // compatSettings so duplicate producer entries round-trip verbatim.
  if (opts.compatSettings) {
    for (const cs of opts.compatSettings) p.push(compatSetting(cs.name, cs.val, cs.uri));
  } else {
    if (opts.version !== undefined) p.push(compatSetting("compatibilityMode", opts.version));
    if (opts.overrideTableStyleFontSizeAndJustification !== undefined)
      p.push(
        compatSetting(
          "overrideTableStyleFontSizeAndJustification",
          opts.overrideTableStyleFontSizeAndJustification ? 1 : 0,
        ),
      );
    if (opts.enableOpenTypeFeatures !== undefined)
      p.push(compatSetting("enableOpenTypeFeatures", opts.enableOpenTypeFeatures ? 1 : 0));
    if (opts.doNotFlipMirrorIndents !== undefined)
      p.push(compatSetting("doNotFlipMirrorIndents", opts.doNotFlipMirrorIndents ? 1 : 0));
  }
  // An explicitly empty options object round-trips a bare `<w:compat/>`
  // (source documents carry it as a placeholder with no flags).
  return p.length ? `<w:compat>${p.join("")}</w:compat>` : "<w:compat/>";
}

// ── Namespace attributes ──

const SETTINGS_NAMESPACE_KEYS = [
  "m",
  "mc",
  "o",
  "r",
  "v",
  "w",
  "w10",
  "w14",
  "w15",
  "wne",
  "wp",
  "wp14",
  "wpc",
  "wpg",
  "wpi",
  "wps",
] as const;

// ── Descriptor ──

type SettingsDescriptorContext = Parameters<typeof documentNamespaceDialect>[0] & {
  _options?: Pick<DocumentOptions, "passthroughRelationships">;
};

/**
 * Preserve a source attached-template relationship id; allocate a free settings
 * relationship id when hand-authored options introduce a new template.
 */
export function attachedTemplateRelationshipId(
  attachedTemplate: string,
  relationships: readonly {
    source: string;
    rId: string;
    relationshipType: string;
    target: string;
  }[] = [],
  reservedIds: readonly string[] = [],
): string {
  const source = relationships.find(
    (relationship) =>
      relationship.source === "word/settings.xml" &&
      relationship.relationshipType === RELATIONSHIP_TYPES.attachedTemplate &&
      relationship.target === attachedTemplate,
  );
  if (source) return source.rId;
  const usedIds = new Set([
    ...relationships
      .filter((relationship) => relationship.source === "word/settings.xml")
      .map((relationship) => relationship.rId),
    ...reservedIds,
  ]);
  let id = 1;
  while (usedIds.has(`rId${id}`)) id += 1;
  return `rId${id}`;
}

export const settingsDesc: CustomDescriptor<SettingsOptions, SettingsDescriptorContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const p: string[] = [];

    // XSD CT_Settings sequence order
    if (opts.writeProtection !== undefined) p.push(stringifyWriteProtect(opts.writeProtection));
    if (opts.view !== undefined) p.push(attrEl("w:view", { "w:val": opts.view }));
    if (opts.zoom !== undefined)
      p.push(attrEl("w:zoom", { "w:percent": opts.zoom.percent ?? 100, "w:val": opts.zoom.val }));
    p.push(onOff("w:removePersonalInformation", opts.removePersonalInformation));
    p.push(onOff("w:removeDateAndTime", opts.removeDateAndTime));
    // XSD order: doNotDisplayPageBoundaries (5) → displayBackgroundShape (6)
    p.push(onOff("w:doNotDisplayPageBoundaries", opts.doNotDisplayPageBoundaries));
    p.push(onOff("w:displayBackgroundShape", opts.displayBackgroundShape));
    p.push(onOff("w:printPostScriptOverText", opts.printPostScriptOverText));
    p.push(onOff("w:printFractionalCharacterWidth", opts.printFractionalCharacterWidth));
    p.push(onOff("w:printFormsData", opts.printFormsData));
    p.push(onOff("w:embedTrueTypeFonts", opts.embedTrueTypeFonts));
    p.push(onOff("w:embedSystemFonts", opts.embedSystemFonts));
    p.push(onOff("w:saveSubsetFonts", opts.saveSubsetFonts));
    p.push(onOff("w:saveFormsData", opts.saveFormsData));
    p.push(onOff("w:mirrorMargins", opts.mirrorMargins));
    p.push(onOff("w:alignBordersAndEdges", opts.alignBordersAndEdges));
    p.push(onOff("w:bordersDoNotSurroundHeader", opts.bordersDoNotSurroundHeader));
    p.push(onOff("w:bordersDoNotSurroundFooter", opts.bordersDoNotSurroundFooter));
    p.push(onOff("w:gutterAtTop", opts.gutterAtTop));
    p.push(onOff("w:hideSpellingErrors", opts.hideSpellingErrors));
    p.push(onOff("w:hideGrammaticalErrors", opts.hideGrammaticalErrors));

    // activeWritingStyle
    if (opts.activeWritingStyle !== undefined) {
      for (const ws of opts.activeWritingStyle) {
        const attrs: Record<string, string | number | boolean> = {};
        if (ws.lang !== undefined) attrs["w:lang"] = ws.lang;
        if (ws.vendorID !== undefined) attrs["w:vendorID"] = ws.vendorID;
        if (ws.dllVersion !== undefined) attrs["w:dllVersion"] = ws.dllVersion;
        if (ws.nlCheck !== undefined) attrs["w:nlCheck"] = ws.nlCheck;
        if (ws.checkStyle !== undefined) attrs["w:checkStyle"] = ws.checkStyle;
        if (ws.appName !== undefined) attrs["w:appName"] = ws.appName;
        p.push(attrEl("w:activeWritingStyle", attrs));
      }
    }

    // proofState
    if (opts.proofState !== undefined) {
      const attrs: Record<string, string> = {};
      if (opts.proofState.spelling !== undefined) attrs["w:spelling"] = opts.proofState.spelling;
      if (opts.proofState.grammar !== undefined) attrs["w:grammar"] = opts.proofState.grammar;
      p.push(attrEl("w:proofState", attrs));
    }

    p.push(onOff("w:formsDesign", opts.formsDesign));

    if (opts.attachedTemplate !== undefined)
      p.push(
        attrEl("w:attachedTemplate", {
          "r:id": attachedTemplateRelationshipId(
            opts.attachedTemplate,
            ctx._options?.passthroughRelationships ?? [],
            opts.mailMerge?.odso?.recipientData,
          ),
        }),
      );
    p.push(onOff("w:linkStyles", opts.linkStyles));

    // stylePaneFormatFilter
    if (opts.stylePaneFormatFilter !== undefined) {
      const f = opts.stylePaneFormatFilter;
      const attrs: Record<string, string> = {};
      const flags: [keyof typeof f, string][] = [
        ["allStyles", "w:allStyles"],
        ["customStyles", "w:customStyles"],
        ["stylesInUse", "w:stylesInUse"],
        ["headingStyles", "w:headingStyles"],
        ["numberingStyles", "w:numberingStyles"],
        ["tableStyles", "w:tableStyles"],
        ["directFormattingOnRuns", "w:directFormattingOnRuns"],
        ["directFormattingOnParagraphs", "w:directFormattingOnParagraphs"],
        ["directFormattingOnNumbering", "w:directFormattingOnNumbering"],
        ["directFormattingOnTables", "w:directFormattingOnTables"],
        ["clearFormatting", "w:clearFormatting"],
        ["top3HeadingStyles", "w:top3HeadingStyles"],
        ["visibleStyles", "w:visibleStyles"],
        ["alternateStyleNames", "w:alternateStyleNames"],
        ["latentStyles", "w:latentStyles"],
      ];
      for (const [prop, xmlKey] of flags) {
        if (f[prop] !== undefined) attrs[xmlKey] = f[prop] ? "1" : "0";
      }
      // Legacy hex bitmask (Word 2007 form) — string keeps leading zeros
      if (f.val !== undefined) attrs["w:val"] = f.val;
      p.push(attrEl("w:stylePaneFormatFilter", attrs));
    }

    p.push(strVal("w:stylePaneSortMethod", opts.stylePaneSortMethod));
    p.push(strVal("w:documentType", opts.documentType));

    // CT_Settings order: mailMerge precedes the revision-tracking group
    if (opts.mailMerge !== undefined) p.push(stringifyMailMerge(opts.mailMerge));

    if (opts.revisionView !== undefined) p.push(stringifyRevisionView(opts.revisionView));
    p.push(onOff("w:trackRevisions", opts.trackRevisions));
    p.push(onOff("w:doNotTrackMoves", opts.doNotTrackMoves));
    p.push(onOff("w:doNotTrackFormatting", opts.doNotTrackFormatting));

    if (opts.documentProtection !== undefined) p.push(stringifyDocProtect(opts.documentProtection));

    p.push(onOff("w:autoFormatOverride", opts.autoFormatOverride));
    p.push(onOff("w:styleLockTheme", opts.styleLockTheme));
    p.push(onOff("w:styleLockQFSet", opts.styleLockQFSet));

    // defaultTabStop — optional (CT_Settings minOccurs=0); emit only when set
    p.push(measureVal("w:defaultTabStop", opts.defaultTabStop));

    // hyphenation — sibling CT_Settings elements, flattened onto the root
    p.push(onOff("w:autoHyphenation", opts.autoHyphenation));
    p.push(numVal("w:consecutiveHyphenLimit", opts.consecutiveHyphenLimit));
    p.push(measureVal("w:hyphenationZone", opts.hyphenationZone));
    p.push(onOff("w:doNotHyphenateCaps", opts.doNotHyphenateCaps));

    p.push(onOff("w:showEnvelope", opts.showEnvelope));
    p.push(numVal("w:summaryLength", opts.summaryLength));
    p.push(strVal("w:clickAndTypeStyle", opts.clickAndTypeStyle));
    p.push(strVal("w:defaultTableStyle", opts.defaultTableStyle));
    p.push(onOff("w:evenAndOddHeaders", opts.evenAndOddHeaders));
    p.push(onOff("w:bookFoldRevPrinting", opts.bookFoldRevPrinting));
    p.push(onOff("w:bookFoldPrinting", opts.bookFoldPrinting));
    p.push(numVal("w:bookFoldPrintingSheets", opts.bookFoldPrintingSheets));
    p.push(measureVal("w:drawingGridHorizontalSpacing", opts.drawingGridHorizontalSpacing));
    p.push(measureVal("w:drawingGridVerticalSpacing", opts.drawingGridVerticalSpacing));
    p.push(numVal("w:displayHorizontalDrawingGridEvery", opts.displayHorizontalDrawingGridEvery));
    p.push(numVal("w:displayVerticalDrawingGridEvery", opts.displayVerticalDrawingGridEvery));
    // XSD order: doNotUseMarginsForDrawingGridOrigin (55) → origins (56,57)
    p.push(
      onOff("w:doNotUseMarginsForDrawingGridOrigin", opts.doNotUseMarginsForDrawingGridOrigin),
    );
    p.push(numVal("w:drawingGridHorizontalOrigin", opts.drawingGridHorizontalOrigin));
    p.push(numVal("w:drawingGridVerticalOrigin", opts.drawingGridVerticalOrigin));
    p.push(onOff("w:doNotShadeFormData", opts.doNotShadeFormData));
    p.push(onOff("w:noPunctuationKerning", opts.noPunctuationKerning));

    // characterSpacingControl — optional (CT_Settings minOccurs=0); emit only when set
    if (opts.characterSpacingControl !== undefined) {
      p.push(strVal("w:characterSpacingControl", opts.characterSpacingControl)!);
    }

    p.push(onOff("w:printTwoOnOne", opts.printTwoOnOne));
    p.push(onOff("w:strictFirstAndLastChars", opts.strictFirstAndLastChars));

    if (opts.noLineBreaksAfter !== undefined) {
      const attrs: Record<string, string> = {};
      if (opts.noLineBreaksAfter.lang !== undefined) attrs["w:lang"] = opts.noLineBreaksAfter.lang;
      if (opts.noLineBreaksAfter.val !== undefined) attrs["w:val"] = opts.noLineBreaksAfter.val;
      p.push(attrEl("w:noLineBreaksAfter", attrs));
    }
    if (opts.noLineBreaksBefore !== undefined) {
      const attrs: Record<string, string> = {};
      if (opts.noLineBreaksBefore.lang !== undefined)
        attrs["w:lang"] = opts.noLineBreaksBefore.lang;
      if (opts.noLineBreaksBefore.val !== undefined) attrs["w:val"] = opts.noLineBreaksBefore.val;
      p.push(attrEl("w:noLineBreaksBefore", attrs));
    }

    p.push(onOff("w:savePreviewPicture", opts.savePreviewPicture));
    p.push(onOff("w:doNotValidateAgainstSchema", opts.doNotValidateAgainstSchema));
    p.push(onOff("w:saveInvalidXml", opts.saveInvalidXml));
    p.push(onOff("w:ignoreMixedContent", opts.ignoreMixedContent));
    p.push(onOff("w:alwaysShowPlaceholderText", opts.alwaysShowPlaceholderText));
    p.push(onOff("w:doNotDemarcateInvalidXml", opts.doNotDemarcateInvalidXml));
    p.push(onOff("w:saveXmlDataOnly", opts.saveXmlDataOnly));
    p.push(onOff("w:useXSLTWhenSaving", opts.useXSLTWhenSaving));

    if (opts.saveThroughXslt !== undefined) {
      const attrs: Record<string, string> = {};
      if (opts.saveThroughXslt.id !== undefined) attrs["r:id"] = opts.saveThroughXslt.id;
      if (opts.saveThroughXslt.solutionID !== undefined)
        attrs["w:solutionID"] = opts.saveThroughXslt.solutionID;
      p.push(attrEl("w:saveThroughXslt", attrs));
    }

    p.push(onOff("w:showXMLTags", opts.showXMLTags));
    p.push(onOff("w:alwaysMergeEmptyNamespace", opts.alwaysMergeEmptyNamespace));
    p.push(onOff("w:updateFields", opts.updateFields));

    if (opts.hdrShapeDefaults !== undefined)
      p.push(
        `<w:hdrShapeDefaults>${stringifyShapeDefaultsInner(opts.hdrShapeDefaults)}</w:hdrShapeDefaults>`,
      );
    if (opts.footnoteProperties !== undefined) p.push(stringifyFootnotePr(opts.footnoteProperties));
    if (opts.endnoteProperties !== undefined) p.push(stringifyEndnotePr(opts.endnoteProperties));

    // Compatibility — tri-state:
    //   undefined → fresh MS Office defaults (4 compatSettings)
    //   object    → exactly the fields given (no version injected — a source
    //               compat with only on/off flags round-trips as-is)
    //   false     → omit <w:compat> entirely
    if (opts.compatibility !== false) {
      const compatOpts =
        opts.compatibility === undefined ? FRESH_COMPATIBILITY : opts.compatibility;
      const compatXml = stringifyCompatibility(compatOpts);
      if (compatXml) p.push(compatXml);
    }

    // docVars
    if (opts.docVars?.length) {
      const vars = opts.docVars
        .map((v) => attrEl("w:docVar", { "w:name": v.name, "w:val": v.val }))
        .join("");
      p.push(`<w:docVars>${vars}</w:docVars>`);
    }

    if (opts.rsids !== undefined) p.push(stringifyRsids(opts.rsids));
    if (opts.mathProperties !== undefined) p.push(stringifyMathPr(opts.mathProperties));

    // Not in the ISO transitional schema, but Word writes this flag and the
    // Open XML SDK schema data declares it — keep for round-trip fidelity.
    // Sits between m:mathPr and w:themeFontLang.
    if (opts.uiCompat97To2003 !== undefined)
      p.push(onOff("w:uiCompat97To2003", opts.uiCompat97To2003));

    if (opts.attachedSchema !== undefined) {
      for (const schema of opts.attachedSchema) p.push(strVal("w:attachedSchema", schema)!);
    }

    // XSD order: attachedSchema → themeFontLang → clrSchemeMapping → doNotIncludeSubdocsInStats
    if (opts.themeFontLang !== undefined) {
      const a: Record<string, string> = {};
      if (opts.themeFontLang.val !== undefined) a["w:val"] = opts.themeFontLang.val;
      if (opts.themeFontLang.eastAsia !== undefined) a["w:eastAsia"] = opts.themeFontLang.eastAsia;
      if (opts.themeFontLang.bidi !== undefined) a["w:bidi"] = opts.themeFontLang.bidi;
      p.push(attrEl("w:themeFontLang", a));
    }
    if (opts.colorSchemeMapping !== undefined)
      p.push(stringifyColorSchemeMapping(opts.colorSchemeMapping));
    p.push(onOff("w:doNotIncludeSubdocsInStats", opts.doNotIncludeSubdocsInStats));
    p.push(onOff("w:doNotAutoCompressPictures", opts.doNotAutoCompressPictures));
    if (opts.forceUpgrade !== undefined) p.push("<w:forceUpgrade/>");
    if (opts.captions !== undefined) p.push(stringifyCaptions(opts.captions));
    if (opts.readModeInkLockDown !== undefined)
      p.push(stringifyReadModeInkLockDown(opts.readModeInkLockDown));

    if (opts.smartTagType !== undefined) {
      for (const st of opts.smartTagType) {
        const attrs: Record<string, string> = {};
        if (st.namespace !== undefined) attrs["w:namespace"] = st.namespace;
        if (st.namespaceuri !== undefined) attrs["w:namespaceuri"] = st.namespaceuri;
        if (st.name !== undefined) attrs["w:name"] = st.name;
        if (st.url !== undefined) attrs["w:url"] = st.url;
        p.push(attrEl("w:smartTagType", attrs));
      }
    }

    // XSD order: smartTagType → shapeDefaults → doNotEmbedSmartTags → decimalSymbol
    if (opts.shapeDefaults !== undefined)
      p.push(
        `<w:shapeDefaults>${stringifyShapeDefaultsInner(opts.shapeDefaults)}</w:shapeDefaults>`,
      );
    p.push(onOff("w:doNotEmbedSmartTags", opts.doNotEmbedSmartTags));
    p.push(strVal("w:decimalSymbol", opts.decimalSymbol));
    p.push(strVal("w:listSeparator", opts.listSeparator));

    // Word 2010/2013 document identifiers — CT_Settings trailing sequence
    // (decimalSymbol → listSeparator → w14:docId → w15:chartTrackingRefBased → w15:docId).
    // Preserved verbatim for round-trip fidelity; val attribute is namespace-scoped.
    p.push(docIdVal("w14:docId", opts.w14DocId, opts.w14DocIdText));
    p.push(onOff("w14:discardImageEditingData", opts.w14DiscardImageEditingData));
    if (opts.w14DefaultImageDpi !== undefined)
      p.push(strVal("w14:defaultImageDpi", String(opts.w14DefaultImageDpi)));
    p.push(onOff("w14:conflictMode", opts.w14ConflictMode));
    if (opts.w15ChartTrackingRefBased) p.push(`<w15:chartTrackingRefBased/>`);
    p.push(docIdVal("w15:docId", opts.w15DocId, opts.w15DocIdText));
    p.push(numVal("w:tmPrefOne", opts.tmPrefOne));
    p.push(numVal("w:tmPrefTwo", opts.tmPrefTwo));
    p.push(numVal("w:tmFmtPref", opts.tmFormatPreference));
    if (opts.tmCommentsProperties !== undefined)
      p.push(stringifyTmCommentsProperties(opts.tmCommentsProperties));
    if (opts.tmReviewProperties !== undefined)
      p.push(stringifyTmReviewProperties(opts.tmReviewProperties));
    if (opts.tmLastPosition !== undefined) p.push(stringifyTmLastPosition(opts.tmLastPosition));
    if (opts.tmApplicationRevision !== undefined)
      p.push(
        attrEl("w:tmAppRevision", {
          "w:date": opts.tmApplicationRevision.date,
          "w:val": opts.tmApplicationRevision.val,
        }),
      );

    const body = p.join("");
    return `<w:settings ${documentNamespaceAttributesInDialect(
      SETTINGS_NAMESPACE_KEYS,
      documentNamespaceDialect(ctx),
    )}>${body}</w:settings>`;
  },

  parse(el, _ctx) {
    const opts: SettingsOptions = {};

    // writeProtection (CT_WriteProtection)
    const wpEl = findChild(el, "w:writeProtection");
    if (wpEl) {
      const wp = readPasswordAttrs(wpEl) as WriteProtectionOptions;
      const recommended = attr(wpEl, "w:recommended");
      if (recommended !== undefined) wp.recommended = parseOnOff(recommended) ?? false;
      if (Object.keys(wp).length > 0) opts.writeProtection = wp;
    }

    // view → w:view/@w:val
    const viewVal = readStr(findChild(el, "w:view"), "w:val");
    if (viewVal !== undefined) {
      if (!isSettingsView(viewVal)) {
        throw new DocxParseError(
          `Invalid settings view "${viewVal}".`,
          "word/settings.xml",
          "w:view",
          "invalid-enumeration",
        );
      }
      opts.view = viewVal;
    }

    // zoom → w:zoom/@w:percent, @w:val
    const zoomEl = findChild(el, "w:zoom");
    if (zoomEl) {
      const zoom: Record<string, unknown> = {};
      const percent = attrMeasure(zoomEl, "w:percent");
      if (percent !== undefined) zoom.percent = percent as number | Percentage;
      const zval = attr(zoomEl, "w:val");
      if (zval) zoom.val = zval;
      if (Object.keys(zoom).length > 0) opts.zoom = zoom;
    }

    // CT_OnOff scalar settings (presence = true unless w:val is explicitly false)
    const onOffScalar: Array<{ key: keyof SettingsOptions; tag: string }> = [
      { key: "removePersonalInformation", tag: "w:removePersonalInformation" },
      { key: "removeDateAndTime", tag: "w:removeDateAndTime" },
      { key: "doNotDisplayPageBoundaries", tag: "w:doNotDisplayPageBoundaries" },
      { key: "displayBackgroundShape", tag: "w:displayBackgroundShape" },
      { key: "printPostScriptOverText", tag: "w:printPostScriptOverText" },
      { key: "printFractionalCharacterWidth", tag: "w:printFractionalCharacterWidth" },
      { key: "printFormsData", tag: "w:printFormsData" },
      { key: "embedTrueTypeFonts", tag: "w:embedTrueTypeFonts" },
      { key: "embedSystemFonts", tag: "w:embedSystemFonts" },
      { key: "saveSubsetFonts", tag: "w:saveSubsetFonts" },
      { key: "saveFormsData", tag: "w:saveFormsData" },
      { key: "mirrorMargins", tag: "w:mirrorMargins" },
      { key: "alignBordersAndEdges", tag: "w:alignBordersAndEdges" },
      { key: "bordersDoNotSurroundHeader", tag: "w:bordersDoNotSurroundHeader" },
      { key: "bordersDoNotSurroundFooter", tag: "w:bordersDoNotSurroundFooter" },
      { key: "gutterAtTop", tag: "w:gutterAtTop" },
      { key: "hideSpellingErrors", tag: "w:hideSpellingErrors" },
      { key: "hideGrammaticalErrors", tag: "w:hideGrammaticalErrors" },
      { key: "formsDesign", tag: "w:formsDesign" },
      { key: "linkStyles", tag: "w:linkStyles" },
      { key: "trackRevisions", tag: "w:trackRevisions" },
      { key: "doNotTrackMoves", tag: "w:doNotTrackMoves" },
      { key: "doNotTrackFormatting", tag: "w:doNotTrackFormatting" },
      { key: "autoFormatOverride", tag: "w:autoFormatOverride" },
      { key: "styleLockTheme", tag: "w:styleLockTheme" },
      { key: "styleLockQFSet", tag: "w:styleLockQFSet" },
      { key: "showEnvelope", tag: "w:showEnvelope" },
      { key: "evenAndOddHeaders", tag: "w:evenAndOddHeaders" },
      { key: "bookFoldRevPrinting", tag: "w:bookFoldRevPrinting" },
      { key: "bookFoldPrinting", tag: "w:bookFoldPrinting" },
      { key: "doNotUseMarginsForDrawingGridOrigin", tag: "w:doNotUseMarginsForDrawingGridOrigin" },
      { key: "doNotShadeFormData", tag: "w:doNotShadeFormData" },
      { key: "noPunctuationKerning", tag: "w:noPunctuationKerning" },
      { key: "printTwoOnOne", tag: "w:printTwoOnOne" },
      { key: "strictFirstAndLastChars", tag: "w:strictFirstAndLastChars" },
      { key: "savePreviewPicture", tag: "w:savePreviewPicture" },
      { key: "doNotValidateAgainstSchema", tag: "w:doNotValidateAgainstSchema" },
      { key: "saveInvalidXml", tag: "w:saveInvalidXml" },
      { key: "ignoreMixedContent", tag: "w:ignoreMixedContent" },
      { key: "alwaysShowPlaceholderText", tag: "w:alwaysShowPlaceholderText" },
      { key: "doNotDemarcateInvalidXml", tag: "w:doNotDemarcateInvalidXml" },
      { key: "saveXmlDataOnly", tag: "w:saveXmlDataOnly" },
      { key: "useXSLTWhenSaving", tag: "w:useXSLTWhenSaving" },
      { key: "showXMLTags", tag: "w:showXMLTags" },
      { key: "alwaysMergeEmptyNamespace", tag: "w:alwaysMergeEmptyNamespace" },
      { key: "updateFields", tag: "w:updateFields" },
      { key: "doNotIncludeSubdocsInStats", tag: "w:doNotIncludeSubdocsInStats" },
      { key: "doNotAutoCompressPictures", tag: "w:doNotAutoCompressPictures" },
      { key: "doNotEmbedSmartTags", tag: "w:doNotEmbedSmartTags" },
    ];
    for (const { key, tag } of onOffScalar) {
      const v = readOnOffValue(findChild(el, tag));
      assignSetting(opts, key, v);
    }

    // activeWritingStyle → w:activeWritingStyle[] (multiple)
    const awsList: Array<Record<string, unknown>> = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "w:activeWritingStyle") continue;
      const entry: Record<string, unknown> = {};
      const lang = attr(child, "w:lang");
      if (lang) entry.lang = lang;
      const vendorID = attr(child, "w:vendorID");
      if (vendorID) entry.vendorID = vendorID;
      const dllVersion = attr(child, "w:dllVersion");
      if (dllVersion) entry.dllVersion = dllVersion;
      const nlCheck = attr(child, "w:nlCheck");
      if (nlCheck !== undefined) entry.nlCheck = parseOnOff(nlCheck) ?? true;
      const checkStyle = attr(child, "w:checkStyle");
      if (checkStyle !== undefined) entry.checkStyle = parseOnOff(checkStyle) ?? true;
      const appName = attr(child, "w:appName");
      if (appName) entry.appName = appName;
      if (Object.keys(entry).length > 0) awsList.push(entry);
    }
    if (awsList.length > 0) opts.activeWritingStyle = awsList;

    // proofState → w:proofState/@w:spelling, @w:grammar
    const proofEl = findChild(el, "w:proofState");
    if (proofEl) {
      const proof: Record<string, unknown> = {};
      const spelling = attr(proofEl, "w:spelling");
      if (spelling) proof.spelling = spelling;
      const grammar = attr(proofEl, "w:grammar");
      if (grammar) proof.grammar = grammar;
      if (Object.keys(proof).length > 0) opts.proofState = proof;
    }

    // attachedTemplate → w:attachedTemplate/@r:id — the raw rId is bridged to
    // the target URL in parse.ts (the settings part's own rels live outside
    // this element tree), so here it is kept only as a presence marker.
    const attachedTplEl = findChild(el, "w:attachedTemplate");
    if (attachedTplEl) {
      const rid = attr(attachedTplEl, "r:id");
      if (rid) opts.attachedTemplate = rid;
    }

    // stylePaneFormatFilter → w:stylePaneFormatFilter (complex attributes)
    const spffEl = findChild(el, "w:stylePaneFormatFilter");
    if (spffEl) {
      const filter: Record<string, boolean | string> = {};
      const spffFlags: [string, string][] = [
        ["allStyles", "w:allStyles"],
        ["customStyles", "w:customStyles"],
        ["stylesInUse", "w:stylesInUse"],
        ["headingStyles", "w:headingStyles"],
        ["numberingStyles", "w:numberingStyles"],
        ["tableStyles", "w:tableStyles"],
        ["directFormattingOnRuns", "w:directFormattingOnRuns"],
        ["directFormattingOnParagraphs", "w:directFormattingOnParagraphs"],
        ["directFormattingOnNumbering", "w:directFormattingOnNumbering"],
        ["directFormattingOnTables", "w:directFormattingOnTables"],
        ["clearFormatting", "w:clearFormatting"],
        ["top3HeadingStyles", "w:top3HeadingStyles"],
        ["visibleStyles", "w:visibleStyles"],
        ["alternateStyleNames", "w:alternateStyleNames"],
        ["latentStyles", "w:latentStyles"],
      ];
      for (const [prop, xmlKey] of spffFlags) {
        const v = attr(spffEl, xmlKey);
        if (v !== undefined) filter[prop] = parseOnOff(v) ?? true;
      }
      // Legacy hex bitmask (Word 2007 form, ST_ShortHexNumber)
      const spffVal = attr(spffEl, "w:val");
      if (spffVal !== undefined) filter.val = spffVal;
      if (Object.keys(filter).length > 0) opts.stylePaneFormatFilter = filter;
    }

    // stylePaneSortMethod / documentType / clickAndTypeStyle / defaultTableStyle
    const stylePaneSortMethod = readStr(findChild(el, "w:stylePaneSortMethod"), "w:val");
    if (stylePaneSortMethod)
      opts.stylePaneSortMethod = stylePaneSortMethod as SettingsOptions["stylePaneSortMethod"];
    const documentType = readStr(findChild(el, "w:documentType"), "w:val");
    if (documentType) opts.documentType = documentType as SettingsOptions["documentType"];
    const clickAndTypeStyle = readStr(findChild(el, "w:clickAndTypeStyle"), "w:val");
    if (clickAndTypeStyle) opts.clickAndTypeStyle = clickAndTypeStyle;
    const defaultTableStyle = readStr(findChild(el, "w:defaultTableStyle"), "w:val");
    if (defaultTableStyle) opts.defaultTableStyle = defaultTableStyle;

    // mailMerge → w:mailMerge (CT_MailMerge)
    const mailMergeEl = findChild(el, "w:mailMerge");
    if (mailMergeEl) {
      const mm = parseMailMerge(mailMergeEl);
      if (mm) opts.mailMerge = mm;
    }

    // revisionView → w:revisionView (CT_TrackChangesView)
    const revViewEl = findChild(el, "w:revisionView");
    if (revViewEl) {
      const rv: Record<string, boolean> = {};
      const rvFlags: [string, string][] = [
        ["markup", "w:markup"],
        ["comments", "w:comments"],
        ["insDel", "w:insDel"],
        ["formatting", "w:formatting"],
        ["inkAnnotations", "w:inkAnnotations"],
      ];
      for (const [k, a] of rvFlags) {
        const v = attr(revViewEl, a);
        if (v !== undefined) rv[k] = parseOnOff(v) ?? true;
      }
      if (Object.keys(rv).length > 0) opts.revisionView = rv as RevisionViewOptions;
    }

    // documentProtection → w:documentProtection (CT_DocProtect)
    const docProtEl = findChild(el, "w:documentProtection");
    if (docProtEl) {
      const prot: DocumentProtectionOptions = {};
      const edit = attr(docProtEl, "w:edit");
      if (edit && (DOC_PROTECT_EDITS as readonly string[]).includes(edit)) {
        prot.edit = edit as DocumentProtectionOptions["edit"];
      }
      const enforcement = attr(docProtEl, "w:enforcement");
      Object.assign(prot, readPasswordAttrs(docProtEl));
      const cryptoAlgorithmSidRaw = attr(docProtEl, "w:cryptAlgorithmSid");
      if (cryptoAlgorithmSidRaw !== undefined) prot.cryptoAlgorithmSidRaw = cryptoAlgorithmSidRaw;
      const hashRaw = attr(docProtEl, "w:hash");
      if (hashRaw !== undefined) prot.hashRaw = hashRaw;
      const saltRaw = attr(docProtEl, "w:salt");
      if (saltRaw !== undefined) prot.saltRaw = saltRaw;
      if (enforcement !== undefined) {
        if (parseOnOff(enforcement) !== undefined) {
          prot.enforcementRaw = enforcement as OnOffLexicalValue;
        }
        prot.enforcement = parseOnOff(enforcement) ?? true;
      }
      const formatting = attr(docProtEl, "w:formatting");
      if (formatting !== undefined) prot.formatting = parseOnOff(formatting) ?? false;
      if (enforcement === undefined) prot.enforcement = null;
      if (Object.keys(prot).length > 0) opts.documentProtection = prot;
    }

    // defaultTabStop → w:defaultTabStop/@w:val
    const defaultTabStop = attrMeasure(findChild(el, "w:defaultTabStop"), "w:val");
    if (defaultTabStop !== undefined) opts.defaultTabStop = parsePositiveMeasure(defaultTabStop);

    // hyphenation — sibling CT_Settings elements, flattened onto the root
    const autoHyph = readOnOffValue(findChild(el, "w:autoHyphenation"));
    if (autoHyph !== undefined) opts.autoHyphenation = autoHyph;
    const noHyphCaps = readOnOffValue(findChild(el, "w:doNotHyphenateCaps"));
    if (noHyphCaps !== undefined) opts.doNotHyphenateCaps = noHyphCaps;
    const consLimit = readNum(findChild(el, "w:consecutiveHyphenLimit"), "w:val");
    if (consLimit !== undefined) opts.consecutiveHyphenLimit = consLimit;
    const zone = attrMeasure(findChild(el, "w:hyphenationZone"), "w:val");
    if (zone !== undefined) opts.hyphenationZone = parsePositiveMeasure(zone);

    // summaryLength → w:summaryLength/@w:val
    const summaryLength = readNum(findChild(el, "w:summaryLength"), "w:val");
    if (summaryLength !== undefined) opts.summaryLength = summaryLength;

    // bookFoldPrintingSheets → w:bookFoldPrintingSheets/@w:val
    const bookFoldSheets = readNum(findChild(el, "w:bookFoldPrintingSheets"), "w:val");
    if (bookFoldSheets !== undefined) opts.bookFoldPrintingSheets = bookFoldSheets;

    // drawing grid numerics
    const drawNum: Array<{ key: keyof SettingsOptions; tag: string }> = [
      { key: "drawingGridHorizontalSpacing", tag: "w:drawingGridHorizontalSpacing" },
      { key: "drawingGridVerticalSpacing", tag: "w:drawingGridVerticalSpacing" },
    ];
    for (const { key, tag } of drawNum) {
      const value = attrMeasure(findChild(el, tag), "w:val");
      assignSetting(opts, key, value === undefined ? undefined : parsePositiveMeasure(value));
    }

    const drawInt: Array<{ key: keyof SettingsOptions; tag: string }> = [
      { key: "displayHorizontalDrawingGridEvery", tag: "w:displayHorizontalDrawingGridEvery" },
      { key: "displayVerticalDrawingGridEvery", tag: "w:displayVerticalDrawingGridEvery" },
      { key: "drawingGridHorizontalOrigin", tag: "w:drawingGridHorizontalOrigin" },
      { key: "drawingGridVerticalOrigin", tag: "w:drawingGridVerticalOrigin" },
    ];
    for (const { key, tag } of drawInt) {
      const v = readNum(findChild(el, tag), "w:val");
      assignSetting(opts, key, v);
    }

    // characterSpacingControl → w:characterSpacingControl/@w:val
    const characterSpacingControl = readStr(findChild(el, "w:characterSpacingControl"), "w:val");
    if (characterSpacingControl) {
      if (!isCharacterSpacingControl(characterSpacingControl)) {
        throw new DocxParseError(
          `Invalid character spacing control "${characterSpacingControl}".`,
          "word/settings.xml",
          "w:characterSpacingControl",
          "invalid-enumeration",
        );
      }
      opts.characterSpacingControl = characterSpacingControl;
    }

    // noLineBreaksAfter/Before → w:noLineBreaksAfter/Before (w:lang, w:val)
    const lineBreakPairs: Array<{ key: keyof SettingsOptions; tag: string }> = [
      { key: "noLineBreaksAfter", tag: "w:noLineBreaksAfter" },
      { key: "noLineBreaksBefore", tag: "w:noLineBreaksBefore" },
    ];
    for (const { key, tag } of lineBreakPairs) {
      const lbEl = findChild(el, tag);
      if (!lbEl) continue;
      const entry: Record<string, string> = {};
      const lang = attr(lbEl, "w:lang");
      if (lang !== undefined) entry.lang = lang;
      const val = attr(lbEl, "w:val");
      if (val !== undefined) entry.val = val;
      if (Object.keys(entry).length > 0) assignSetting(opts, key, entry);
    }

    // saveThroughXslt → w:saveThroughXslt (r:id, w:solutionID)
    const stxEl = findChild(el, "w:saveThroughXslt");
    if (stxEl) {
      const stx: Record<string, string> = {};
      const id = attr(stxEl, "r:id");
      if (id) stx.id = id;
      const solutionID = attr(stxEl, "w:solutionID");
      if (solutionID) stx.solutionID = solutionID;
      if (Object.keys(stx).length > 0) opts.saveThroughXslt = stx;
    }

    // hdrShapeDefaults / shapeDefaults → structured CT_ShapeDefaults content
    const hdrSdEl = findChild(el, "w:hdrShapeDefaults");
    if (hdrSdEl) opts.hdrShapeDefaults = parseShapeDefaultsInner(hdrSdEl);
    const sdEl = findChild(el, "w:shapeDefaults");
    if (sdEl) opts.shapeDefaults = parseShapeDefaultsInner(sdEl);

    // footnotePr / endnotePr (CT_FtnDocProps / CT_EdnDocProps)
    const fnPrEl = findChild(el, "w:footnotePr");
    if (fnPrEl) {
      const fn = parseFtnEdnPr(fnPrEl, "w:footnote");
      if (fn) opts.footnoteProperties = fn;
    }
    const enPrEl = findChild(el, "w:endnotePr");
    if (enPrEl) {
      const en = parseFtnEdnPr(enPrEl, "w:endnote");
      if (en) opts.endnoteProperties = en;
    }

    // compatibility — tri-state: object when any <w:compat> holds content,
    // false when absent (prevents fresh defaults from being injected on
    // round-trip). Producer bugs can emit more than one container; parse them
    // all and rebuild the XSD's single-container form.
    const compatEls = (el.elements ?? []).filter(
      (child): child is Element => child.type === "element" && child.name === "w:compat",
    );
    let compat: CompatibilityOptions | undefined;
    for (const compatEl of compatEls) {
      const parsed = parseCompatibility(compatEl);
      if (!parsed) continue;
      compat ??= {};
      for (const [key, value] of Object.entries(parsed)) {
        if (key === "compatSettings") {
          compat.compatSettings = [
            ...(compat.compatSettings ?? []),
            ...(value as CompatSettingOptions[]),
          ];
        } else {
          (compat as Record<string, unknown>)[key] = value;
        }
      }
    }
    // Word 2010 transitional documents can strand compatSetting entries at
    // the settings root, outside w:compat — fold them into the same domain so
    // they re-emit at their XSD position inside w:compat.
    for (const child of el.elements ?? []) {
      if (child.type === "element" && child.name === "w:compatSetting") {
        compat ??= {};
        applyCompatSetting(compat, child);
      }
    }
    opts.compatibility = compat ?? false;

    // docVars → w:docVars/w:docVar
    const docVarsEl = findChild(el, "w:docVars");
    if (docVarsEl) {
      const vars: Array<{ name: string; val: string }> = [];
      for (const child of docVarsEl.elements ?? []) {
        if (child.name !== "w:docVar") continue;
        const name = attr(child, "w:name");
        const val = attr(child, "w:val");
        if (name !== undefined && val !== undefined) vars.push({ name, val });
      }
      if (vars.length > 0) opts.docVars = vars;
    }

    // rsids → w:rsids (CT_DocRsids)
    const rsidsEl = findChild(el, "w:rsids");
    if (rsidsEl) {
      const rsids: RsidsOptions = {};
      const root = readStr(findChild(rsidsEl, "w:rsidRoot"), "w:val");
      if (root) rsids.rsidRoot = root;
      const list: NonNullable<RsidsOptions["rsids"]> = [];
      for (const child of rsidsEl.elements ?? []) {
        if (child.name !== "w:rsid") continue;
        const value = attr(child, "w:val");
        list.push(value ?? "");
      }
      if (list.length > 0) rsids.rsids = list;
      // An empty w:rsids element still round-trips as { rsids: {} } — element
      // existence is part of the fidelity contract (pandoc emits it empty).
      opts.rsids = rsids;
    }

    // mathPr → m:mathPr (CT_MathPr)
    const mathPrEl = findChild(el, "m:mathPr");
    if (mathPrEl) {
      const mp = parseMathPr(mathPrEl);
      if (mp) opts.mathProperties = mp;
    }

    // uiCompat97To2003 — Word 2010+ flag between mathPr and themeFontLang
    const uiCompat = findChild(el, "w:uiCompat97To2003");
    if (uiCompat) opts.uiCompat97To2003 = parseOnOff(attr(uiCompat, "w:val")) ?? true;

    // attachedSchema → w:attachedSchema/@w:val (multiple)
    const attachedSchemas: string[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "w:attachedSchema") continue;
      const val = attr(child, "w:val");
      if (val) attachedSchemas.push(val);
    }
    if (attachedSchemas.length > 0) opts.attachedSchema = attachedSchemas;

    // themeFontLang → w:themeFontLang (CT_Language: val/eastAsia/bidi) —
    // empty-string attributes are meaningful (explicitly unset), keep them.
    const tflEl = findChild(el, "w:themeFontLang");
    if (tflEl) {
      const tfl: Record<string, string> = {};
      const val = attr(tflEl, "w:val");
      if (val !== undefined) tfl.val = val;
      const eastAsia = attr(tflEl, "w:eastAsia");
      if (eastAsia !== undefined) tfl.eastAsia = eastAsia;
      const bidi = attr(tflEl, "w:bidi");
      if (bidi !== undefined) tfl.bidi = bidi;
      if (Object.keys(tfl).length > 0) opts.themeFontLang = tfl;
    }

    // clrSchemeMapping → w:clrSchemeMapping
    const csmEl = findChild(el, "w:clrSchemeMapping");
    if (csmEl) {
      const csm: Record<string, string> = {};
      const csmMap: [string, string][] = [
        ["bg1", "w:bg1"],
        ["t1", "w:t1"],
        ["bg2", "w:bg2"],
        ["t2", "w:t2"],
        ["accent1", "w:accent1"],
        ["accent2", "w:accent2"],
        ["accent3", "w:accent3"],
        ["accent4", "w:accent4"],
        ["accent5", "w:accent5"],
        ["accent6", "w:accent6"],
        ["hyperlink", "w:hyperlink"],
        ["followedHyperlink", "w:followedHyperlink"],
      ];
      for (const [key, xmlAttr] of csmMap) {
        const v = attr(csmEl, xmlAttr);
        if (v) csm[key] = v;
      }
      if (Object.keys(csm).length > 0) opts.colorSchemeMapping = csm;
    }

    // forceUpgrade → w:forceUpgrade (CT_Empty, presence)
    if (findChild(el, "w:forceUpgrade")) opts.forceUpgrade = true;

    // captions → w:captions (CT_Captions)
    const captionsEl = findChild(el, "w:captions");
    if (captionsEl) {
      const captions = parseCaptions(captionsEl);
      if (captions) opts.captions = captions;
    }

    // readModeInkLockDown → w:readModeInkLockDown
    const rmilEl = findChild(el, "w:readModeInkLockDown");
    if (rmilEl) {
      opts.readModeInkLockDown = {
        actualPg: parseOnOff(attr(rmilEl, "w:actualPg")) ?? true,
        w: parseInt(attr(rmilEl, "w:w") ?? "0", 10),
        h: parseInt(attr(rmilEl, "w:h") ?? "0", 10),
        fontSz: parseInt(attr(rmilEl, "w:fontSz") ?? "0", 10),
      };
    }

    // smartTagType → w:smartTagType[] (multiple)
    const smartTags: Array<Record<string, string>> = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "w:smartTagType") continue;
      const entry: Record<string, string> = {};
      const ns = attr(child, "w:namespace");
      if (ns) entry.namespace = ns;
      const nsuri = attr(child, "w:namespaceuri");
      if (nsuri) entry.namespaceuri = nsuri;
      const name = attr(child, "w:name");
      if (name) entry.name = name;
      const url = attr(child, "w:url");
      if (url) entry.url = url;
      if (Object.keys(entry).length > 0) smartTags.push(entry);
    }
    if (smartTags.length > 0) opts.smartTagType = smartTags;

    // decimalSymbol / listSeparator → w:decimalSymbol|w:listSeparator/@w:val
    const decimalSymbol = readStr(findChild(el, "w:decimalSymbol"), "w:val");
    if (decimalSymbol) opts.decimalSymbol = decimalSymbol;
    const listSeparator = readStr(findChild(el, "w:listSeparator"), "w:val");
    if (listSeparator) opts.listSeparator = listSeparator;

    // Word 2010/2013 document identifiers — preserve for round-trip fidelity
    const w14DocId = readStr(findChild(el, "w14:docId"), "w14:val");
    if (w14DocId) opts.w14DocId = w14DocId;
    const w14DocIdText = textOf(findChild(el, "w14:docId"));
    if (w14DocIdText !== undefined && w14DocIdText !== "") opts.w14DocIdText = w14DocIdText;
    const w14Discard = findChild(el, "w14:discardImageEditingData");
    if (w14Discard)
      opts.w14DiscardImageEditingData = parseOnOff(attr(w14Discard, "w14:val")) ?? true;
    const w14DefaultImageDpi = readStr(findChild(el, "w14:defaultImageDpi"), "w14:val");
    if (w14DefaultImageDpi !== undefined) {
      const dpi = Number(w14DefaultImageDpi);
      if (!isNaN(dpi)) opts.w14DefaultImageDpi = dpi;
    }
    const w14ConflictMode = findChild(el, "w14:conflictMode");
    if (w14ConflictMode)
      opts.w14ConflictMode = parseOnOff(attr(w14ConflictMode, "w14:val")) ?? true;
    if (findChild(el, "w15:chartTrackingRefBased")) opts.w15ChartTrackingRefBased = true;
    const w15DocId = readStr(findChild(el, "w15:docId"), "w15:val");
    if (w15DocId) opts.w15DocId = w15DocId;
    const w15DocIdText = textOf(findChild(el, "w15:docId"));
    if (w15DocIdText !== undefined && w15DocIdText !== "") opts.w15DocIdText = w15DocIdText;
    const tmPrefOne = readNum(findChild(el, "w:tmPrefOne"), "w:val");
    if (tmPrefOne !== undefined) opts.tmPrefOne = tmPrefOne;
    const tmPrefTwo = readNum(findChild(el, "w:tmPrefTwo"), "w:val");
    if (tmPrefTwo !== undefined) opts.tmPrefTwo = tmPrefTwo;
    const tmFormatPreference = readNum(findChild(el, "w:tmFmtPref"), "w:val");
    if (tmFormatPreference !== undefined) opts.tmFormatPreference = tmFormatPreference;
    const tmCommentsProperties = findChild(el, "w:tmCommentsPr");
    if (tmCommentsProperties)
      opts.tmCommentsProperties = parseTmCommentsProperties(tmCommentsProperties);
    const tmReviewProperties = findChild(el, "w:tmReviewPr");
    if (tmReviewProperties) opts.tmReviewProperties = parseTmReviewProperties(tmReviewProperties);
    const tmLastPosition = findChild(el, "w:tmLastPos");
    if (tmLastPosition) opts.tmLastPosition = parseTmLastPosition(tmLastPosition);
    const tmApplicationRevision = findChild(el, "w:tmAppRevision");
    if (tmApplicationRevision)
      opts.tmApplicationRevision = {
        date: readNum(tmApplicationRevision, "w:date"),
        val: readNum(tmApplicationRevision, "w:val"),
      };

    return opts;
  },
};

/** Serialize the o: element sequence hosted by w:hdrShapeDefaults/w:shapeDefaults. */
function stringifyShapeDefaultsInner(opts: ShapeDefaultsOptions): string {
  const serializers = {
    shapedefaults: stringifyVmlShapeDefaults,
    shapelayout: stringifyVmlShapeLayout,
  } as const;
  const emitted = new Set<string>();
  const parts: string[] = [];
  for (const key of opts.childOrder ?? []) {
    if (opts[key] === undefined || emitted.has(key)) continue;
    parts.push(serializers[key](opts[key]));
    emitted.add(key);
  }
  for (const key of ["shapedefaults", "shapelayout"] as const) {
    if (opts[key] === undefined || emitted.has(key)) continue;
    parts.push(serializers[key](opts[key]));
  }
  return parts.join("");
}

/** Parse the o: element sequence hosted by w:hdrShapeDefaults/w:shapeDefaults. */
function parseShapeDefaultsInner(el: Element): ShapeDefaultsOptions {
  const out: Partial<ShapeDefaultsOptions> = {};
  const childOrder: NonNullable<ShapeDefaultsOptions["childOrder"]> = [];
  for (const child of el.elements ?? []) {
    if (child.name === "o:shapedefaults") {
      out.shapedefaults = parseVmlShapeDefaults(child);
      childOrder.push("shapedefaults");
    } else if (child.name === "o:shapelayout") {
      out.shapelayout = parseVmlShapeLayout(child);
      childOrder.push("shapelayout");
    }
  }
  if (childOrder.length > 1) out.childOrder = childOrder;
  return out as ShapeDefaultsOptions;
}

/** Serialize one recipientData entry in its source vocabulary. */
function stringifyRecipientData(opts: RecipientDataOptions, prefix: "w" | "wne"): string {
  const active = "active" in opts ? opts.active : undefined;
  const activeXml =
    active === true
      ? `<${prefix}:active ${valAttr(`${prefix}:active`)}="1"/>`
      : onOff(`${prefix}:active`, active);
  const children =
    "hash" in opts
      ? [activeXml, strVal(`${prefix}:hash`, opts.hash)]
      : [
          activeXml,
          numVal(`${prefix}:column`, opts.column),
          strVal(`${prefix}:uniqueTag`, opts.uniqueTag),
        ];
  return `<${prefix}:recipientData>${children.join("")}</${prefix}:recipientData>`;
}

/** Mail merge recipients part (w:recipients). */
export const mailMergeRecipientsDesc: CustomDescriptor<MailMergeRecipientsOptions> = {
  kind: "custom",

  stringify(opts) {
    if (opts.namespace === "wne") {
      const children = opts.recipients
        .map((recipient) => stringifyRecipientData(recipient, "wne"))
        .join("");
      return `<wne:recipients xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">${children}</wne:recipients>`;
    }
    const children = opts.recipients
      .map((recipient) => stringifyRecipientData(recipient, "w"))
      .join("");
    return `<w:recipients xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${children}</w:recipients>`;
  },

  parse(el) {
    const prefix = el.name === "wne:recipients" ? "wne" : "w";
    const recipients: RecipientDataOptions[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== `${prefix}:recipientData`) continue;
      const activeEl = findChild(child, `${prefix}:active`);
      const active = readOnOff(activeEl);
      const hash = readStr(findChild(child, `${prefix}:hash`), `${prefix}:val`);
      if (hash !== undefined) {
        const recipient: RecipientDataOptions = {
          ...(active !== undefined ? { active } : {}),
          hash,
        };
        recipients.push(recipient);
        continue;
      }
      const column = Number(
        attr(findChild(child, `${prefix}:column`) ?? {}, `${prefix}:val`) ?? Number.NaN,
      );
      const uniqueTag = readStr(findChild(child, `${prefix}:uniqueTag`), `${prefix}:val`);
      if (Number.isFinite(column) && uniqueTag) {
        const recipient: RecipientDataOptions = {
          ...(active !== undefined ? { active } : {}),
          column,
          uniqueTag,
        };
        recipients.push(recipient);
      }
    }
    return { ...(prefix === "wne" ? { namespace: "wne" } : {}), recipients };
  },
};
