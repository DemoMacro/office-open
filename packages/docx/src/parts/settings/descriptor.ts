/**
 * Settings descriptor — produces word/settings.xml.
 *
 * Stringifies pure JSON SettingsOptions into XML without creating
 * Settings/Compatibility/DocumentProtection/etc. class instances.
 * Zero XmlComponent allocation — pure string concatenation.
 *
 * @module
 */

import { derivePasswordHash } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrBool, findChild } from "@office-open/xml";

import type { FeaturesOptions } from "../core-properties";
import type {
  SettingsOptions,
  DocumentProtectionOptions,
  WriteProtectionOptions,
  MailMergeOptions,
  OdsoOptions,
  OdsoFieldMapDataOptions,
  CompatibilityOptions,
  CaptionsOptions,
  MathPropertiesOptions,
  RsidsOptions,
  ReadModeInkLockDownOptions,
  RevisionViewOptions,
  FootnotePropertiesOptions,
  EndnotePropertiesOptions,
} from "./settings";

// ── XML string helpers ──

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Derive the namespace-prefixed val attribute from the element tag. */
function valAttr(tag: string): string {
  const ns = tag.split(":")[0];
  return `${ns}:val`;
}

function onOff(tag: string, val: boolean | undefined): string {
  return val !== undefined ? `<${tag} ${valAttr(tag)}="${val ? 1 : 0}"/>` : "";
}

function numVal(tag: string, val: number | undefined): string {
  return val !== undefined ? `<${tag} ${valAttr(tag)}="${val}"/>` : "";
}

function strVal(tag: string, val: string | undefined): string {
  return val !== undefined ? `<${tag} ${valAttr(tag)}="${escapeAttr(val)}"/>` : "";
}

/** Build attribute string from key-value pairs, skipping undefined. */
function attrStr(attrs: Record<string, string | number | boolean | undefined>): string {
  const parts: string[] = [];
  for (const [k, v] of Object.entries(attrs)) {
    if (v !== undefined) parts.push(`${k}="${escapeAttr(String(v))}"`);
  }
  return parts.join(" ");
}

/** Self-closing element with attributes only. */
function attrEl(tag: string, attrs: Record<string, string | number | boolean | undefined>): string {
  const a = attrStr(attrs);
  return a ? `<${tag} ${a}/>` : `<${tag}/>`;
}

function compatSetting(name: string, val: string | number): string {
  return `<w:compatSetting w:name="${escapeAttr(name)}" w:uri="http://schemas.microsoft.com/office/word" w:val="${val}"/>`;
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

/** Valid w:documentProtection/@w:edit values (ST_DocProtect). */
const DOC_PROTECT_EDITS = ["none", "readOnly", "comments", "trackedChanges", "forms"] as const;

function stringifyDocProtect(opts: DocumentProtectionOptions): string {
  const derived = maybeDerive(opts.password, opts.hashValue);
  const attrs: Record<string, string | number> = { "w:enforcement": "1" };
  if (opts.edit !== undefined) attrs["w:edit"] = opts.edit;
  if (opts.formatting !== undefined) attrs["w:formatting"] = opts.formatting ? "1" : "0";
  if (opts.algorithmName ?? derived?.algorithmName)
    attrs["w:algorithmName"] = opts.algorithmName ?? derived!.algorithmName;
  if (opts.hashValue ?? derived?.hashValue)
    attrs["w:hashValue"] = opts.hashValue ?? derived!.hashValue;
  if (opts.saltValue ?? derived?.saltValue)
    attrs["w:saltValue"] = opts.saltValue ?? derived!.saltValue;
  if (opts.hash !== undefined) attrs["w:hash"] = opts.hash;
  if (opts.salt !== undefined) attrs["w:salt"] = opts.salt;
  if (opts.spinCount ?? derived?.spinCount)
    attrs["w:spinCount"] = opts.spinCount ?? derived!.spinCount;
  if (opts.cryptoAlgorithmClass !== undefined)
    attrs["w:cryptAlgorithmClass"] = opts.cryptoAlgorithmClass;
  if (opts.cryptoAlgorithmSid !== undefined) attrs["w:cryptAlgorithmSid"] = opts.cryptoAlgorithmSid;
  if (opts.cryptoAlgorithmType !== undefined)
    attrs["w:cryptAlgorithmType"] = opts.cryptoAlgorithmType;
  if (opts.cryptoProvider !== undefined) attrs["w:cryptProvider"] = opts.cryptoProvider;
  if (opts.cryptoProviderType !== undefined) attrs["w:cryptProviderType"] = opts.cryptoProviderType;
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
  p.push(onOff("w:active", opts.active));
  if (opts.recipients !== undefined) p.push(attrEl("w:recipients", { "r:id": opts.recipients }));
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
  p.push(strVal("w:uniqueTag", opts.uniqueTag));
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

function stringifyFootnotePr(opts: FootnotePropertiesOptions): string {
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
  return `<w:footnotePr>${p.join("")}</w:footnotePr>`;
}

function stringifyEndnotePr(opts: EndnotePropertiesOptions): string {
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
  return `<w:endnotePr>${p.join("")}</w:endnotePr>`;
}

function stringifyRsids(opts: RsidsOptions): string {
  const p: string[] = [];
  if (opts.rsidRoot !== undefined) p.push(attrEl("w:rsidRoot", { "w:val": opts.rsidRoot }));
  if (opts.rsids !== undefined) {
    for (const rsid of opts.rsids) p.push(attrEl("w:rsid", { "w:val": rsid }));
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
    if (cap.chapNum !== undefined) attrs["w:chapNum"] = cap.chapNum ? "1" : "0";
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
  if (opts.brkBin !== undefined) p.push(attrEl("m:brkBin", { "m:val": opts.brkBin }));
  if (opts.brkBinSub !== undefined) p.push(attrEl("m:brkBinSub", { "m:val": opts.brkBinSub }));
  p.push(onOff("m:smallFrac", opts.smallFrac));
  p.push(onOff("m:dispDef", opts.dispDef));
  p.push(numVal("m:lMargin", opts.lMargin));
  p.push(numVal("m:rMargin", opts.rMargin));
  if (opts.defJc !== undefined) p.push(attrEl("m:defJc", { "m:val": opts.defJc }));
  p.push(numVal("m:wrapIndent", opts.wrapIndent));
  if (opts.intLim !== undefined) p.push(attrEl("m:intLim", { "m:val": opts.intLim }));
  if (opts.naryLim !== undefined) p.push(attrEl("m:naryLim", { "m:val": opts.naryLim }));
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

function stringifyCompatibility(opts: CompatibilityOptions): string {
  const p: string[] = [];
  // Individual compat on/off elements (XSD order)
  if (opts.useSingleBorderforContiguousCells)
    p.push(onOff("w:useSingleBorderforContiguousCells", true));
  if (opts.wordPerfectJustification) p.push(onOff("w:wpJustification", true));
  if (opts.noTabStopForHangingIndent) p.push(onOff("w:noTabHangInd", true));
  if (opts.noLeading) p.push(onOff("w:noLeading", true));
  if (opts.spaceForUnderline) p.push(onOff("w:spaceForUL", true));
  if (opts.noColumnBalance) p.push(onOff("w:noColumnBalance", true));
  if (opts.balanceSingleByteDoubleByteWidth)
    p.push(onOff("w:balanceSingleByteDoubleByteWidth", true));
  if (opts.noExtraLineSpacing) p.push(onOff("w:noExtraLineSpacing", true));
  if (opts.doNotLeaveBackslashAlone) p.push(onOff("w:doNotLeaveBackslashAlone", true));
  if (opts.underlineTrailingSpaces) p.push(onOff("w:ulTrailSpace", true));
  if (opts.doNotExpandShiftReturn) p.push(onOff("w:doNotExpandShiftReturn", true));
  if (opts.spacingInWholePoints) p.push(onOff("w:spacingInWholePoints", true));
  if (opts.lineWrapLikeWord6) p.push(onOff("w:lineWrapLikeWord6", true));
  if (opts.printBodyTextBeforeHeader) p.push(onOff("w:printBodyTextBeforeHeader", true));
  if (opts.printColorsBlack) p.push(onOff("w:printColBlack", true));
  if (opts.spaceWidth) p.push(onOff("w:wpSpaceWidth", true));
  if (opts.showBreaksInFrames) p.push(onOff("w:showBreaksInFrames", true));
  if (opts.subFontBySize) p.push(onOff("w:subFontBySize", true));
  if (opts.suppressBottomSpacing) p.push(onOff("w:suppressBottomSpacing", true));
  if (opts.suppressTopSpacing) p.push(onOff("w:suppressTopSpacing", true));
  if (opts.suppressSpacingAtTopOfPage) p.push(onOff("w:suppressSpacingAtTopOfPage", true));
  if (opts.suppressTopSpacingWP) p.push(onOff("w:suppressTopSpacingWP", true));
  if (opts.suppressSpBfAfterPgBrk) p.push(onOff("w:suppressSpBfAfterPgBrk", true));
  if (opts.swapBordersFacingPages) p.push(onOff("w:swapBordersFacingPages", true));
  if (opts.convertMailMergeEsc) p.push(onOff("w:convMailMergeEsc", true));
  if (opts.truncateFontHeightsLikeWP6) p.push(onOff("w:truncateFontHeightsLikeWP6", true));
  if (opts.macWordSmallCaps) p.push(onOff("w:mwSmallCaps", true));
  if (opts.usePrinterMetrics) p.push(onOff("w:usePrinterMetrics", true));
  if (opts.doNotSuppressParagraphBorders) p.push(onOff("w:doNotSuppressParagraphBorders", true));
  if (opts.wrapTrailSpaces) p.push(onOff("w:wrapTrailSpaces", true));
  if (opts.footnoteLayoutLikeWW8) p.push(onOff("w:footnoteLayoutLikeWW8", true));
  if (opts.shapeLayoutLikeWW8) p.push(onOff("w:shapeLayoutLikeWW8", true));
  if (opts.alignTablesRowByRow) p.push(onOff("w:alignTablesRowByRow", true));
  if (opts.forgetLastTabAlignment) p.push(onOff("w:forgetLastTabAlignment", true));
  if (opts.adjustLineHeightInTable) p.push(onOff("w:adjustLineHeightInTable", true));
  if (opts.autoSpaceLikeWord95) p.push(onOff("w:autoSpaceLikeWord95", true));
  if (opts.noSpaceRaiseLower) p.push(onOff("w:noSpaceRaiseLower", true));
  if (opts.doNotUseHTMLParagraphAutoSpacing)
    p.push(onOff("w:doNotUseHTMLParagraphAutoSpacing", true));
  if (opts.layoutRawTableWidth) p.push(onOff("w:layoutRawTableWidth", true));
  if (opts.layoutTableRowsApart) p.push(onOff("w:layoutTableRowsApart", true));
  if (opts.useWord97LineBreakRules) p.push(onOff("w:useWord97LineBreakRules", true));
  if (opts.doNotBreakWrappedTables) p.push(onOff("w:doNotBreakWrappedTables", true));
  if (opts.doNotSnapToGridInCell) p.push(onOff("w:doNotSnapToGridInCell", true));
  if (opts.selectFieldWithFirstOrLastCharacter)
    p.push(onOff("w:selectFldWithFirstOrLastChar", true));
  if (opts.applyBreakingRules) p.push(onOff("w:applyBreakingRules", true));
  if (opts.doNotWrapTextWithPunctuation) p.push(onOff("w:doNotWrapTextWithPunct", true));
  if (opts.doNotUseEastAsianBreakRules) p.push(onOff("w:doNotUseEastAsianBreakRules", true));
  if (opts.useWord2002TableStyleRules) p.push(onOff("w:useWord2002TableStyleRules", true));
  if (opts.growAutofit) p.push(onOff("w:growAutofit", true));
  if (opts.useFELayout) p.push(onOff("w:useFELayout", true));
  if (opts.useNormalStyleForList) p.push(onOff("w:useNormalStyleForList", true));
  if (opts.doNotUseIndentAsNumberingTabStop)
    p.push(onOff("w:doNotUseIndentAsNumberingTabStop", true));
  if (opts.useAlternateEastAsianLineBreakRules)
    p.push(onOff("w:useAltKinsokuLineBreakRules", true));
  if (opts.allowSpaceOfSameStyleInTable) p.push(onOff("w:allowSpaceOfSameStyleInTable", true));
  if (opts.doNotSuppressIndentation) p.push(onOff("w:doNotSuppressIndentation", true));
  if (opts.doNotAutofitConstrainedTables) p.push(onOff("w:doNotAutofitConstrainedTables", true));
  if (opts.autofitToFirstFixedWidthCell) p.push(onOff("w:autofitToFirstFixedWidthCell", true));
  if (opts.underlineTabInNumberingList) p.push(onOff("w:underlineTabInNumList", true));
  if (opts.displayHangulFixedWidth) p.push(onOff("w:displayHangulFixedWidth", true));
  if (opts.splitPgBreakAndParaMark) p.push(onOff("w:splitPgBreakAndParaMark", true));
  if (opts.doNotVerticallyAlignCellWithSp) p.push(onOff("w:doNotVertAlignCellWithSp", true));
  if (opts.doNotBreakConstrainedForcedTable)
    p.push(onOff("w:doNotBreakConstrainedForcedTable", true));
  if (opts.ignoreVerticalAlignmentInTextboxes) p.push(onOff("w:doNotVertAlignInTxbx", true));
  if (opts.useAnsiKerningPairs) p.push(onOff("w:useAnsiKerningPairs", true));
  if (opts.cachedColumnBalance) p.push(onOff("w:cachedColBalance", true));
  // compatSetting elements last (XSD order)
  if (opts.version) p.push(compatSetting("compatibilityMode", opts.version));
  if (opts.overrideTableStyleFontSizeAndJustification)
    p.push(compatSetting("overrideTableStyleFontSizeAndJustification", 1));
  if (opts.enableOpenTypeFeatures) p.push(compatSetting("enableOpenTypeFeatures", 1));
  if (opts.doNotFlipMirrorIndents) p.push(compatSetting("doNotFlipMirrorIndents", 1));
  return p.length ? `<w:compat>${p.join("")}</w:compat>` : "";
}

// ── Namespace attributes ──

const SETTINGS_NS =
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'mc:Ignorable="w14 w15 wp14"';

// ── Descriptor ──

export const settingsDesc: CustomDescriptor<SettingsOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [];

    // XSD CT_Settings sequence order
    if (opts.writeProtection !== undefined) p.push(stringifyWriteProtect(opts.writeProtection));
    if (opts.view !== undefined) p.push(attrEl("w:view", { "w:val": opts.view }));
    if (opts.zoom !== undefined)
      p.push(attrEl("w:zoom", { "w:percent": opts.zoom.percent ?? 100, "w:val": opts.zoom.val }));
    p.push(onOff("w:removePersonalInformation", opts.removePersonalInformation));
    p.push(onOff("w:removeDateAndTime", opts.removeDateAndTime));
    p.push(onOff("w:displayBackgroundShape", opts.displayBackgroundShape));
    p.push(onOff("w:doNotDisplayPageBoundaries", opts.doNotDisplayPageBoundaries));
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
        if (ws.appCheck !== undefined) attrs["w:appCheck"] = ws.appCheck;
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
      p.push(attrEl("w:attachedTemplate", { "r:id": opts.attachedTemplate }));
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
      ];
      for (const [prop, xmlKey] of flags) {
        if (f[prop] !== undefined) attrs[xmlKey] = f[prop] ? "1" : "0";
      }
      p.push(attrEl("w:stylePaneFormatFilter", attrs));
    }

    p.push(strVal("w:stylePaneSortMethod", opts.stylePaneSortMethod));
    p.push(strVal("w:documentType", opts.documentType));

    if (opts.revisionView !== undefined) p.push(stringifyRevisionView(opts.revisionView));
    p.push(onOff("w:trackRevisions", opts.trackRevisions));
    p.push(onOff("w:doNotTrackMoves", opts.doNotTrackMoves));
    p.push(onOff("w:doNotTrackFormatting", opts.doNotTrackFormatting));

    if (opts.mailMerge !== undefined) p.push(stringifyMailMerge(opts.mailMerge));
    if (opts.documentProtection !== undefined) p.push(stringifyDocProtect(opts.documentProtection));

    p.push(onOff("w:autoFormatOverride", opts.autoFormatOverride));
    p.push(onOff("w:styleLockTheme", opts.styleLockTheme));
    p.push(onOff("w:styleLockQFSet", opts.styleLockQFSet));

    // defaultTabStop — always present with default
    p.push(numVal("w:defaultTabStop", opts.defaultTabStop ?? 420)!);

    // hyphenation
    p.push(onOff("w:autoHyphenation", opts.hyphenation?.autoHyphenation));
    p.push(numVal("w:consecutiveHyphenLimit", opts.hyphenation?.consecutiveHyphenLimit));
    p.push(numVal("w:hyphenationZone", opts.hyphenation?.hyphenationZone));
    p.push(onOff("w:doNotHyphenateCaps", opts.hyphenation?.doNotHyphenateCaps));

    p.push(onOff("w:showEnvelope", opts.showEnvelope));
    p.push(numVal("w:summaryLength", opts.summaryLength));
    p.push(strVal("w:clickAndTypeStyle", opts.clickAndTypeStyle));
    p.push(strVal("w:defaultTableStyle", opts.defaultTableStyle));
    p.push(onOff("w:evenAndOddHeaders", opts.evenAndOddHeaders));
    p.push(onOff("w:bookFoldRevPrinting", opts.bookFoldRevPrinting));
    p.push(onOff("w:bookFoldPrinting", opts.bookFoldPrinting));
    p.push(numVal("w:bookFoldPrintingSheets", opts.bookFoldPrintingSheets));
    p.push(numVal("w:drawingGridHorizontalSpacing", opts.drawingGridHorizontalSpacing));
    p.push(numVal("w:drawingGridVerticalSpacing", opts.drawingGridVerticalSpacing));
    p.push(numVal("w:displayHorizontalDrawingGridEvery", opts.displayHorizontalDrawingGridEvery));
    p.push(numVal("w:displayVerticalDrawingGridEvery", opts.displayVerticalDrawingGridEvery));
    p.push(numVal("w:drawingGridHorizontalOrigin", opts.drawingGridHorizontalOrigin));
    p.push(numVal("w:drawingGridVerticalOrigin", opts.drawingGridVerticalOrigin));
    p.push(
      onOff("w:doNotUseMarginsForDrawingGridOrigin", opts.doNotUseMarginsForDrawingGridOrigin),
    );
    p.push(onOff("w:doNotShadeFormData", opts.doNotShadeFormData));

    // characterSpacingControl — always present with default
    p.push(
      strVal("w:characterSpacingControl", opts.characterSpacingControl ?? "compressPunctuation")!,
    );

    p.push(onOff("w:noPunctuationKerning", opts.noPunctuationKerning));
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
      if (opts.saveThroughXslt.val !== undefined) attrs["w:val"] = opts.saveThroughXslt.val;
      if (opts.saveThroughXslt.solutionID !== undefined)
        attrs["w:solutionID"] = opts.saveThroughXslt.solutionID;
      p.push(attrEl("w:saveThroughXslt", attrs));
    }

    p.push(onOff("w:showXMLTags", opts.showXMLTags));
    p.push(onOff("w:alwaysMergeEmptyNamespace", opts.alwaysMergeEmptyNamespace));
    p.push(onOff("w:updateFields", opts.updateFields));

    if (opts.hdrShapeDefaults !== undefined) p.push("<w:hdrShapeDefaults/>");
    if (opts.footnotePr !== undefined) p.push(stringifyFootnotePr(opts.footnotePr));
    if (opts.endnotePr !== undefined) p.push(stringifyEndnotePr(opts.endnotePr));

    // Compatibility — always present with defaults
    p.push(
      stringifyCompatibility({
        ...opts.compatibility,
        version: opts.compatibility?.version ?? opts.compatibilityModeVersion ?? 15,
        spaceForUnderline: opts.compatibility?.spaceForUnderline ?? true,
        balanceSingleByteDoubleByteWidth:
          opts.compatibility?.balanceSingleByteDoubleByteWidth ?? true,
        doNotLeaveBackslashAlone: opts.compatibility?.doNotLeaveBackslashAlone ?? true,
        underlineTrailingSpaces: opts.compatibility?.underlineTrailingSpaces ?? true,
        doNotExpandShiftReturn: opts.compatibility?.doNotExpandShiftReturn ?? true,
        adjustLineHeightInTable: opts.compatibility?.adjustLineHeightInTable ?? true,
        useFELayout: opts.compatibility?.useFELayout ?? true,
        overrideTableStyleFontSizeAndJustification:
          opts.compatibility?.overrideTableStyleFontSizeAndJustification ?? true,
        enableOpenTypeFeatures: opts.compatibility?.enableOpenTypeFeatures ?? true,
        doNotFlipMirrorIndents: opts.compatibility?.doNotFlipMirrorIndents ?? true,
      }),
    );

    // docVars
    if (opts.docVars?.length) {
      const vars = opts.docVars
        .map((v) => attrEl("w:docVar", { "w:name": v.name, "w:val": v.val }))
        .join("");
      p.push(`<w:docVars>${vars}</w:docVars>`);
    }

    if (opts.rsids !== undefined) p.push(stringifyRsids(opts.rsids));
    if (opts.mathPr !== undefined) p.push(stringifyMathPr(opts.mathPr));

    if (opts.attachedSchema !== undefined) {
      for (const schema of opts.attachedSchema) p.push(strVal("w:attachedSchema", schema)!);
    }

    if (opts.colorSchemeMapping !== undefined)
      p.push(stringifyColorSchemeMapping(opts.colorSchemeMapping));
    if (opts.themeFontLang !== undefined)
      p.push(attrEl("w:themeFontLang", { "w:val": opts.themeFontLang }));
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

    p.push(onOff("w:doNotEmbedSmartTags", opts.doNotEmbedSmartTags));
    if (opts.shapeDefaults !== undefined) p.push("<w:shapeDefaults/>");
    p.push(strVal("w:decimalSymbol", opts.decimalSymbol));
    p.push(strVal("w:listSeparator", opts.listSeparator));

    const body = p.join("");
    return `<w:settings ${SETTINGS_NS}>${body}</w:settings>`;
  },

  parse(el, _ctx) {
    const opts: Record<string, unknown> = {};

    // evenAndOddHeaderAndFooters → w:evenAndOddHeaders
    const eohEl = findChild(el, "w:evenAndOddHeaders");
    if (eohEl) {
      const val = attr(eohEl, "w:val");
      opts.evenAndOddHeaderAndFooters = val !== "false" && val !== "0" && val !== "off";
    }

    // view → w:view/@w:val
    const viewEl = findChild(el, "w:view");
    if (viewEl) {
      const val = attr(viewEl, "w:val");
      if (val) opts.view = val;
    }

    // zoom → w:zoom/@w:percent, @w:val
    const zoomEl = findChild(el, "w:zoom");
    if (zoomEl) {
      const zoom: Record<string, unknown> = {};
      const percent = attr(zoomEl, "w:percent");
      if (percent) zoom.percent = parseInt(percent, 10);
      const val = attr(zoomEl, "w:val");
      if (val) zoom.val = val;
      if (Object.keys(zoom).length > 0) opts.zoom = zoom;
    }

    // defaultTabStop → w:defaultTabStop/@w:val
    const tabStopEl = findChild(el, "w:defaultTabStop");
    if (tabStopEl) {
      const val = attr(tabStopEl, "w:val");
      if (val) opts.defaultTabStop = parseInt(val, 10);
    }

    // features.trackRevisions → w:trackRevisions (onOff: read w:val)
    const trackRevEl = findChild(el, "w:trackRevisions");
    if (trackRevEl) {
      const features: FeaturesOptions = (opts.features as FeaturesOptions | undefined) ?? {};
      features.trackRevisions = attrBool(trackRevEl, "w:val") ?? true;
      opts.features = features;
    }

    // features.updateFields → w:updateFields (onOff: read w:val)
    const updateFieldsEl = findChild(el, "w:updateFields");
    if (updateFieldsEl) {
      const features: FeaturesOptions = (opts.features as FeaturesOptions | undefined) ?? {};
      features.updateFields = attrBool(updateFieldsEl, "w:val") ?? true;
      opts.features = features;
    }

    // compatibilityModeVersion → w:compat/w:compatSetting
    const compatEl = findChild(el, "w:compat");
    if (compatEl) {
      for (const child of compatEl.elements ?? []) {
        if (child.name === "w:compatSetting") {
          const name = attr(child, "w:name");
          if (name === "compatibilityMode") {
            const val = attr(child, "w:val");
            if (val) opts.compatabilityModeVersion = parseInt(val, 10);
          }
        }
      }
    }

    // docVars → w:docVars/w:docVar
    const docVarsEl = findChild(el, "w:docVars");
    if (docVarsEl) {
      const vars: Array<{ name: string; val: string }> = [];
      for (const child of docVarsEl.elements ?? []) {
        if (child.name === "w:docVar") {
          const name = attr(child, "w:name");
          const val = attr(child, "w:val");
          if (name && val) vars.push({ name, val });
        }
      }
      if (vars.length > 0) opts.docVars = vars;
    }

    // characterSpacingControl → w:characterSpacingControl/@w:val
    const cscEl = findChild(el, "w:characterSpacingControl");
    if (cscEl) {
      const val = attr(cscEl, "w:val");
      if (val) opts.characterSpacingControl = val;
    }

    // displayBackgroundShape → w:displayBackgroundShape (presence)
    if (findChild(el, "w:displayBackgroundShape")) {
      opts.displayBackgroundShape = true;
    }

    // embedTrueTypeFonts → w:embedTrueTypeFonts (presence)
    if (findChild(el, "w:embedTrueTypeFonts")) {
      opts.embedTrueTypeFonts = true;
    }

    // embedSystemFonts → w:embedSystemFonts (presence)
    if (findChild(el, "w:embedSystemFonts")) {
      opts.embedSystemFonts = true;
    }

    // saveSubsetFonts → w:saveSubsetFonts (presence)
    if (findChild(el, "w:saveSubsetFonts")) {
      opts.saveSubsetFonts = true;
    }

    // removePersonalInformation → w:removePersonalInformation (presence)
    if (findChild(el, "w:removePersonalInformation")) {
      opts.removePersonalInformation = true;
    }

    // removeDateAndTime → w:removeDateAndTime (presence)
    if (findChild(el, "w:removeDateAndTime")) {
      opts.removeDateAndTime = true;
    }

    // hideSpellingErrors → w:hideSpellingErrors (presence)
    if (findChild(el, "w:hideSpellingErrors")) {
      opts.hideSpellingErrors = true;
    }

    // hideGrammaticalErrors → w:hideGrammaticalErrors (presence)
    if (findChild(el, "w:hideGrammaticalErrors")) {
      opts.hideGrammaticalErrors = true;
    }

    // mirrorMargins → w:mirrorMargins (presence)
    if (findChild(el, "w:mirrorMargins")) {
      opts.mirrorMargins = true;
    }

    // saveFormsData → w:saveFormsData (presence)
    if (findChild(el, "w:saveFormsData")) {
      opts.saveFormsData = true;
    }

    // alignBordersAndEdges → w:alignBordersAndEdges (presence)
    if (findChild(el, "w:alignBordersAndEdges")) {
      opts.alignBordersAndEdges = true;
    }

    // bordersDoNotSurroundHeader → w:bordersDoNotSurroundHeader (presence)
    if (findChild(el, "w:bordersDoNotSurroundHeader")) {
      opts.bordersDoNotSurroundHeader = true;
    }

    // bordersDoNotSurroundFooter → w:bordersDoNotSurroundFooter (presence)
    if (findChild(el, "w:bordersDoNotSurroundFooter")) {
      opts.bordersDoNotSurroundFooter = true;
    }

    // gutterAtTop → w:gutterAtTop (presence)
    if (findChild(el, "w:gutterAtTop")) {
      opts.gutterAtTop = true;
    }

    // formsDesign → w:formsDesign (presence)
    if (findChild(el, "w:formsDesign")) {
      opts.formsDesign = true;
    }

    // linkStyles → w:linkStyles (presence)
    if (findChild(el, "w:linkStyles")) {
      opts.linkStyles = true;
    }

    // autoHyphenation → w:autoHyphenation
    const autoHyphEl = findChild(el, "w:autoHyphenation");
    if (autoHyphEl) {
      const hyphenation = (opts.hyphenation as Record<string, unknown>) ?? {};
      hyphenation.autoHyphenation = true;
      opts.hyphenation = hyphenation;
    }

    // doNotHyphenateCaps → w:doNotHyphenateCaps
    const noHyphCapsEl = findChild(el, "w:doNotHyphenateCaps");
    if (noHyphCapsEl) {
      const hyphenation = (opts.hyphenation as Record<string, unknown>) ?? {};
      hyphenation.doNotHyphenateCaps = true;
      opts.hyphenation = hyphenation;
    }

    // consecutiveHyphenLimit → w:consecutiveHyphenLimit/@w:val
    const consHyphEl = findChild(el, "w:consecutiveHyphenLimit");
    if (consHyphEl) {
      const val = attr(consHyphEl, "w:val");
      if (val) {
        const hyphenation = (opts.hyphenation as Record<string, unknown>) ?? {};
        hyphenation.consecutiveHyphenLimit = parseInt(val, 10);
        opts.hyphenation = hyphenation;
      }
    }

    // hyphenationZone → w:hyphenationZone/@w:val
    const hyphZoneEl = findChild(el, "w:hyphenationZone");
    if (hyphZoneEl) {
      const val = attr(hyphZoneEl, "w:val");
      if (val) {
        const hyphenation = (opts.hyphenation as Record<string, unknown>) ?? {};
        hyphenation.hyphenationZone = parseInt(val, 10);
        opts.hyphenation = hyphenation;
      }
    }

    // attachedTemplate → w:attachedTemplate/@r:id
    const attachedTplEl = findChild(el, "w:attachedTemplate");
    if (attachedTplEl) {
      const val = attr(attachedTplEl, "r:id");
      if (val) opts.attachedTemplate = val;
    }

    // stylePaneSortMethod → w:stylePaneSortMethod/@w:val
    const spsmEl = findChild(el, "w:stylePaneSortMethod");
    if (spsmEl) {
      const val = attr(spsmEl, "w:val");
      if (val) opts.stylePaneSortMethod = val;
    }

    // documentType → w:documentType/@w:val
    const docTypeEl = findChild(el, "w:documentType");
    if (docTypeEl) {
      const val = attr(docTypeEl, "w:val");
      if (val) opts.documentType = val;
    }

    // defaultTableStyle → w:defaultTableStyle/@w:val
    const defTblStyleEl = findChild(el, "w:defaultTableStyle");
    if (defTblStyleEl) {
      const val = attr(defTblStyleEl, "w:val");
      if (val) opts.defaultTableStyle = val;
    }

    // stylePaneFormatFilter → w:stylePaneFormatFilter (complex attributes)
    const spffEl = findChild(el, "w:stylePaneFormatFilter");
    if (spffEl) {
      const filter: Record<string, boolean> = {};
      const flags: [string, string][] = [
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
      ];
      for (const [prop, xmlKey] of flags) {
        const v = attr(spffEl, xmlKey);
        if (v !== undefined) filter[prop] = v !== "0" && v !== "false" && v !== "off";
      }
      if (Object.keys(filter).length > 0) opts.stylePaneFormatFilter = filter;
    }

    // clickAndTypeStyle → w:clickAndTypeStyle/@w:val
    const catEl = findChild(el, "w:clickAndTypeStyle");
    if (catEl) {
      const val = attr(catEl, "w:val");
      if (val) opts.clickAndTypeStyle = val;
    }

    // activeWritingStyle → iterate w:activeWritingStyle children
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
      if (nlCheck !== undefined) entry.nlCheck = nlCheck !== "0" && nlCheck !== "false";
      const checkStyle = attr(child, "w:checkStyle");
      if (checkStyle !== undefined) entry.checkStyle = checkStyle !== "0" && checkStyle !== "false";
      const appCheck = attr(child, "w:appCheck");
      if (appCheck) entry.appCheck = appCheck;
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

    // documentProtection → w:documentProtection (written under features)
    const docProtEl = findChild(el, "w:documentProtection");
    if (docProtEl) {
      const prot: DocumentProtectionOptions = {};
      const edit = attr(docProtEl, "w:edit");
      if (edit && (DOC_PROTECT_EDITS as readonly string[]).includes(edit)) {
        prot.edit = edit as DocumentProtectionOptions["edit"];
      }
      const enforcement = attr(docProtEl, "w:enforcement");
      if (enforcement !== undefined) {
        // Only include documentProtection if enforcement is set
        const hash = attr(docProtEl, "w:hash");
        if (hash) prot.hash = hash;
        const salt = attr(docProtEl, "w:salt");
        if (salt) prot.salt = salt;
        const cryptProviderType = attr(docProtEl, "w:cryptProviderType");
        if (cryptProviderType) prot.cryptoProviderType = cryptProviderType;
        const cryptAlgorithmClass = attr(docProtEl, "w:cryptAlgorithmClass");
        if (cryptAlgorithmClass) prot.cryptoAlgorithmClass = cryptAlgorithmClass;
        const cryptAlgorithmType = attr(docProtEl, "w:cryptAlgorithmType");
        if (cryptAlgorithmType) prot.cryptoAlgorithmType = cryptAlgorithmType;
        const cryptAlgorithmSid = attr(docProtEl, "w:cryptAlgorithmSid");
        if (cryptAlgorithmSid) prot.cryptoAlgorithmSid = parseInt(cryptAlgorithmSid, 10);
        const cryptSpinCount = attr(docProtEl, "w:cryptSpinCount");
        if (cryptSpinCount) prot.cryptoSpinCount = parseInt(cryptSpinCount, 10);
        const hashValue = attr(docProtEl, "w:hashValue");
        if (hashValue) prot.hashValue = hashValue;
        const saltValue = attr(docProtEl, "w:saltValue");
        if (saltValue) prot.saltValue = saltValue;
        const spinCount = attr(docProtEl, "w:spinCount");
        if (spinCount) prot.spinCount = parseInt(spinCount, 10);
        const algorithmName = attr(docProtEl, "w:algorithmName");
        if (algorithmName) prot.algorithmName = algorithmName;
        const formatting = attr(docProtEl, "w:formatting");
        if (formatting !== undefined) prot.formatting = formatting === "1" || formatting === "true";
      }
      if (Object.keys(prot).length > 0) {
        const features: FeaturesOptions = (opts.features as FeaturesOptions | undefined) ?? {};
        features.documentProtection = prot;
        opts.features = features;
      }
    }

    // doNotTrackMoves → w:doNotTrackMoves (onOff: read w:val)
    const noMovesEl = findChild(el, "w:doNotTrackMoves");
    if (noMovesEl) opts.doNotTrackMoves = attrBool(noMovesEl, "w:val") ?? true;

    // doNotTrackFormatting → w:doNotTrackFormatting (onOff: read w:val)
    const noFmtEl = findChild(el, "w:doNotTrackFormatting");
    if (noFmtEl) opts.doNotTrackFormatting = attrBool(noFmtEl, "w:val") ?? true;

    return opts as unknown as SettingsOptions;
  },
};
