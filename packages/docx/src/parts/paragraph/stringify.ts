/**
 * Direct XML string builders for paragraph and run properties.
 *
 * Replaces `buildParagraphProperties() + xml()` and `buildRunProperties() + xml()`
 * with direct string concatenation — zero intermediate IXmlableObject allocation,
 * zero recursive xml() traversal. Follows PPTX/XLSX pattern.
 *
 * @module
 */

import {
  decimalNumber,
  eighthPointMeasureValue,
  hexColorValue,
  hpsMeasureValue,
  pointMeasureValue,
  signedTwipsMeasureValue,
  twipsMeasureValue,
  uCharHexNumber,
} from "@office-open/core";
import { escapeXml } from "@office-open/xml";
import type { CnfConditionalOptions } from "@parts/paragraph/formatting/cnf-style";
import type { IndentAttributesProperties } from "@parts/paragraph/formatting/indent";
import type { SpacingProperties } from "@parts/paragraph/formatting/spacing";
import type { TabStopDefinition } from "@parts/paragraph/formatting/tab-stop";
import type { FrameOptions } from "@parts/paragraph/frame/frame-properties";
import type { ParagraphPropertiesOptions } from "@parts/paragraph/properties";
import type { EastAsianLayoutOptions } from "@parts/paragraph/run/east-asian-layout";
import type { ColorOptions } from "@parts/paragraph/run/formatting";
import type { LanguageOptions } from "@parts/paragraph/run/language";
import type {
  ParagraphRunPropertiesOptions,
  RunPropertiesChangeOptions,
  RunPropertiesOptions,
} from "@parts/paragraph/run/properties";
import type { FontAttributesProperties } from "@parts/paragraph/run/run-fonts";
import type { BorderOptions } from "@shared/border";
import { BorderStyle } from "@shared/border";
import type { ShadingAttributesProperties } from "@shared/shading";

// ── Inline helpers ──

/** On/off: `<w:name/>` for true, `<w:name w:val="0"/>` for false */
export function onOff(name: string, val: boolean): string {
  return val ? `<${name}/>` : `<${name} w:val="0"/>`;
}

/** Build attrs string from key-value pairs, skipping undefined */
export function attrParts(attrs: Record<string, string | number | boolean | undefined>): string {
  const parts: string[] = [];
  for (const [key, val] of Object.entries(attrs)) {
    if (val !== undefined) parts.push(`${key}="${val}"`);
  }
  return parts.join(" ");
}

// ── Border ──

export function borderStr(name: string, opts: BorderOptions): string {
  const a = attrParts({
    "w:val": opts.style,
    "w:color": opts.color !== undefined ? hexColorValue(opts.color) : undefined,
    "w:sz": opts.size !== undefined ? eighthPointMeasureValue(opts.size) : undefined,
    "w:space": opts.space !== undefined ? pointMeasureValue(opts.space) : undefined,
    "w:themeColor": opts.themeColor,
    "w:themeTint": opts.themeTint !== undefined ? uCharHexNumber(opts.themeTint) : undefined,
    "w:themeShade": opts.themeShade !== undefined ? uCharHexNumber(opts.themeShade) : undefined,
    "w:shadow": opts.shadow !== undefined ? (opts.shadow ? 1 : 0) : undefined,
    "w:frame": opts.frame !== undefined ? (opts.frame ? 1 : 0) : undefined,
  });
  return `<${name} ${a}/>`;
}

// ── Shading ──

export function shadingStr(opts: ShadingAttributesProperties): string {
  const a = attrParts({
    "w:val": opts.type ?? "clear",
    "w:color": opts.color !== undefined ? hexColorValue(opts.color) : undefined,
    "w:fill": opts.fill !== undefined ? hexColorValue(opts.fill) : undefined,
    "w:themeColor": opts.themeColor,
    "w:themeTint": opts.themeTint !== undefined ? uCharHexNumber(opts.themeTint) : undefined,
    "w:themeShade": opts.themeShade !== undefined ? uCharHexNumber(opts.themeShade) : undefined,
    "w:themeFill": opts.themeFill,
    "w:themeFillTint":
      opts.themeFillTint !== undefined ? uCharHexNumber(opts.themeFillTint) : undefined,
    "w:themeFillShade":
      opts.themeFillShade !== undefined ? uCharHexNumber(opts.themeFillShade) : undefined,
  });
  return `<w:shd ${a}/>`;
}

// ── Spacing ──

function spacingStr(opts: SpacingProperties): string {
  const a = attrParts({
    "w:after": opts.after,
    "w:afterAutospacing":
      opts.afterAutoSpacing !== undefined ? (opts.afterAutoSpacing ? 1 : 0) : undefined,
    "w:afterLines": opts.afterLines !== undefined ? decimalNumber(opts.afterLines) : undefined,
    "w:before": opts.before,
    "w:beforeAutospacing":
      opts.beforeAutoSpacing !== undefined ? (opts.beforeAutoSpacing ? 1 : 0) : undefined,
    "w:beforeLines": opts.beforeLines !== undefined ? decimalNumber(opts.beforeLines) : undefined,
    "w:line": opts.line,
    "w:lineRule": opts.lineRule,
  });
  return `<w:spacing ${a}/>`;
}

// ── Indent ──

function indentStr(opts: IndentAttributesProperties): string {
  const a = attrParts({
    "w:start": opts.start !== undefined ? signedTwipsMeasureValue(opts.start) : undefined,
    "w:startChars": opts.startChars !== undefined ? decimalNumber(opts.startChars) : undefined,
    "w:end": opts.end !== undefined ? signedTwipsMeasureValue(opts.end) : undefined,
    "w:endChars": opts.endChars !== undefined ? decimalNumber(opts.endChars) : undefined,
    "w:left": opts.left !== undefined ? signedTwipsMeasureValue(opts.left) : undefined,
    "w:leftChars": opts.leftChars !== undefined ? decimalNumber(opts.leftChars) : undefined,
    "w:right": opts.right !== undefined ? signedTwipsMeasureValue(opts.right) : undefined,
    "w:rightChars": opts.rightChars !== undefined ? decimalNumber(opts.rightChars) : undefined,
    "w:hanging": opts.hanging !== undefined ? twipsMeasureValue(opts.hanging) : undefined,
    "w:hangingChars":
      opts.hangingChars !== undefined ? decimalNumber(opts.hangingChars) : undefined,
    "w:firstLine": opts.firstLine !== undefined ? twipsMeasureValue(opts.firstLine) : undefined,
    "w:firstLineChars":
      opts.firstLineChars !== undefined ? decimalNumber(opts.firstLineChars) : undefined,
  });
  return `<w:ind ${a}/>`;
}

// ── Tab stops ──

function tabStopsStr(defs: TabStopDefinition[]): string {
  const items = defs.map(({ type, position, leader }) => {
    const a = attrParts({ "w:val": type, "w:pos": position, "w:leader": leader });
    return `<w:tab ${a}/>`;
  });
  return `<w:tabs>${items.join("")}</w:tabs>`;
}

// ── CNF style ──

function cnfStyleStr(opts: CnfConditionalOptions): string {
  const a = attrParts({
    "w:firstRow": opts.firstRow ? "1" : "0",
    "w:lastRow": opts.lastRow ? "1" : "0",
    "w:firstColumn": opts.firstColumn ? "1" : "0",
    "w:lastColumn": opts.lastColumn ? "1" : "0",
    "w:oddVBand": opts.oddVBand ? "1" : "0",
    "w:evenVBand": opts.evenVBand ? "1" : "0",
    "w:oddHBand": opts.oddHBand ? "1" : "0",
    "w:evenHBand": opts.evenHBand ? "1" : "0",
    "w:firstRowFirstColumn": opts.firstRowFirstColumn ? "1" : "0",
    "w:firstRowLastColumn": opts.firstRowLastColumn ? "1" : "0",
    "w:lastRowFirstColumn": opts.lastRowFirstColumn ? "1" : "0",
    "w:lastRowLastColumn": opts.lastRowLastColumn ? "1" : "0",
  });
  return `<w:cnfStyle ${a}/>`;
}

// ── Frame properties ──

function framePrStr(opts: FrameOptions): string {
  const alignment = (opts as { alignment?: { x?: string; y?: string } }).alignment;
  const position = (opts as { position?: { x?: number; y?: number } }).position;
  const a = attrParts({
    "w:xAlign": alignment?.x,
    "w:yAlign": alignment?.y,
    "w:hAnchor": opts.anchor?.horizontal,
    "w:anchorLock": opts.anchorLock,
    "w:vAnchor": opts.anchor?.vertical,
    "w:dropCap": opts.dropCap,
    "w:h": opts.height,
    "w:lines": opts.lines,
    "w:hRule": opts.rule,
    "w:hSpace": opts.space?.horizontal,
    "w:vSpace": opts.space?.vertical,
    "w:w": opts.width,
    "w:wrap": opts.wrap,
    "w:x": position?.x,
    "w:y": position?.y,
  });
  return `<w:framePr ${a}/>`;
}

// ── Number properties ──

function numPrStr(
  numberId: number | string,
  indentLevel: number,
  numberingChange?: { original: string; id: string; author: string; date?: string },
): string {
  const idVal = typeof numberId === "string" ? `{${numberId}}` : numberId;
  const parts = [`<w:ilvl w:val="${Math.min(indentLevel, 9)}"/>`, `<w:numId w:val="${idVal}"/>`];
  if (numberingChange) {
    const a = attrParts({
      "w:original": numberingChange.original,
      "w:id": numberingChange.id,
      "w:author": numberingChange.author,
      "w:date": numberingChange.date,
    });
    parts.push(`<w:numberingChange ${a}/>`);
  }
  return `<w:numPr>${parts.join("")}</w:numPr>`;
}

// ── Run-level formatting helpers ──

function colorStr(colorOrOptions: string | ColorOptions): string {
  if (typeof colorOrOptions === "string") {
    return `<w:color w:val="${hexColorValue(colorOrOptions)}"/>`;
  }
  const opts = colorOrOptions;
  const a = attrParts({
    "w:val": opts.val !== undefined ? hexColorValue(opts.val) : undefined,
    "w:themeColor": opts.themeColor,
    "w:themeTint": opts.themeTint !== undefined ? uCharHexNumber(opts.themeTint) : undefined,
    "w:themeShade": opts.themeShade !== undefined ? uCharHexNumber(opts.themeShade) : undefined,
  });
  return `<w:color ${a}/>`;
}

function runFontsStr(nameOrAttrs: string | FontAttributesProperties, hint?: string): string {
  if (typeof nameOrAttrs === "string") {
    const a = attrParts({
      "w:ascii": nameOrAttrs,
      "w:cs": nameOrAttrs,
      "w:eastAsia": nameOrAttrs,
      "w:hAnsi": nameOrAttrs,
      "w:hint": hint,
    });
    return `<w:rFonts ${a}/>`;
  }
  const attrs = nameOrAttrs;
  const a = attrParts({
    "w:ascii": attrs.ascii,
    "w:asciiTheme": attrs.asciiTheme,
    "w:cs": attrs.cs,
    "w:cstheme": attrs.cstheme,
    "w:eastAsia": attrs.eastAsia,
    "w:eastAsiaTheme": attrs.eastAsiaTheme,
    "w:hAnsi": attrs.hAnsi,
    "w:hAnsiTheme": attrs.hAnsiTheme,
    "w:hint": attrs.hint,
  });
  return `<w:rFonts ${a}/>`;
}

function underlineStr(type: string | undefined, color?: string): string {
  const a = attrParts({
    "w:val": type ?? "single",
    "w:color": color !== undefined ? hexColorValue(color) : undefined,
  });
  return `<w:u ${a}/>`;
}

function eastAsianLayoutStr(opts: EastAsianLayoutOptions): string {
  const a = attrParts({
    "w:id": opts.id !== undefined ? decimalNumber(opts.id) : undefined,
    "w:combine": opts.combine !== undefined ? (opts.combine ? 1 : 0) : undefined,
    "w:combineBrackets": opts.combineBrackets,
    "w:vert": opts.vert !== undefined ? (opts.vert ? 1 : 0) : undefined,
    "w:vertCompress": opts.vertCompress !== undefined ? (opts.vertCompress ? 1 : 0) : undefined,
  });
  return `<w:eastAsianLayout ${a}/>`;
}

function languageStr(opts: LanguageOptions): string {
  const a = attrParts({
    "w:val": opts.value,
    "w:eastAsia": opts.eastAsia,
    "w:bidi": opts.bidirectional,
  });
  return `<w:lang ${a}/>`;
}

// ════════════════════════════════════════════════════════════════════════════
// Paragraph Properties
// ════════════════════════════════════════════════════════════════════════════

export interface StringifyPPrResult {
  xml: string | undefined;
  numberingReferences: { reference: string; instance: number }[];
}

/**
 * Build `<w:pPr>` XML string directly from options — zero IXmlableObject allocation.
 *
 * Replaces `buildParagraphProperties() + xml()` with a single-pass string builder.
 */
export function stringifyParagraphProperties(
  options?: ParagraphPropertiesOptions,
): StringifyPPrResult {
  const numberingReferences: { reference: string; instance: number }[] = [];

  if (!options) return { xml: undefined, numberingReferences };

  const parts: string[] = [];

  // Style / heading / bullet / numbering style references
  if (options.heading) {
    parts.push(`<w:pStyle w:val="${escapeXml(options.heading)}"/>`);
  }

  if (options.bullet) {
    parts.push('<w:pStyle w:val="ListParagraph"/>');
  }

  if (options.numbering) {
    if (!options.style && !options.heading) {
      if (!options.numbering.custom) {
        parts.push('<w:pStyle w:val="ListParagraph"/>');
      }
    }
  }

  if (options.style) {
    parts.push(`<w:pStyle w:val="${escapeXml(options.style)}"/>`);
  }

  // CT_PPrBase element order per XSD (wml.xsd) — strictly ordered sequence.
  // 1-4: keepNext, keepLines, pageBreakBefore
  if (options.keepNext !== undefined) parts.push(onOff("w:keepNext", options.keepNext));
  if (options.keepLines !== undefined) parts.push(onOff("w:keepLines", options.keepLines));
  if (options.pageBreakBefore !== undefined)
    parts.push(onOff("w:pageBreakBefore", options.pageBreakBefore));

  // 5: framePr
  if (options.frame) parts.push(framePrStr(options.frame));

  // 6: widowControl
  if (options.widowControl !== undefined) parts.push(onOff("w:widowControl", options.widowControl));

  // 7: numPr
  if (options.bullet) {
    parts.push(
      `<w:numPr><w:ilvl w:val="${Math.min(options.bullet.level, 9)}"/><w:numId w:val="1"/></w:numPr>`,
    );
  }

  if (options.numbering) {
    numberingReferences.push({
      instance: options.numbering.instance ?? 0,
      reference: options.numbering.reference,
    });

    const numId = `${options.numbering.reference}-${options.numbering.instance ?? 0}`;
    parts.push(numPrStr(numId, options.numbering.level, options.numbering.numberingChange));
  } else if (options.numbering === false) {
    parts.push(numPrStr(0, 0));
  }

  // 8: suppressLineNumbers
  if (options.suppressLineNumbers !== undefined)
    parts.push(onOff("w:suppressLineNumbers", options.suppressLineNumbers));

  // 9: pBdr
  if (options.border) {
    const bParts: string[] = [];
    if (options.border.top) bParts.push(borderStr("w:top", options.border.top));
    if (options.border.left) bParts.push(borderStr("w:left", options.border.left));
    if (options.border.bottom) bParts.push(borderStr("w:bottom", options.border.bottom));
    if (options.border.right) bParts.push(borderStr("w:right", options.border.right));
    if (options.border.between) bParts.push(borderStr("w:between", options.border.between));
    if (options.border.bar) bParts.push(borderStr("w:bar", options.border.bar));
    if (bParts.length) parts.push(`<w:pBdr>${bParts.join("")}</w:pBdr>`);
  }

  if (options.thematicBreak) {
    parts.push(
      `<w:pBdr>${borderStr("w:bottom", { color: "auto", size: 6, space: 1, style: BorderStyle.SINGLE })}</w:pBdr>`,
    );
  }

  // 10: shd
  if (options.shading) parts.push(shadingStr(options.shading));

  // 11: tabs
  const tabDefs: TabStopDefinition[] = [
    ...(options.rightTabStop !== undefined
      ? [{ position: options.rightTabStop, type: "right" as const }]
      : []),
    ...(options.tabStops ? options.tabStops : []),
    ...(options.leftTabStop !== undefined
      ? [{ position: options.leftTabStop, type: "left" as const }]
      : []),
  ];
  if (tabDefs.length > 0) parts.push(tabStopsStr(tabDefs));

  // 12-18: suppressAutoHyphens, kinsoku, wordWrap, overflowPunct, topLinePunct, autoSpaceDE, autoSpaceDN
  if (options.suppressAutoHyphens !== undefined)
    parts.push(onOff("w:suppressAutoHyphens", options.suppressAutoHyphens));
  if (options.kinsoku !== undefined) parts.push(onOff("w:kinsoku", options.kinsoku));
  if (options.wordWrap !== undefined) parts.push(onOff("w:wordWrap", options.wordWrap));
  if (options.overflowPunctuation)
    parts.push(onOff("w:overflowPunct", options.overflowPunctuation));
  if (options.topLinePunct !== undefined) parts.push(onOff("w:topLinePunct", options.topLinePunct));
  if (options.autoSpaceDE !== undefined) parts.push(onOff("w:autoSpaceDE", options.autoSpaceDE));
  if (options.autoSpaceEastAsianText !== undefined)
    parts.push(onOff("w:autoSpaceDN", options.autoSpaceEastAsianText));

  // 19: bidi
  if (options.bidirectional !== undefined) parts.push(onOff("w:bidi", options.bidirectional));

  // 20-21: adjustRightInd, snapToGrid
  if (options.adjustRightInd !== undefined)
    parts.push(onOff("w:adjustRightInd", options.adjustRightInd));
  if (options.snapToGrid !== undefined) parts.push(onOff("w:snapToGrid", options.snapToGrid));

  // 22-24: spacing, ind, contextualSpacing
  if (options.spacing) parts.push(spacingStr(options.spacing));
  if (options.indent) parts.push(indentStr(options.indent));
  if (options.contextualSpacing !== undefined)
    parts.push(onOff("w:contextualSpacing", options.contextualSpacing));

  // 25-26: mirrorIndents, suppressOverlap
  if (options.mirrorIndents !== undefined)
    parts.push(onOff("w:mirrorIndents", options.mirrorIndents));
  if (options.suppressOverlap !== undefined)
    parts.push(onOff("w:suppressOverlap", options.suppressOverlap));

  // 27: jc
  if (options.alignment) parts.push(`<w:jc w:val="${options.alignment}"/>`);

  // 28-30: textDirection, textAlignment, textboxTightWrap
  if (options.textDirection !== undefined)
    parts.push(`<w:textDirection w:val="${options.textDirection}"/>`);
  if (options.textAlignment !== undefined)
    parts.push(`<w:textAlignment w:val="${options.textAlignment}"/>`);
  if (options.textboxTightWrap !== undefined)
    parts.push(`<w:textboxTightWrap w:val="${options.textboxTightWrap}"/>`);

  // 31-33: outlineLvl, divId, cnfStyle
  if (options.outlineLevel !== undefined)
    parts.push(`<w:outlineLvl w:val="${options.outlineLevel}"/>`);
  if (options.divId !== undefined) parts.push(`<w:divId w:val="${options.divId}"/>`);
  if (options.cnfStyle) parts.push(cnfStyleStr(options.cnfStyle));

  // Embedded run properties (w:rPr inside w:pPr)
  if (options.run) {
    const inner = stringifyRunPropertiesInner(options.run);
    if (inner !== undefined) {
      const extra: string[] = [];
      const runOpts = options.run as ParagraphRunPropertiesOptions;
      if (runOpts.insertion) {
        const { id, author, date } = runOpts.insertion;
        extra.push(`<w:ins w:id="${id}" w:author="${escapeXml(author)}" w:date="${date}"/>`);
      }
      if (runOpts.deletion) {
        const { id, author, date } = runOpts.deletion;
        extra.push(`<w:del w:id="${id}" w:author="${escapeXml(author)}" w:date="${date}"/>`);
      }
      const body = inner + extra.join("");
      parts.push(`<w:rPr>${body}</w:rPr>`);
    }
  }

  // Revision (pPrChange)
  if (options.revision) {
    const rev = options.revision;
    const { author: _a, date: _d, id: _i, ...originalProps } = rev;
    const inner = stringifyParagraphProperties({ ...originalProps, includeIfEmpty: true });
    parts.push(
      `<w:pPrChange w:author="${escapeXml(rev.author)}" w:date="${rev.date}" w:id="${rev.id}">${inner.xml ?? "<w:pPr/>"}</w:pPrChange>`,
    );
  }

  const body = parts.join("");
  const xml = options.includeIfEmpty || body.length > 0 ? `<w:pPr>${body}</w:pPr>` : undefined;
  return { xml, numberingReferences };
}

// ════════════════════════════════════════════════════════════════════════════
// Run Properties
// ════════════════════════════════════════════════════════════════════════════

/**
 * Build the inner content of `<w:rPr>` as a string.
 * Returns undefined if no properties are set.
 */
export function stringifyRunPropertiesInner(opts?: RunPropertiesOptions): string | undefined {
  if (!opts) return undefined;

  const parts: string[] = [];

  // Style
  if (opts.style) parts.push(`<w:rStyle w:val="${escapeXml(opts.style)}"/>`);

  // Font
  if (opts.font) {
    if (typeof opts.font === "string") {
      parts.push(runFontsStr(opts.font));
    } else if ("name" in opts.font) {
      parts.push(runFontsStr(opts.font.name, opts.font.hint));
    } else {
      parts.push(runFontsStr(opts.font));
    }
  }

  // Bold
  if (opts.bold !== undefined) parts.push(onOff("w:b", opts.bold));
  const bCs =
    (opts.boldComplexScript === undefined && opts.bold !== undefined) || opts.boldComplexScript;
  if (bCs !== undefined) parts.push(onOff("w:bCs", opts.boldComplexScript ?? opts.bold!));

  // Italic
  if (opts.italic !== undefined) parts.push(onOff("w:i", opts.italic));
  const iCs =
    (opts.italicComplexScript === undefined && opts.italic !== undefined) ||
    opts.italicComplexScript;
  if (iCs !== undefined) parts.push(onOff("w:iCs", opts.italicComplexScript ?? opts.italic!));

  // Caps
  if (opts.smallCaps !== undefined) {
    parts.push(onOff("w:smallCaps", opts.smallCaps));
  } else if (opts.allCaps !== undefined) {
    parts.push(onOff("w:caps", opts.allCaps));
  }

  // Strike
  if (opts.strike !== undefined) parts.push(onOff("w:strike", opts.strike));
  if (opts.doubleStrike !== undefined) parts.push(onOff("w:dstrike", opts.doubleStrike));
  if (opts.emboss !== undefined) parts.push(onOff("w:emboss", opts.emboss));
  if (opts.imprint !== undefined) parts.push(onOff("w:imprint", opts.imprint));
  if (opts.outline !== undefined) parts.push(onOff("w:outline", opts.outline));
  if (opts.shadow !== undefined) parts.push(onOff("w:shadow", opts.shadow));
  if (opts.webHidden !== undefined) parts.push(onOff("w:webHidden", opts.webHidden));
  if (opts.noProof !== undefined) parts.push(onOff("w:noProof", opts.noProof));
  if (opts.snapToGrid !== undefined) parts.push(onOff("w:snapToGrid", opts.snapToGrid));
  if (opts.vanish) parts.push(onOff("w:vanish", opts.vanish));

  // Color
  if (opts.color) parts.push(colorStr(opts.color));

  // Character spacing
  if (opts.characterSpacing) {
    parts.push(`<w:spacing w:val="${signedTwipsMeasureValue(opts.characterSpacing)}"/>`);
  }

  // Scale
  if (opts.scale !== undefined) parts.push(`<w:w w:val="${opts.scale}"/>`);

  // Kern — w:val="0" is meaningful (explicitly disables kerning), so emit
  // whenever the field is set rather than truthy-checking it.
  if (opts.kern !== undefined) parts.push(`<w:kern w:val="${hpsMeasureValue(opts.kern)}"/>`);

  // Position
  if (opts.position) parts.push(`<w:position w:val="${opts.position}"/>`);

  // Size (points → half-points)
  if (opts.size !== undefined) parts.push(`<w:sz w:val="${hpsMeasureValue(opts.size * 2)}"/>`);
  const szCs =
    opts.sizeComplexScript === undefined || opts.sizeComplexScript === true
      ? opts.size
      : opts.sizeComplexScript;
  if (szCs) parts.push(`<w:szCs w:val="${hpsMeasureValue((szCs as number) * 2)}"/>`);

  // Highlight
  if (opts.highlight) parts.push(`<w:highlight w:val="${opts.highlight}"/>`);
  if (opts.highlightComplexScript === true) {
    if (opts.highlight) parts.push(`<w:highlightCs w:val="${opts.highlight}"/>`);
  } else if (opts.highlightComplexScript !== undefined && opts.highlightComplexScript !== false) {
    parts.push(`<w:highlightCs w:val="${opts.highlightComplexScript}"/>`);
  }

  // Underline
  if (opts.underline) parts.push(underlineStr(opts.underline.type, opts.underline.color));

  // Effect
  if (opts.effect) parts.push(`<w:effect w:val="${opts.effect}"/>`);

  // Border
  if (opts.border) parts.push(borderStr("w:bdr", opts.border));

  // Shading
  if (opts.shading) parts.push(shadingStr(opts.shading));

  // Vertical alignment
  if (opts.subScript) parts.push('<w:vertAlign w:val="subscript"/>');
  if (opts.superScript) parts.push('<w:vertAlign w:val="superscript"/>');

  // RTL
  if (opts.rightToLeft !== undefined) parts.push(onOff("w:rtl", opts.rightToLeft));

  // Emphasis mark
  if (opts.emphasisMark) parts.push(`<w:em w:val="${opts.emphasisMark.type ?? "dot"}"/>`);

  // Language
  if (opts.language) parts.push(languageStr(opts.language));

  // Spec vanish
  if (opts.specVanish) parts.push("<w:specVanish/>");

  // Math
  if (opts.math) parts.push(onOff("w:oMath", opts.math));

  // Fit text
  if (opts.fitText !== undefined) parts.push(`<w:fitText w:val="${opts.fitText}"/>`);

  // Complex script
  if (opts.complexScript !== undefined) parts.push(onOff("w:cs", opts.complexScript));

  // East Asian layout
  if (opts.eastAsianLayout) parts.push(eastAsianLayoutStr(opts.eastAsianLayout));

  // Content part
  if (opts.contentPartRId) parts.push(`<w:contentPart r:id="${opts.contentPartRId}"/>`);

  // Revision (rPrChange)
  if (opts.revision) {
    const rev = opts.revision as RunPropertiesChangeOptions;
    const { author: _a, date: _d, id: _i, ...originalProps } = rev;
    const inner = stringifyRunPropertiesInner(originalProps as RunPropertiesOptions);
    parts.push(
      `<w:rPrChange w:author="${escapeXml(rev.author)}" w:date="${rev.date}" w:id="${rev.id}"><w:rPr>${inner ?? ""}</w:rPr></w:rPrChange>`,
    );
  }

  // w14:* text effects — raw passthrough, emitted last (EG_RPrBase extension slot)
  if (opts.w14RawXml) parts.push(opts.w14RawXml);

  return parts.length > 0 ? parts.join("") : undefined;
}

/**
 * Build `<w:rPr>` XML string directly from options — zero IXmlableObject allocation.
 *
 * Replaces `buildRunProperties() + xml()` with a single-pass string builder.
 */
export function stringifyRunProperties(opts?: RunPropertiesOptions): string | undefined {
  const inner = stringifyRunPropertiesInner(opts);
  return inner ? `<w:rPr>${inner}</w:rPr>` : undefined;
}
