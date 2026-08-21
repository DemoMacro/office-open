/**
 * Cross-format text adapter — DrawingML paragraph (a:p, the core text model
 * shared by pptx/xlsx) ↔ WordprocessingML paragraph (w:p, docx).
 *
 * pptx and xlsx already model text as the core `a:p` shape, so this adapter only
 * bridges docx's w:p to/from that shared model. It lives in the aggregate
 * office-open convert layer alongside the other cross-format converters
 * (picture/shape/...): cross-format code references multiple format packages,
 * and the aggregate package is the single place that already depends on all of
 * them, so converters stay dependency-cycle-free. Single-format users never need
 * cross-format conversion.
 *
 * Round-trip is lossy by design, mirroring MS Office paste between apps.
 *
 * docx → a:p drops docx-only fields with no DrawingML text equivalent:
 *   paragraph: numbering, heading, borders, shading, keepNext/keepLines,
 *     bidirectional, widowControl, rsid, frame, outlineLevel, ...
 *   run: highlight, shading, border, kern, scale, position, effect,
 *     emphasisMark, w14 effects, language.eastAsia/bidirectional, ...
 *
 * a:p → docx drops DrawingML-only fields with no WordprocessingML equivalent:
 *   paragraph: defTabSize; bullet color/size/font/char/format (only the level
 *     survives); text fields (a:fld) are dropped.
 *   run: non-solid fill and non-sRGB color variants (gradient/pattern/blip/group
 *     fill; scheme/system/HSL/scRgb/preset colors → no w:p color);
 *     a:br run properties.
 *
 * Magnitude loss: a:p `baseline` is a signed percentage; docx has only
 * subscript/superscript on/off flags. Hyperlink round-trips url + tooltip only.
 *
 * Units (both APIs take plain numbers in their native unit):
 *   font size      points on both APIs (direct).
 *   char spacing   a:p spc (1/100 pt) ↔ w:p spacing (twips), ÷5 / ×5.
 *   before/after   a:p spcPts (1/100 pt) ↔ w:p spacing (twips), ÷5 / ×5.
 *   line spacing   a:p lineSpacingPercent (percent, 100 = single) ↔ w:p line +
 *                  lineRule "auto" (240 = single); a:p lineSpacingPoints (pt)
 *                  ↔ line + lineRule "exact" (×20).
 *   indents/tabs   a:p marL/marR/pos (EMU) ↔ w:p (twips), ÷635 / ×635.
 *
 * @module
 */

import { convertToTwip, stripColorHashPrefix } from "@office-open/core";
import type {
  FillOptions,
  ParagraphDescriptorOptions,
  TextParagraphPropertiesOptions as DrawingParagraphProperties,
  RunFont,
  TextRunOptions as DrawingRunOptions,
  TextCharacterPropertiesOptions as DrawingRunProperties,
  TextFont,
} from "@office-open/core";
import type { ParagraphOptions, RunOptions } from "@office-open/docx";

// ── unit factors ──

/** 1 inch = 914400 EMU = 1440 twips → 1 twip = 635 EMU. */
const EMU_PER_TWIP = 635;
/** 1 point = 100 hundredths = 20 twips → 1 hundredth = 0.2 twip. */
const TWIPS_PER_HUNDREDTH = 1 / 5;
const HUNDREDTHS_PER_TWIP = 5;
/** w:p "auto" line: 240 = single (100%). */
const AUTO_LINE_SINGLE = 240;
const POINTS_PER_TWIP = 1 / 20;

const round = Math.round;

/** Discriminant keys of ParagraphChild variants that are NOT a text run. */
const NON_RUN_KEYS = new Set([
  "hyperlink",
  "pageBreak",
  "columnBreak",
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
  "comment",
  "insertion",
  "deletion",
  "bookmarkStart",
  "bookmarkEnd",
  "bookmark",
  "wpsShape",
  "wpgGroup",
  "proofErr",
  "positionalTab",
  "permStart",
  "permEnd",
  "pageReference",
  "section",
  "symbol",
  "footnoteReference",
  "endnoteReference",
  "footnote",
  "endnote",
  "chart",
  "picture",
  "object",
  "sdt",
  "customXml",
  "pageNumber",
  "tableOfContents",
]);

// ── a:p → w:p ──

/**
 * Convert a DrawingML paragraph (core a:p) to a WordprocessingML paragraph (w:p).
 *
 * pptx/xlsx text is already a:p, so call this when pasting shape or cell text
 * into a docx. See module header for lossy fields.
 */
export function fromDrawingParagraph(drawing: ParagraphDescriptorOptions): ParagraphOptions {
  const docx: ParagraphOptions = {};

  const props = drawing.properties;
  if (props) {
    if (props.alignment) {
      const a = alignToDocx(props.alignment);
      if (a) docx.alignment = a;
    }

    const spacing = spacingToDocx(props);
    if (spacing) docx.spacing = spacing;

    const indent = indentToDocx(props);
    if (indent) docx.indent = indent;

    // Only fabricate a docx bullet when the source is actually bulleted; a bare
    // indentLevel (common in pptx placeholders) carries no bullet semantics.
    if (props.bullet && props.bullet.type !== "none") {
      docx.bullet = { level: props.indentLevel ?? 0 };
    }

    if (props.fontAlignment) {
      const t = fontAlignToDocx(props.fontAlignment);
      if (t) docx.textAlignment = t;
    }

    if (props.tabStops?.length) {
      const tabs = props.tabStops.map(tabToDocx);
      if (tabs.length) docx.tabStops = tabs;
    }
  }

  // Text shorthand takes priority when the source is a single text-only run.
  if (drawing.text !== undefined) {
    docx.text = drawing.text;
  } else if (drawing.children?.length) {
    const children = drawingToDocxChildren(drawing.children);
    if (children.length) docx.children = children;
  }

  return docx;
}

function drawingToDocxChildren(
  children: NonNullable<ParagraphDescriptorOptions["children"]>,
): NonNullable<ParagraphOptions["children"]> {
  const out: NonNullable<ParagraphOptions["children"]> = [];
  for (const child of children) {
    // String shorthand (core children allow bare strings) → one text run.
    if (typeof child === "string") {
      out.push({ text: child });
      continue;
    }
    // Soft break (a:br) → a run carrying w:br (count collapses to 1).
    if (typeof child === "object" && child !== null && "break" in child) {
      out.push({ break: 1 });
      continue;
    }
    // Text field (a:fld) — no w:p equivalent; dropped.
    if (typeof child === "object" && child !== null && "type" in child) {
      continue;
    }
    const run = child as DrawingRunOptions;
    // Run with an external hyperlink becomes a w:hyperlink child wrapping the
    // (de-hyperlinked) run so its run formatting survives.
    if (run.hyperlink?.url) {
      const { hyperlink, ...rest } = run;
      out.push({
        hyperlink: {
          url: hyperlink.url,
          ...(hyperlink.tooltip ? { tooltip: hyperlink.tooltip } : {}),
          children: [drawingRunToDocx(rest)],
        },
      });
      continue;
    }
    out.push(drawingRunToDocx(run));
  }
  return out;
}

function drawingRunToDocx(run: DrawingRunOptions): RunOptions {
  return {
    ...drawingRunPropertiesToDocx(run),
    ...(run.text !== undefined ? { text: run.text } : {}),
  };
}

function drawingRunPropertiesToDocx(run: DrawingRunProperties): Partial<RunOptions> {
  const out: Partial<RunOptions> = {};
  if (run.size !== undefined) out.size = run.size;
  if (run.bold !== undefined) out.bold = run.bold;
  if (run.italic !== undefined) out.italic = run.italic;
  if (run.underline && run.underline !== "none") out.underline = { type: run.underline };
  if (run.strike === "singleStrike") out.strike = true;
  else if (run.strike === "doubleStrike") out.doubleStrike = true;
  if (run.baseline !== undefined && run.baseline !== 0) {
    out.verticalAlign = run.baseline > 0 ? "superscript" : "subscript";
  }
  if (run.spacing !== undefined) out.characterSpacing = round(run.spacing * TWIPS_PER_HUNDREDTH);
  if (run.capitalization === "all") out.allCaps = true;
  else if (run.capitalization === "small") out.smallCaps = true;
  if (run.shadow) out.shadow = true;
  if (run.outline) out.outline = true;
  if (run.rightToLeft !== undefined) out.rightToLeft = run.rightToLeft;
  if (run.font !== undefined) out.font = drawingFontToDocx(run.font);
  if (run.fill !== undefined) {
    const hex = solidFillToHex(run.fill);
    if (hex) out.color = hex;
  }
  if (run.lang !== undefined) out.language = { value: run.lang };
  return out;
}

// ── w:p → a:p ──

/**
 * Convert a WordprocessingML paragraph (w:p) to a DrawingML paragraph (core a:p).
 *
 * Use this when pasting docx text (body or textbox) into a pptx/xlsx shape.
 * See module header for lossy fields.
 */
export function toDrawingParagraph(docx: ParagraphOptions): ParagraphDescriptorOptions {
  const drawing: ParagraphDescriptorOptions = {};

  const props = paragraphPropertiesToDrawing(docx);
  if (props) drawing.properties = props;

  // Prefer structured children over the text shorthand when both exist.
  if (docx.children?.length) {
    const children = docxToDrawingChildren(docx.children);
    if (children.length) drawing.children = children;
  } else if (docx.text !== undefined) {
    drawing.text = docx.text;
  }

  return drawing;
}

function docxToDrawingChildren(
  children: NonNullable<ParagraphOptions["children"]>,
): NonNullable<ParagraphDescriptorOptions["children"]> {
  const out: NonNullable<ParagraphDescriptorOptions["children"]> = [];
  for (const child of children) {
    if (typeof child === "string") {
      out.push({ text: child });
      continue;
    }
    if (typeof child !== "object" || child === null) continue;

    // External hyperlink child → flatten: each inner run inherits the link.
    if ("hyperlink" in child) {
      const hl = child.hyperlink;
      const url = hl.url;
      if (url === undefined) continue; // anchor-only/internal link — no a:p equivalent
      const link = { url, ...(hl.tooltip ? { tooltip: hl.tooltip } : {}) };
      const subs =
        hl.children && hl.children.length
          ? hl.children
          : child.text !== undefined
            ? [child.text]
            : [];
      for (const sub of subs) {
        if (typeof sub !== "string" && !isRunChild(sub)) continue;
        const run: DrawingRunOptions =
          typeof sub === "string" ? { text: sub } : docxRunToDrawing(sub);
        out.push({ ...run, hyperlink: link });
      }
      continue;
    }

    if (!isRunChild(child)) continue; // docx-only child (pageBreak, bookmark, …) dropped
    const run = child as RunOptions;
    if (run.text !== undefined) out.push(docxRunToDrawing(run));
    // A run carrying w:br becomes a soft break (count collapses to one).
    if (run.break) out.push({ break: true });
  }
  return out;
}

function docxRunToDrawing(run: RunOptions): DrawingRunOptions {
  const out: DrawingRunOptions = docxRunPropertiesToDrawing(run);
  if (run.text !== undefined) out.text = run.text;
  return out;
}

function docxRunPropertiesToDrawing(run: RunOptions): DrawingRunProperties {
  const out: DrawingRunProperties = {};
  if (run.size !== undefined) out.size = run.size;
  if (run.bold !== undefined) out.bold = run.bold;
  if (run.italic !== undefined) out.italic = run.italic;
  if (run.underline?.type) {
    // a:p models only single/double; other Word underline styles collapse to single.
    out.underline = run.underline.type === "double" ? "double" : "single";
  }
  if (run.doubleStrike) out.strike = "doubleStrike";
  else if (run.strike) out.strike = "singleStrike";
  if (run.verticalAlign === "superscript") out.baseline = 30000;
  else if (run.verticalAlign === "subscript") out.baseline = -25000;
  if (run.characterSpacing !== undefined) {
    out.spacing = round(convertToTwip(run.characterSpacing) * HUNDREDTHS_PER_TWIP);
  }
  if (run.allCaps) out.capitalization = "all";
  else if (run.smallCaps) out.capitalization = "small";
  if (run.shadow) out.shadow = true;
  if (run.outline) out.outline = true;
  if (run.rightToLeft !== undefined) out.rightToLeft = run.rightToLeft;
  if (run.font !== undefined) {
    const typeface = fontToString(run.font);
    if (typeface) out.font = typeface;
  }
  if (run.color !== undefined) {
    const hex = colorToHex(run.color);
    if (hex) out.fill = hex;
  }
  if (run.language?.value) out.lang = run.language.value;
  return out;
}

// ── paragraph property helpers ──

function paragraphPropertiesToDrawing(
  docx: ParagraphOptions,
): DrawingParagraphProperties | undefined {
  const out: DrawingParagraphProperties = {};
  if (docx.alignment) {
    const a = alignToDrawing(docx.alignment);
    if (a) out.alignment = a;
  }
  if (docx.spacing) {
    const sp = docx.spacing;
    if (sp.before !== undefined)
      out.spaceBefore = round(convertToTwip(sp.before) * HUNDREDTHS_PER_TWIP);
    if (sp.after !== undefined)
      out.spaceAfter = round(convertToTwip(sp.after) * HUNDREDTHS_PER_TWIP);
    if (sp.line !== undefined) {
      const twips = convertToTwip(sp.line);
      if (sp.lineRule === "auto") out.lineSpacingPercent = round((twips / AUTO_LINE_SINGLE) * 100);
      else out.lineSpacingPoints = round(twips * POINTS_PER_TWIP);
    }
  }
  if (docx.indent) {
    const start = docx.indent.start ?? docx.indent.left;
    const end = docx.indent.end ?? docx.indent.right;
    if (start !== undefined) out.marginIndent = round(convertToTwip(start) * EMU_PER_TWIP);
    if (end !== undefined) out.marginRight = round(convertToTwip(end) * EMU_PER_TWIP);
  }
  if (docx.bullet) {
    out.indentLevel = docx.bullet.level;
    out.bullet = { type: "char", char: "•" };
  }
  if (docx.textAlignment) {
    const f = fontAlignToDrawing(docx.textAlignment);
    if (f) out.fontAlignment = f;
  }
  if (docx.tabStops?.length) {
    const tabs = docx.tabStops
      .map(tabToDrawing)
      .filter((t): t is NonNullable<typeof t> => t !== undefined);
    if (tabs.length) out.tabStops = tabs;
  }
  return Object.keys(out).length ? out : undefined;
}

function spacingToDocx(
  props: DrawingParagraphProperties,
): NonNullable<ParagraphOptions["spacing"]> | undefined {
  const sp: NonNullable<ParagraphOptions["spacing"]> = {};
  if (props.spaceBefore !== undefined) sp.before = round(props.spaceBefore * TWIPS_PER_HUNDREDTH);
  if (props.spaceAfter !== undefined) sp.after = round(props.spaceAfter * TWIPS_PER_HUNDREDTH);
  // Percent line spacing (100 = single) takes precedence over the points form.
  if (props.lineSpacingPercent !== undefined) {
    sp.line = round((props.lineSpacingPercent / 100) * AUTO_LINE_SINGLE);
    sp.lineRule = "auto";
  } else if (props.lineSpacingPoints !== undefined) {
    sp.line = round(props.lineSpacingPoints / POINTS_PER_TWIP);
    sp.lineRule = "exact";
  }
  return Object.keys(sp).length ? sp : undefined;
}

function indentToDocx(
  props: DrawingParagraphProperties,
): NonNullable<ParagraphOptions["indent"]> | undefined {
  const ind: NonNullable<ParagraphOptions["indent"]> = {};
  if (props.marginIndent !== undefined) ind.start = round(props.marginIndent / EMU_PER_TWIP);
  if (props.marginRight !== undefined) ind.end = round(props.marginRight / EMU_PER_TWIP);
  return Object.keys(ind).length ? ind : undefined;
}

function tabToDocx(
  tab: NonNullable<DrawingParagraphProperties["tabStops"]>[number],
): NonNullable<ParagraphOptions["tabStops"]>[number] {
  return {
    type: tabAlignToDocx(tab.alignment),
    position: tab.position !== undefined ? round(tab.position / EMU_PER_TWIP) : 0,
  };
}

function tabToDrawing(
  tab: NonNullable<ParagraphOptions["tabStops"]>[number],
): NonNullable<DrawingParagraphProperties["tabStops"]>[number] | undefined {
  const alignment = tabAlignToDrawing(tab.type);
  if (!alignment) return undefined;
  const out: NonNullable<DrawingParagraphProperties["tabStops"]>[number] = { alignment };
  if (typeof tab.position === "number") out.position = round(tab.position * EMU_PER_TWIP);
  return out;
}

// ── enum / value mappers ──

function alignToDocx(a: DrawingParagraphProperties["alignment"]): ParagraphOptions["alignment"] {
  switch (a) {
    case "left":
      return "left";
    case "center":
      return "center";
    case "right":
      return "right";
    case "justify":
      return "both";
    default:
      return undefined;
  }
}

function alignToDrawing(a: ParagraphOptions["alignment"]): DrawingParagraphProperties["alignment"] {
  switch (a) {
    case "left":
    case "start":
      return "left";
    case "center":
      return "center";
    case "right":
    case "end":
      return "right";
    case "both":
      return "justify";
    default:
      return undefined;
  }
}

function fontAlignToDocx(
  f: DrawingParagraphProperties["fontAlignment"],
): ParagraphOptions["textAlignment"] {
  switch (f) {
    case "top":
      return "top";
    case "center":
      return "center";
    case "bottom":
      return "bottom";
    case "base":
      return "baseline";
    case "auto":
      return "auto";
    default:
      return undefined;
  }
}

function fontAlignToDrawing(
  t: ParagraphOptions["textAlignment"],
): DrawingParagraphProperties["fontAlignment"] {
  switch (t) {
    case "top":
      return "top";
    case "center":
      return "center";
    case "bottom":
      return "bottom";
    case "baseline":
      return "base";
    case "auto":
      return "auto";
    default:
      return undefined;
  }
}

type DrawingTabAlignment = NonNullable<DrawingParagraphProperties["tabStops"]>[number]["alignment"];
type DocxTabType = NonNullable<ParagraphOptions["tabStops"]>[number]["type"];

function tabAlignToDocx(a: DrawingTabAlignment): DocxTabType {
  switch (a) {
    case "left":
      return "left";
    case "center":
      return "center";
    case "right":
      return "right";
    case "decimal":
      return "decimal";
    default:
      return "left";
  }
}

function tabAlignToDrawing(t: DocxTabType): DrawingTabAlignment {
  switch (t) {
    case "left":
      return "left";
    case "center":
      return "center";
    case "right":
      return "right";
    case "decimal":
      return "decimal";
    default:
      return undefined;
  }
}

/** 6- or 8-digit sRGB hex; excludes scheme/system/preset color-name values. */
const SRGB_HEX = /^[0-9A-Fa-f]{6}([0-9A-Fa-f]{2})?$/;

/**
 * Extract an sRGB hex string from a solid fill; undefined for non-solid fills
 * and for non-RGB color variants (scheme/system/HSL/scRgb/preset) that have no
 * w:p color equivalent.
 */
function solidFillToHex(fill: FillOptions): string | undefined {
  if (typeof fill === "string") return stripColorHashPrefix(fill);
  if (fill.type === "solid") {
    const c = fill.color;
    if (typeof c === "string") return stripColorHashPrefix(c);
    if ("value" in c && typeof c.value === "string" && SRGB_HEX.test(c.value)) return c.value;
  }
  return undefined;
}

function colorToHex(color: NonNullable<RunOptions["color"]>): string | undefined {
  return typeof color === "string" ? stripColorHashPrefix(color) : color.val;
}

function fontToString(font: NonNullable<RunOptions["font"]>): string | undefined {
  if (typeof font === "string") return font;
  if ("name" in font) return font.name; // RunFontReference
  return font.ascii ?? font.hAnsi ?? font.eastAsia ?? font.complexScript; // FontProperties
}

/** Map a DrawingML RunFont to a docx run font: latin→ascii+hAnsi, eastAsia→eastAsia, complexScript→complexScript. symbol has no docx equivalent and is dropped. */
function drawingFontToDocx(font: RunFont): NonNullable<RunOptions["font"]> {
  if (typeof font === "string") return font;
  const typeface = (tf: TextFont | undefined): string | undefined =>
    tf === undefined ? undefined : typeof tf === "string" ? tf : tf.typeface;
  const latin = typeface(font.latin);
  const ea = typeface(font.eastAsia);
  const cs = typeface(font.complexScript);
  return {
    ...(latin ? { ascii: latin, hAnsi: latin } : {}),
    ...(ea ? { eastAsia: ea } : {}),
    ...(cs ? { complexScript: cs } : {}),
  };
}

function isRunChild(child: unknown): child is RunOptions {
  if (typeof child !== "object" || child === null) return false;
  for (const key of Object.keys(child)) {
    if (NON_RUN_KEYS.has(key)) return false;
  }
  return true;
}
