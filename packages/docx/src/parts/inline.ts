/**
 * Shared inline run/paragraph stringification for DOCX descriptors.
 *
 * Used by table.ts, comments.ts, body.ts, and other descriptors that need to
 * serialize paragraph/run content. Includes JSON child dispatch for all
 * ParagraphChild variants (image, chart, hyperlink, etc.).
 *
 * Pure string concatenation — zero IXmlableObject, zero BaseXmlComponent.
 *
 * @module
 */

import { toUint8Array } from "@office-open/core";
import { TargetModeType } from "@office-open/core";
import { uniqueId } from "@office-open/core";
import { chartSpaceDesc } from "@office-open/core/chart";
import type { SourceRectangleOptions } from "@office-open/core/drawingml";
import { createDataModel } from "@office-open/core/smartart";
import { escapeXml } from "@office-open/xml";
import type { BackgroundRawMediaOptions } from "@parts/document/document-background/document-background";
import type { MarkupRangeOptions, MoveRangeStartOptions } from "@parts/paragraph/links/bookmark";
import type { ParagraphChild, ParagraphOptions } from "@parts/paragraph/paragraph";
import type { CommentChildOptions } from "@parts/paragraph/run/comment-run";
import type { ImageOptions } from "@parts/paragraph/run/image-run";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { RubyOptions } from "@parts/paragraph/run/ruby";
import { breakXml, type RunOptions } from "@parts/paragraph/run/run";
import type { SmartArtOptions } from "@parts/paragraph/run/smartart-run";
import type {
  ChartMediaData,
  GroupChildMediaData,
  MediaData,
  SmartArtMediaData,
  WpgMediaData,
  WpsMediaData,
} from "@shared/media";
import { createTransformation } from "@shared/media";
import type { NonVisualPropertiesOptions } from "@shared/media/data";

import type { BodyContext } from "../context";
import { checkboxSymbolRunInner, stringifyCustomXmlShell, stringifySdtShell } from "./bodychildren";
import { drawingDesc } from "./drawing";
import { stringifyMath } from "./paragraph/math/stringify";
import { createBegin, createSeparate, createEnd } from "./paragraph/run/field";
import { stringifyParagraphProperties, stringifyRunProperties } from "./paragraph/stringify";

// ── Run ──

/** Serialize a deleted run: rPr + delText (or field delInstrText). */
function stringifyDeletedRun(c: RunOptions | string): string {
  const opts = typeof c === "string" ? { text: c } : c;
  const parts: string[] = [];
  const rPr = stringifyRunProperties(opts);
  if (rPr) parts.push(rPr);
  if (opts.break) parts.push(breakXml(opts.break));
  const fieldMap: Record<string, string> = {
    CURRENT: "PAGE",
    TOTAL_PAGES: "NUMPAGES",
    TOTAL_PAGES_IN_SECTION: "SECTIONPAGES",
  };
  if (opts.children) {
    for (const cc of opts.children) {
      if (typeof cc === "string") {
        // Page number fields use delInstrText instead of instrText
        const instrText = fieldMap[cc];
        if (instrText) {
          parts.push(
            '<w:fldChar w:fldCharType="begin"/>' +
              `<w:delInstrText xml:space="preserve">${instrText}</w:delInstrText>` +
              '<w:fldChar w:fldCharType="separate"/>' +
              '<w:fldChar w:fldCharType="end"/>',
          );
        } else {
          parts.push(`<w:delText xml:space="preserve">${escapeXml(cc)}</w:delText>`);
        }
      }
    }
  } else if (opts.text) {
    parts.push(`<w:delText xml:space="preserve">${escapeXml(String(opts.text))}</w:delText>`);
  }
  return `<w:r>${parts.join("")}</w:r>`;
}

export function stringifyRunInline(opts: RunOptions, ctx: BodyContext): string {
  const parts: string[] = [];

  const rPr = stringifyRunProperties(opts);
  if (rPr) parts.push(rPr);

  if (opts.break) parts.push(breakXml(opts.break));

  if (opts.children) {
    for (const child of opts.children) {
      if (typeof child === "string") {
        parts.push(`<w:t xml:space="preserve">${escapeXml(child)}</w:t>`);
      } else {
        // JSON child dispatch
        const jsonResult = stringifyChildDispatch(child as ParagraphChild, ctx);
        if (jsonResult !== undefined) {
          if (Array.isArray(jsonResult)) {
            parts.push(...jsonResult);
          } else {
            parts.push(jsonResult);
          }
        } else if ("text" in child || "children" in child || "break" in child) {
          parts.push(stringifyRunInline(child as RunOptions, ctx));
        }
      }
    }
  } else if (opts.text !== undefined) {
    parts.push(`<w:t xml:space="preserve">${escapeXml(String(opts.text))}</w:t>`);
  }

  const rsidAttrs: string[] = [];
  if (opts.rsid) rsidAttrs.push(` w:rsidR="${opts.rsid}"`);
  if (opts.runPropertiesRsid) rsidAttrs.push(` w:rsidRPr="${opts.runPropertiesRsid}"`);
  if (opts.deletionRsid) rsidAttrs.push(` w:rsidDel="${opts.deletionRsid}"`);
  const attr = rsidAttrs.join("");

  const body = parts.join("");
  return body.length === 0 ? (attr ? `<w:r${attr}/>` : "<w:r/>") : `<w:r${attr}>${body}</w:r>`;
}

// ── Image helpers ──

function createImageData(
  data: Uint8Array,
  transformation: ImageOptions["transformation"],
  key: string,
  sourceRectangle?: SourceRectangleOptions,
  nonVisualProperties?: NonVisualPropertiesOptions,
): Pick<
  MediaData,
  "data" | "fileName" | "transformation" | "sourceRectangle" | "nonVisualProperties"
> {
  return {
    data,
    fileName: key,
    sourceRectangle,
    nonVisualProperties,
    transformation: createTransformation(transformation),
  };
}

let nextChartId = 1;

// ── JSON child dispatch ──

/**
 * Stringify a ParagraphChild into one or more XML strings.
 *
 * Handles side effects (media, chart, smartArt, relationship registration)
 * directly without creating temporary class instances.
 *
 * Returns `undefined` if the child is not a recognized JSON wrapper.
 */
/**
 * Wrap a `<w:drawing>` run, rebuilding an mc:AlternateContent wrapper when a
 * VML fallback was carried from parse (Choice stays structured/editable,
 * Fallback round-trips as raw XML for fidelity).
 */
function wrapDrawingRun(
  drawingXml: string | undefined,
  opts: { vmlFallback?: string; mcChoiceRequires?: string; runProperties?: RunPropertiesOptions },
): string {
  const xml = drawingXml ?? "";
  const rPr = stringifyRunProperties(opts.runProperties) ?? "";
  if (opts.vmlFallback) {
    const requires = opts.mcChoiceRequires ?? "wps";
    // opts.vmlFallback is the serialized <mc:Fallback>…</mc:Fallback> element,
    // so splice it in directly (no extra wrapper).
    return `<w:r>${rPr}<mc:AlternateContent><mc:Choice Requires="${requires}">${xml}</mc:Choice>${opts.vmlFallback}</mc:AlternateContent></w:r>`;
  }
  return `<w:r>${rPr}${xml}</w:r>`;
}

/**
 * Register media carried by a VML fallback (mc:AlternateContent Fallback) so the
 * compiler resolves the fallback's `{fileName}` placeholders into rIds.
 *
 * A VML fallback image mirrors its Choice blip (same source bytes). When the
 * blip is already registered, reuse it and remap the fallback's `{fileName}`
 * placeholder to the shared media — matching Office, which emits one
 * relationship/file per image rather than a duplicate for the VML branch.
 */
function registerVmlFallbackMedia(
  opts: { vmlFallback?: string; vmlFallbackMedia?: BackgroundRawMediaOptions[] },
  ctx: BodyContext,
): void {
  if (!opts.vmlFallbackMedia) return;
  for (const m of opts.vmlFallbackMedia) {
    const data = toUint8Array(m.data);
    const entry = ctx.file.media.addMedia(
      data,
      m.type,
      (fileName) =>
        ({
          type: m.type,
          data,
          fileName,
          transformation: { emus: { x: 0, y: 0 }, pixels: { x: 0, y: 0 } },
        }) as MediaData,
      m.fileName,
    );
    // Dedup may reuse the Choice blip's file name; remap the VML fallback
    // placeholder so both branches share one relationship/file (matches Office).
    if (entry.fileName !== m.fileName && opts.vmlFallback) {
      opts.vmlFallback = opts.vmlFallback.split(`{${m.fileName}}`).join(`{${entry.fileName}}`);
    }
  }
}

/**
 * Build the rPr XML for a break/tab run from its structured run properties.
 */
/** Shared attribute string for CT_MarkupRange end markers (commentRange, move range end). */
function buildMarkupRangeAttrs(m: MarkupRangeOptions): string {
  const a: string[] = [`w:id="${m.id}"`];
  if (m.displacedByCustomXml) a.push(`w:displacedByCustomXml="${m.displacedByCustomXml}"`);
  return a.join(" ");
}

/** Shared attribute string for w:moveFromRangeStart / w:moveToRangeStart (CT_MoveBookmark). */
function buildMoveRangeStartAttrs(m: MoveRangeStartOptions): string {
  const a: string[] = [`w:id="${m.id}"`];
  if (m.name) a.push(`w:name="${escapeXml(m.name)}"`);
  if (m.author) a.push(`w:author="${escapeXml(m.author)}"`);
  if (m.date) a.push(`w:date="${m.date}"`);
  if (m.displacedByCustomXml) a.push(`w:displacedByCustomXml="${m.displacedByCustomXml}"`);
  if (m.colFirst !== undefined) a.push(`w:colFirst="${m.colFirst}"`);
  if (m.colLast !== undefined) a.push(`w:colLast="${m.colLast}"`);
  return a.join(" ");
}

/**
 * Expand a `{ comment }` sugar child: allocate the comment id, register the
 * comment entry (side effect, consumed when word/comments.xml is stringified),
 * and emit the range markers + anchored content + reference with one shared id.
 *
 * The caller never supplies an id — the library owns id allocation and pairing.
 */
function stringifyCommentChild(c: CommentChildOptions, ctx: BodyContext): string {
  const id = ctx.file.comments.nextId++;
  ctx.file.comments.entries.push({
    id,
    author: c.author,
    initials: c.initials,
    date: c.date,
    children: c.children,
  });

  // Anchored document content emitted between the range markers (inline runs/text).
  const wrapParts: string[] = [];
  for (const item of c.wrap ?? []) {
    wrapParts.push(
      typeof item === "string"
        ? stringifyRunInline({ text: item }, ctx)
        : stringifyRunInline(item, ctx),
    );
  }

  return (
    `<w:commentRangeStart w:id="${id}"/>` +
    wrapParts.join("") +
    `<w:commentRangeEnd w:id="${id}"/>` +
    `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`
  );
}

function runPropertiesXml(child: ParagraphChild): string {
  return stringifyRunProperties(child as RunOptions) ?? "";
}

export function stringifyChildDispatch(
  child: ParagraphChild,
  ctx: BodyContext,
): string | string[] | undefined {
  // Simple break types — pure XML, no side effects. A break run may carry run
  // properties (round-tripped from <w:r><w:rPr>…</w:rPr><w:br…/></w:r>).
  if ("pageBreak" in child) {
    return `<w:r>${runPropertiesXml(child)}<w:br w:type="page"/></w:r>`;
  }
  if ("columnBreak" in child) {
    return `<w:r>${runPropertiesXml(child)}<w:br w:type="column"/></w:r>`;
  }
  if ("tab" in child) {
    return `<w:r>${runPropertiesXml(child)}<w:tab/></w:r>`;
  }

  // Reference types — pure XML, no side effects
  if ("footnoteReference" in child) {
    const ref = child.footnoteReference;
    const id = typeof ref === "number" ? ref : ref.id;
    const cmf =
      typeof ref === "object" && ref.customMarkFollows ? ' w:customMarkFollows="true"' : "";
    return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"${cmf}/></w:r>`;
  }
  if ("endnoteReference" in child) {
    const ref = child.endnoteReference;
    const id = typeof ref === "number" ? ref : ref.id;
    const cmf =
      typeof ref === "object" && ref.customMarkFollows ? ' w:customMarkFollows="true"' : "";
    return `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"${cmf}/></w:r>`;
  }

  // Comment sugar — library allocates the id, emits the range markers +
  // reference, and registers the comment entry (see stringifyCommentChild).
  if ("comment" in child) return stringifyCommentChild(child.comment, ctx);

  // Comment markers — pure XML
  if ("commentRangeStart" in child)
    return `<w:commentRangeStart ${buildMarkupRangeAttrs(child.commentRangeStart)}/>`;
  if ("commentRangeEnd" in child)
    return `<w:commentRangeEnd ${buildMarkupRangeAttrs(child.commentRangeEnd)}/>`;
  if ("commentReference" in child)
    return `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${child.commentReference}"/></w:r>`;

  // Bookmark markers — pure XML
  if ("bookmarkStart" in child) {
    const bs = child.bookmarkStart;
    const a: string[] = [`w:id="${bs.id}"`, `w:name="${escapeXml(bs.name)}"`];
    if (bs.displacedByCustomXml) a.push(`w:displacedByCustomXml="${bs.displacedByCustomXml}"`);
    if (bs.colFirst !== undefined) a.push(`w:colFirst="${bs.colFirst}"`);
    if (bs.colLast !== undefined) a.push(`w:colLast="${bs.colLast}"`);
    return `<w:bookmarkStart ${a.join(" ")}/>`;
  }
  if ("bookmarkEnd" in child) {
    const be = child.bookmarkEnd;
    const a: string[] = [`w:id="${be.id}"`];
    if (be.displacedByCustomXml) a.push(`w:displacedByCustomXml="${be.displacedByCustomXml}"`);
    return `<w:bookmarkEnd ${a.join(" ")}/>`;
  }

  // Symbol run — direct XML output.
  // <w:sym> is a self-closing element, not text: emit it directly so it is
  // not escaped into a <w:t> by the run children path.
  if ("symbolRun" in child) {
    const opts = child.symbolRun;
    const rPr = stringifyRunProperties(opts) ?? "";
    return `<w:r>${rPr}<w:sym w:char="${opts.char}" w:font="${opts.symbolfont ?? "Wingdings"}"/></w:r>`;
  }

  // Form field (checkbox / dropdown list / text input) — fldChar sequence.
  // Word needs the field code (instrText) between begin and separate to
  // recognize the field type and render its result.
  if ("formField" in child) {
    const ff = child.formField;
    let result = "";
    let instrCode = "";
    let symbolFont = false;
    if (ff.checkBox) {
      result = ff.checkBox.checked ? "☒" : "☐";
      instrCode = "FORMCHECKBOX";
      // U+2610/U+2612 are absent from common body fonts (Calibri/Times);
      // MS Gothic holds them and matches the SDT w14:checkbox default.
      symbolFont = true;
    } else if (ff.dropDownList) {
      const idx = ff.dropDownList.result ?? ff.dropDownList.default;
      result = idx !== undefined ? (ff.dropDownList.entries[idx] ?? "") : "";
      instrCode = "FORMDROPDOWN";
    } else if (ff.textInput) {
      // Prefer the user-entered value (result run) over the placeholder default.
      result = ff.textInput.value ?? ff.textInput.default ?? "";
      instrCode = "FORMTEXT";
    }
    const rPr = symbolFont
      ? '<w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr>'
      : "";
    return (
      `<w:r>${createBegin(false, ff)}</w:r>` +
      `<w:r><w:instrText xml:space="preserve"> ${instrCode} </w:instrText></w:r>` +
      `<w:r>${createSeparate()}</w:r>` +
      `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(result)}</w:t></w:r>` +
      `<w:r>${createEnd()}</w:r>`
    );
  }

  // Image — side effect: media registration (content-deduplicated via core Media)
  if ("image" in child) {
    const opts = child.image;
    const rawData = toUint8Array(opts.data) as Uint8Array;

    let mediaData: MediaData;
    if (opts.type === "svg") {
      const fallbackData = toUint8Array(opts.fallback.data) as Uint8Array;
      const fallbackType = opts.fallback.type;
      // Register the raster fallback first so its file name is allocated, then
      // build the svg entry referencing it. Dedup applies to both independently.
      const fallback = ctx.file.media.addMedia(
        fallbackData,
        fallbackType,
        (fileName) =>
          ({
            type: fallbackType,
            ...createImageData(fallbackData, opts.transformation, fileName),
          }) as MediaData,
      );
      mediaData = ctx.file.media.addMedia(
        rawData,
        "svg",
        (fileName) =>
          ({
            type: "svg" as const,
            ...createImageData(
              rawData,
              opts.transformation,
              fileName,
              opts.sourceRectangle,
              opts.nonVisualProperties,
            ),
            useLocalDpi: opts.useLocalDpi,
            fallback,
          }) as MediaData,
      );
    } else {
      const type = opts.type;
      mediaData = ctx.file.media.addMedia(
        rawData,
        type,
        (fileName) =>
          ({
            type,
            ...createImageData(
              rawData,
              opts.transformation,
              fileName,
              opts.sourceRectangle,
              opts.nonVisualProperties,
            ),
            useLocalDpi: opts.useLocalDpi,
          }) as MediaData,
      );
    }

    // Build drawing XML via descriptor (zero XmlComponent instances)
    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
        outline: opts.outline,
        fill: opts.fill,
        effects: opts.effects,
        blipEffects: opts.blipEffects,
        tile: opts.tile,
        graphicFrameLocks: opts.graphicFrameLocks,
      },
      ctx,
    );
    return wrapDrawingRun(drawingXml, opts);
  }

  // Chart — side effect: chart registration
  if ("chart" in child) {
    const opts = child.chart;
    const chartKey = `chart_${nextChartId++}`;
    const mediaData: ChartMediaData = {
      chartKey,
      transformation: createTransformation(opts.transformation),
      type: "chart",
    };

    // Register chart
    const chartXml = chartSpaceDesc.stringify(
      {
        categories: opts.categories,
        series: opts.series,
        showLegend: opts.showLegend,
        style: opts.style,
        title: opts.title,
        type: opts.type,
        threeD: opts.threeD,
      },
      ctx.file,
    );
    ctx.file.charts.addChart(chartKey, {
      key: chartKey,
      chartSpaceXml: chartXml ?? "",
    });

    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
      },
      ctx,
    );
    return `<w:r>${drawingXml}</w:r>`;
  }

  // SmartArt — side effect: smartArt registration
  if ("smartArt" in child) {
    const opts = child.smartArt;
    const hash = hashSmartArtData(opts);
    const smartArtKey = `smartart_${hash}`;
    const mediaData: SmartArtMediaData = {
      smartArtKey,
      transformation: createTransformation(opts.transformation),
      type: "smartart",
    };

    // Register SmartArt
    const layoutId = opts.layout ?? "default";
    const styleId = opts.style ?? "simple1";
    const colorId = opts.color ?? "accent1_2";
    const dataModelXml = createDataModel(opts.data.nodes, layoutId, styleId, colorId);

    ctx.file.smartArts.addSmartArt(smartArtKey, {
      dataModelXml,
      key: smartArtKey,
      layout: layoutId,
      style: styleId,
      color: colorId,
    });

    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
      },
      ctx,
    );
    return `<w:r>${drawingXml}</w:r>`;
  }

  // WPS Shape (WordProcessing Shape) — side effect: blip fill media registration
  if ("wpsShape" in child) {
    const opts = child.wpsShape;
    const mediaData: WpsMediaData = {
      data: opts,
      transformation: createTransformation(opts.transformation),
      type: "wps",
    };

    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
        outline: opts.outline,
        fill: opts.fill,
        graphicFrameLocks: opts.graphicFrameLocks,
      },
      ctx,
    );
    registerVmlFallbackMedia(opts, ctx);
    return wrapDrawingRun(drawingXml, opts);
  }

  // WPG Group (WordProcessing Group) — group of shapes/pictures
  if ("wpgGroup" in child) {
    const opts = child.wpgGroup;
    const mediaData: WpgMediaData = {
      children: opts.children,
      transformation: createTransformation(opts.transformation),
      childOffset: opts.childOffset,
      childExtent: opts.childExtent,
      fill: opts.fill,
      effects: opts.effects,
      groupShapeLocks: opts.groupShapeLocks,
      type: "wpg",
    };

    // Register pic children media so {fileName} placeholders resolve, recursing
    // into nested wpg groups. wps children carry shape data, not media.
    const registerMedia = (children: readonly GroupChildMediaData[]): void => {
      for (const c of children) {
        if (c.type === "wps") continue;
        if (c.type === "wpg") {
          registerMedia(c.children);
          continue;
        }
        ctx.file.media.addMedia(c.data, c.type, () => c as MediaData, c.fileName);
      }
    };
    registerMedia(opts.children);

    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
        graphicFrameLocks: opts.graphicFrameLocks,
      },
      ctx,
    );
    registerVmlFallbackMedia(opts, ctx);
    return wrapDrawingRun(drawingXml, opts);
  }

  // Ruby annotation — pure string concatenation
  if ("ruby" in child && typeof child.ruby === "object" && child.ruby !== null) {
    const r = child.ruby as RubyOptions;
    const align = r.alignment ?? "center";
    const hps = (r.fontSize ?? 10) * 2;
    const hpsRaise = (r.raise ?? 10) * 2;
    const hpsBaseText = (r.baseFontSize ?? 20) * 2;
    const lid = r.languageId ?? "ja-JP";

    const prParts = [
      `<w:rubyAlign w:val="${align}"/>`,
      `<w:hps w:val="${hps}"/>`,
      `<w:hpsRaise w:val="${hpsRaise}"/>`,
      `<w:hpsBaseText w:val="${hpsBaseText}"/>`,
      `<w:lid w:val="${lid}"/>`,
    ];
    if (r.dirty) prParts.push("<w:dirty/>");

    const rt = `<w:rt><w:r><w:t xml:space="preserve">${escapeXml(r.text)}</w:t></w:r></w:rt>`;
    const rubyBase = `<w:rubyBase><w:r><w:t xml:space="preserve">${escapeXml(r.base)}</w:t></w:r></w:rubyBase>`;

    return `<w:ruby><w:rubyPr>${prParts.join("")}</w:rubyPr>${rt}${rubyBase}</w:ruby>`;
  }

  // Math — pure string concatenation
  if ("math" in child && typeof child.math === "object" && child.math !== null) {
    const mathOpts = child.math;
    const children = mathOpts.children ?? [];
    return stringifyMath(children);
  }

  // Inserted text run(s) — w:ins wraps one or more runs (CT_RunTrackChange)
  if ("insertion" in child) {
    const { id, author, date, children } = child.insertion;
    const body = children
      .map((c) => stringifyRunInline(typeof c === "string" ? { text: c } : c, ctx))
      .join("");
    return `<w:ins w:id="${id}" w:author="${escapeXml(String(author))}" w:date="${date}">${body}</w:ins>`;
  }

  // Deleted text run(s) — w:del wraps one or more runs (delText content)
  if ("deletion" in child) {
    const { id, author, date, children } = child.deletion;
    const body = children.map((c) => stringifyDeletedRun(c)).join("");
    return `<w:del w:id="${id}" w:author="${escapeXml(String(author))}" w:date="${date}">${body}</w:del>`;
  }

  // Hyperlink — side effect: relationship registration
  if ("hyperlink" in child) {
    const hl = child.hyperlink;

    // Serialize children using stringifyRunInline. A top-level `text` is a
    // shorthand for a single text run; without it `{ text, hyperlink }` would
    // emit an empty <w:hyperlink>.
    const childParts: string[] = [];
    if (child.text !== undefined) {
      childParts.push(stringifyRunInline({ text: child.text }, ctx));
    }
    if (hl.children) {
      for (const rc of hl.children) {
        if (typeof rc === "string") {
          childParts.push(stringifyRunInline({ text: rc }, ctx));
        } else {
          childParts.push(stringifyRunInline(rc, ctx));
        }
      }
    }
    const body = childParts.join("");

    const pushHlAttrs = (attrs: string[]): void => {
      if (hl.history !== false) attrs.push('w:history="1"');
      if (hl.tooltip) attrs.push(`w:tooltip="${escapeXml(hl.tooltip)}"`);
      if (hl.tgtFrame) attrs.push(`w:tgtFrame="${escapeXml(hl.tgtFrame)}"`);
      if (hl.docLocation) attrs.push(`w:docLocation="${escapeXml(hl.docLocation)}"`);
    };
    if (hl.link) {
      const linkId = uniqueId();
      ctx.viewWrapper.relationships.addRelationship(
        linkId,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        hl.link,
        TargetModeType.EXTERNAL,
      );
      const attrs = [`r:id="rId${linkId}"`];
      pushHlAttrs(attrs);
      return `<w:hyperlink ${attrs.join(" ")}>${body}</w:hyperlink>`;
    }
    if (hl.anchor) {
      const attrs = [`w:anchor="${escapeXml(hl.anchor)}"`];
      pushHlAttrs(attrs);
      return `<w:hyperlink ${attrs.join(" ")}>${body}</w:hyperlink>`;
    }
    return "";
  }

  // ── Proof error markers ──
  if ("proofErr" in child) return `<w:proofErr w:type="${child.proofErr}"/>`;

  // ── Positional tab ──
  if ("positionalTab" in child) {
    const pt = child.positionalTab;
    return `<w:ptab w:alignment="${pt.alignment}" w:leader="${pt.leader}" w:relativeTo="${pt.relativeTo}"/>`;
  }

  // ── Permission range markers ──
  if ("permStart" in child) {
    const ps = child.permStart;
    const a: string[] = [`w:id="${ps.id}"`];
    if (ps.ed !== undefined) a.push(`w:ed="${escapeXml(String(ps.ed))}"`);
    if (ps.editGroup !== undefined) a.push(`w:edGrp="${ps.editGroup}"`);
    if (ps.colFirst !== undefined) a.push(`w:colFirst="${ps.colFirst}"`);
    if (ps.colLast !== undefined) a.push(`w:colLast="${ps.colLast}"`);
    return `<w:permStart ${a.join(" ")}/>`;
  }
  if ("permEnd" in child) return `<w:permEnd w:id="${child.permEnd}"/>`;

  // ── Move revision range markers ──
  if ("moveFromRangeStart" in child) {
    return `<w:moveFromRangeStart ${buildMoveRangeStartAttrs(child.moveFromRangeStart)}/>`;
  }
  if ("moveFromRangeEnd" in child)
    return `<w:moveFromRangeEnd ${buildMarkupRangeAttrs(child.moveFromRangeEnd)}/>`;
  if ("moveToRangeStart" in child) {
    return `<w:moveToRangeStart ${buildMoveRangeStartAttrs(child.moveToRangeStart)}/>`;
  }
  if ("moveToRangeEnd" in child)
    return `<w:moveToRangeEnd ${buildMarkupRangeAttrs(child.moveToRangeEnd)}/>`;

  // ── Move revision text runs ──
  if ("movedFrom" in child) {
    const { id, author, date, children } = child.movedFrom;
    const body = children
      .map((c) => stringifyRunInline(typeof c === "string" ? { text: c } : c, ctx))
      .join("");
    return `<w:moveFrom w:id="${id}" w:author="${escapeXml(String(author))}" w:date="${date}">${body}</w:moveFrom>`;
  }
  if ("movedTo" in child) {
    const { id, author, date, children } = child.movedTo;
    const body = children
      .map((c) => stringifyRunInline(typeof c === "string" ? { text: c } : c, ctx))
      .join("");
    return `<w:moveTo w:id="${id}" w:author="${escapeXml(String(author))}" w:date="${date}">${body}</w:moveTo>`;
  }

  // ── Custom XML range markers (track changes) ──
  if ("customXmlInsRangeStart" in child) {
    const o = child.customXmlInsRangeStart;
    return `<w:customXmlInsRangeStart w:id="${o.id}" w:author="${escapeXml(o.author)}"${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlInsRangeEnd" in child)
    return `<w:customXmlInsRangeEnd w:id="${child.customXmlInsRangeEnd}"/>`;
  if ("customXmlDelRangeStart" in child) {
    const o = child.customXmlDelRangeStart;
    return `<w:customXmlDelRangeStart w:id="${o.id}" w:author="${escapeXml(o.author)}"${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlDelRangeEnd" in child)
    return `<w:customXmlDelRangeEnd w:id="${child.customXmlDelRangeEnd}"/>`;
  if ("customXmlMoveFromRangeStart" in child) {
    const o = child.customXmlMoveFromRangeStart;
    return `<w:customXmlMoveFromRangeStart w:id="${o.id}" w:author="${escapeXml(o.author)}"${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlMoveFromRangeEnd" in child)
    return `<w:customXmlMoveFromRangeEnd w:id="${child.customXmlMoveFromRangeEnd}"/>`;
  if ("customXmlMoveToRangeStart" in child) {
    const o = child.customXmlMoveToRangeStart;
    return `<w:customXmlMoveToRangeStart w:id="${o.id}" w:author="${escapeXml(o.author)}"${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlMoveToRangeEnd" in child)
    return `<w:customXmlMoveToRangeEnd w:id="${child.customXmlMoveToRangeEnd}"/>`;

  // ── Simple field ──
  if ("simpleField" in child) {
    const sf = child.simpleField;
    const sfAttrs = [`w:instr="${escapeXml(sf.instruction)}"`];
    if (sf.fldLock !== undefined) sfAttrs.push(`w:fldLock="${sf.fldLock ? 1 : 0}"`);
    if (sf.dirty !== undefined) sfAttrs.push(`w:dirty="${sf.dirty ? 1 : 0}"`);
    if (sf.cachedValue !== undefined) {
      return `<w:fldSimple ${sfAttrs.join(" ")}><w:r><w:t>${escapeXml(sf.cachedValue)}</w:t></w:r></w:fldSimple>`;
    }
    return `<w:fldSimple ${sfAttrs.join(" ")}/>`;
  }

  // ── Complex field (PAGE/DATE/TOC/... — fldChar field without w:ffData) ──
  if ("complexField" in child) {
    const cf = child.complexField;
    // Run-properties: Word writes identical rPr across a field's runs. Apply
    // the captured control-run rPr to begin/instrText/separate/end and the
    // result-run rPr to the result (defaults to the control rPr when the
    // result had none, matching Word's uniform behavior).
    const ctrl = cf.rPrXml ?? "";
    const res = cf.resultRPrXml ?? ctrl;
    // `separate` + the result run are emitted only when there is a cached
    // result; a result-less field round-trips as begin/instrText/end.
    const resultXml =
      cf.result !== undefined
        ? `<w:r>${ctrl}<w:fldChar w:fldCharType="separate"/></w:r>` +
          `<w:r>${res}<w:t xml:space="preserve">${escapeXml(cf.result)}</w:t></w:r>`
        : "";
    return (
      `<w:r>${ctrl}<w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r>${ctrl}<w:instrText xml:space="preserve">${escapeXml(
        cf.instruction,
      )}</w:instrText></w:r>` +
      resultXml +
      `<w:r>${ctrl}<w:fldChar w:fldCharType="end"/></w:r>`
    );
  }

  // ── Sequential identifier (SEQ field) ──
  if ("seqIdentifier" in child) {
    const id = child.seqIdentifier;
    return (
      "<w:r>" +
      '<w:fldChar w:fldCharType="begin"/>' +
      `<w:instrText xml:space="preserve"> SEQ ${escapeXml(id)} </w:instrText>` +
      '<w:fldChar w:fldCharType="separate"/>' +
      '<w:fldChar w:fldCharType="end"/>' +
      "</w:r>"
    );
  }

  // ── Page reference (PAGEREF field) ──
  if ("pageReference" in child) {
    const pr = child.pageReference;
    let instr = ` PAGEREF ${escapeXml(pr.bookmarkId)} `;
    if (pr.hyperlink) instr += "\\h ";
    if (pr.useRelativePosition) instr += "\\p ";
    return (
      "<w:r>" +
      '<w:fldChar w:fldCharType="begin"/>' +
      `<w:instrText xml:space="preserve">${instr}</w:instrText>` +
      '<w:fldChar w:fldCharType="end"/>' +
      "</w:r>"
    );
  }

  // ── Bidirectional text containers ──
  if ("dir" in child) {
    const d = child.dir;
    const childXml = serializeDirChildren(d.children, ctx);
    return `<w:dir w:val="${d.val}">${childXml}</w:dir>`;
  }
  if ("bdo" in child) {
    const b = child.bdo;
    const childXml = serializeDirChildren(b.children, ctx);
    return `<w:bdo w:val="${b.val}">${childXml}</w:bdo>`;
  }

  // ── Smart tag ──
  if ("smartTag" in child) {
    const st = child.smartTag;
    const attrs: string[] = [];
    if (st.uri) attrs.push(`w:uri="${escapeXml(st.uri)}"`);
    attrs.push(`w:element="${escapeXml(st.element)}"`);

    const parts: string[] = [];
    if (st.properties?.length) {
      const propParts: string[] = [];
      for (const p of st.properties) {
        const pa: string[] = [];
        if (p.uri) pa.push(`w:uri="${escapeXml(p.uri)}"`);
        pa.push(`w:name="${escapeXml(p.name)}"`, `w:val="${escapeXml(p.val)}"`);
        propParts.push(`<w:attr ${pa.join(" ")}/>`);
      }
      parts.push(`<w:smartTagPr>${propParts.join("")}</w:smartTagPr>`);
    }
    if (st.children) {
      for (const c of st.children) {
        parts.push(
          typeof c === "string" ? stringifyRunInline({ text: c }, ctx) : stringifyRunInline(c, ctx),
        );
      }
    }
    return `<w:smartTag ${attrs.join(" ")}>${parts.join("")}</w:smartTag>`;
  }

  // ── Custom XML run (CT_CustomXmlRun) ──
  if ("customXml" in child) {
    const cx = child.customXml;
    const contentParts: string[] = [];
    if (cx.children) {
      for (const c of cx.children) {
        if (typeof c === "string") {
          contentParts.push(stringifyRunInline({ text: c }, ctx));
        } else {
          const jr = stringifyChildDispatch(c as ParagraphChild, ctx);
          if (jr !== undefined) {
            contentParts.push(Array.isArray(jr) ? jr.join("") : jr);
          } else {
            contentParts.push(stringifyRunInline(c as RunOptions, ctx));
          }
        }
      }
    }
    return stringifyCustomXmlShell(cx, contentParts.join(""));
  }

  // ── Inline structured document tag (CT_SdtRun) ──
  if ("sdt" in child) {
    const s = child.sdt;
    let contentXml = "";
    if (s.properties.checkbox) {
      // Inline checkbox: render the state symbol as a run (no <w:p> wrapper).
      contentXml = checkboxSymbolRunInner(s.properties.checkbox);
    } else if (s.children && s.children.length > 0) {
      const cparts: string[] = [];
      for (const c of s.children) {
        if (typeof c === "string") {
          cparts.push(stringifyRunInline({ text: c }, ctx));
        } else {
          const jr = stringifyChildDispatch(c as ParagraphChild, ctx);
          if (jr !== undefined) {
            cparts.push(Array.isArray(jr) ? jr.join("") : jr);
          } else if ("text" in c || "children" in c || "break" in c) {
            cparts.push(stringifyRunInline(c as RunOptions, ctx));
          }
        }
      }
      contentXml = cparts.join("");
    }
    return stringifySdtShell(s.properties, s.endProperties, contentXml);
  }

  return undefined;
}

/** Serialize children of Dir/Bdo containers. */
function serializeDirChildren(
  children: (RunOptions | string)[] | undefined,
  ctx: BodyContext,
): string {
  if (!children) return "";
  const parts: string[] = [];
  for (const c of children) {
    if (typeof c === "string") {
      parts.push(stringifyRunInline({ text: c }, ctx));
    } else {
      parts.push(stringifyRunInline(c, ctx));
    }
  }
  return parts.join("");
}

/** Hash SmartArt data for unique key generation (duplicated from SmartArtRun). */
function hashSmartArtData(options: SmartArtOptions): number {
  const data = JSON.stringify(options.data);
  let hash = 0;
  for (let i = 0; i < data.length; i++) {
    const char = data.charCodeAt(i);
    hash = ((hash << 5) - hash + char) | 0;
  }
  return Math.abs(hash);
}

// ── Paragraph ──

export function stringifyParagraphInline(
  opts: string | ParagraphOptions,
  ctx: BodyContext,
): string {
  const resolved: ParagraphOptions = typeof opts === "string" ? { text: opts } : opts;
  const parts: string[] = [];

  const props = stringifyParagraphProperties(resolved);
  if (props.xml) parts.push(props.xml);

  // Register numbering references from inline paragraphs (footnotes, endnotes, etc.)
  // so that concrete numbering instances are created and placeholders get resolved.
  if (props.numberingReferences.length > 0) {
    for (const ref of props.numberingReferences) {
      ctx.file.numbering.createConcreteNumberingInstance(ref.reference, ref.instance);
    }
  }

  if (resolved.text !== undefined) {
    parts.push(stringifyRunInline({ text: resolved.text }, ctx));
  }

  if (resolved.children) {
    for (const child of resolved.children) {
      if (typeof child === "string") {
        parts.push(stringifyRunInline({ text: child }, ctx));
      } else if (typeof child === "object" && child !== null) {
        // Try JSON child dispatch first (image, chart, hyperlink, etc.)
        const jsonResult = stringifyChildDispatch(child as ParagraphChild, ctx);
        if (jsonResult !== undefined) {
          if (Array.isArray(jsonResult)) {
            parts.push(...jsonResult);
          } else {
            parts.push(jsonResult);
          }
        } else if ("text" in child || "children" in child || "break" in child) {
          parts.push(stringifyRunInline(child as RunOptions, ctx));
        }
      }
    }
  }

  const body = parts.join("");
  return body ? `<w:p>${body}</w:p>` : "<w:p/>";
}
