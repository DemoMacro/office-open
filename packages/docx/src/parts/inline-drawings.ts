/**
 * Drawing-family paragraph children: picture, chart, SmartArt, WPS shape,
 * content part, and WPG group runs.
 *
 * Each branch registers its side effects on the write context (media bytes,
 * chart and diagram parts, VML fallback media), then stringifies the
 * `<w:drawing>` run through the drawing descriptor. Shared by the paragraph
 * and body-level JSON dispatches via stringifyDrawingChild.
 *
 * @module
 */

import type { DataType } from "@office-open/core";
import { encodeBase64, imageTypeFromPath, toUint8Array } from "@office-open/core";
import { buildUserShapesData, chartExSpaceDesc, chartSpaceDesc } from "@office-open/core/chart";
import type { ChartExSpaceOptions } from "@office-open/core/chart";
import {
  createDataModel,
  definitionId,
  stringifyDataModelOptions,
} from "@office-open/core/smartart";
import type { SmartArtRawParts } from "@office-open/core/smartart";
import { documentNamespaceDialect } from "@parts/document/document-attributes";
import type { BackgroundRawMediaOptions } from "@parts/document/document-background/document-background";
import type { ParagraphChild } from "@parts/paragraph/paragraph";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { SmartArtOptions } from "@parts/paragraph/run/smartart-run";
import { stringifyUnsupportedDrawing } from "@parts/paragraph/run/unsupported-drawing-run";
import type {
  ChartMediaData,
  ChartExMediaData,
  CoreMediaData,
  GroupChildMediaData,
  MediaData,
  RegularMediaData,
  SmartArtMediaData,
  GroupMediaData,
  ShapeMediaData,
  SvgMediaData,
  CanvasMediaData,
} from "@shared/media";
import { createTransformation } from "@shared/media";

import type { BodyContext } from "../context";
import { drawingDesc } from "./drawing";
import {
  takeChartSourcePath,
  takeChartSourceRelationships,
  takeSmartArtSourcePaths,
  takeSmartArtSourceRelationships,
  takeUserShapesSourcePath,
} from "./drawing/drawing-parse";
import { stringifyRunProperties } from "./paragraph/stringify";

let nextChartId = 1;
let nextChartExId = 1;
const drawingDialect = (ctx: BodyContext) => documentNamespaceDialect(ctx.fileData);

/**
 * Wrap a `<w:drawing>` run, rebuilding an mc:AlternateContent wrapper when a
 * VML fallback was carried from parse (Choice stays structured/editable,
 * Fallback round-trips as raw XML for fidelity).
 */
function wrapDrawingRun(
  drawingXml: string | undefined,
  opts: {
    vmlFallback?: string;
    mcChoiceRequires?: string;
    runProperties?: RunPropertiesOptions;
    additionRsid?: string;
    deletionRsid?: string;
    runPropertiesRsid?: string;
    lastRenderedPageBreak?: boolean;
  },
  // The remapped fallback copy from registerVmlFallbackMedia — spliced in
  // place of opts.vmlFallback so the caller's options object stays untouched.
  vmlFallback: string | undefined = opts.vmlFallback,
  bareInRun = false,
): string {
  const xml = drawingXml ?? "";
  const rPr = stringifyRunProperties(opts.runProperties) ?? "";
  const lrpb = opts.lastRenderedPageBreak ? "<w:lastRenderedPageBreak/>" : "";
  const runAttrs =
    (opts.additionRsid ? ` w:rsidR="${opts.additionRsid}"` : "") +
    (opts.deletionRsid ? ` w:rsidDel="${opts.deletionRsid}"` : "") +
    (opts.runPropertiesRsid ? ` w:rsidRPr="${opts.runPropertiesRsid}"` : "");
  if (vmlFallback) {
    const requires = opts.mcChoiceRequires ?? "wps";
    // vmlFallback is the serialized <mc:Fallback>…</mc:Fallback> element, so
    // splice it in directly (no extra wrapper).
    const content = `<mc:AlternateContent><mc:Choice Requires="${requires}">${xml}</mc:Choice>${vmlFallback}</mc:AlternateContent>`;
    return bareInRun ? content : `<w:r${runAttrs}>${rPr}${lrpb}${content}</w:r>`;
  }
  return bareInRun ? xml : `<w:r${runAttrs}>${rPr}${lrpb}${xml}</w:r>`;
}

/**
 * Register media carried by a VML fallback (mc:AlternateContent Fallback) so the
 * compiler resolves the fallback's `{fileName}` placeholders into rIds.
 *
 * A VML fallback image mirrors its Choice blip (same source bytes). When the
 * blip is already registered, reuse it and remap the fallback's `{fileName}`
 * placeholder to the shared media — matching Office, which emits one
 * relationship/file per image rather than a duplicate for the VML branch.
 *
 * Returns the fallback XML with the renames applied — a local copy, never the
 * caller's `opts.vmlFallback` (options objects are shared, serializable data;
 * mutating them would leak this run's media numbering into the next run's
 * input).
 */
function registerVmlFallbackMedia(
  opts: { vmlFallback?: string; vmlFallbackMedia?: BackgroundRawMediaOptions[] },
  ctx: BodyContext,
): string | undefined {
  const vml = opts.vmlFallback;
  if (!vml || !opts.vmlFallbackMedia) return vml;
  let out = vml;
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
          transformation: { width: 0, height: 0 },
        }) as MediaData,
      m.fileName,
    );
    // Dedup may reuse the Choice blip's file name; remap the VML fallback
    // placeholder so both branches share one relationship/file (matches Office).
    if (entry.fileName !== m.fileName) {
      out = out.split(`{${m.fileName}}`).join(`{${entry.fileName}}`);
    }
  }
  return out;
}

/** Content fingerprint of the raw source parts: every field participates so
 *  two instances that differ in any raw part (or in their companion rels or
 *  media names) never merge into one diagram part set. */
function rawSmartArtFingerprint(raw: SmartArtRawParts | undefined): string | undefined {
  if (!raw) return undefined;
  const b64 = (v: DataType | undefined) =>
    v === undefined ? undefined : encodeBase64(toUint8Array(v));
  return JSON.stringify({
    drawing: b64(raw.drawing),
    media: raw.media?.map(({ fileName, partPath }) => ({ fileName, partPath })),
  });
}

/** Clone a WPG/canvas child tree so media-registration mutations stay local. */
function cloneDrawingChild(child: GroupChildMediaData): GroupChildMediaData {
  return child.type === "wpg"
    ? { ...child, children: child.children.map(cloneDrawingChild) }
    : child.type === "svg"
      ? { ...child, fallback: { ...child.fallback } }
      : { ...child };
}

function registerChartEx(options: ChartExSpaceOptions, ctx: BodyContext): string {
  const key = `chartEx_${nextChartExId++}`;
  ctx.file.chartExes.addChartEx(key, {
    key,
    chartExXml: chartExSpaceDesc.stringify(options, ctx.file) ?? "",
    ...(takeChartSourcePath(options) ? { sourcePath: takeChartSourcePath(options) } : {}),
    ...(takeChartSourceRelationships(options)
      ? { sourceRelationships: takeChartSourceRelationships(options) }
      : {}),
  });
  return key;
}

/** Register child media and chart parts for WPG groups and WordprocessingCanvas. */
function registerDrawingChildMedia(
  children: readonly GroupChildMediaData[],
  ctx: BodyContext,
): void {
  const adopt =
    (media: MediaData): ((fileName: string) => MediaData) =>
    (fileName) => {
      media.fileName = fileName;
      return media;
    };
  for (const child of children) {
    if (child.type === "wps") continue;
    if (child.type === "wpg") {
      registerDrawingChildMedia(child.children, ctx);
      continue;
    }
    if (child.type === "chart") {
      if (child.chartOptions && !child.chartKey) {
        child.chartKey = `chart_${nextChartId++}`;
        const externalData = child.chartOptions.externalData;
        ctx.file.charts.addChart(child.chartKey, {
          key: child.chartKey,
          chartSpaceXml: chartSpaceDesc.stringify(child.chartOptions, ctx.file) ?? "",
          ...(takeChartSourcePath(child.chartOptions)
            ? { sourcePath: takeChartSourcePath(child.chartOptions) }
            : {}),
          ...(takeChartSourceRelationships(child.chartOptions)
            ? { sourceRelationships: takeChartSourceRelationships(child.chartOptions) }
            : {}),
          ...(externalData?.data !== undefined &&
          externalData.fileName !== undefined &&
          externalData.relationshipId !== undefined
            ? {
                embedding: {
                  relationshipId: externalData.relationshipId,
                  fileName: externalData.fileName,
                  data: externalData.data,
                  ...(externalData.relationshipType
                    ? { relationshipType: externalData.relationshipType }
                    : {}),
                },
              }
            : {}),
          ...(child.chartOptions.userShapes
            ? {
                userShapes: {
                  ...buildUserShapesData(child.chartOptions.userShapes),
                  ...(takeUserShapesSourcePath(child.chartOptions.userShapes)
                    ? { path: takeUserShapesSourcePath(child.chartOptions.userShapes) }
                    : {}),
                },
              }
            : {}),
        });
      }
      continue;
    }
    if (child.type === "chartEx") {
      if (child.chartExOptions && !child.chartExKey) {
        child.chartExKey = registerChartEx(child.chartExOptions, ctx);
      }
      continue;
    }
    if (child.type === "contentPart" || child.type === "unsupported") continue;
    if (child.type === "svg") {
      const fallback = child.fallback;
      const fallbackEntry = ctx.file.media.addMedia(
        fallback.data,
        fallback.type,
        adopt(fallback),
        fallback.fileName,
      );
      fallback.fileName = fallbackEntry.fileName;
      const svgEntry = ctx.file.media.addMedia(child.data, "svg", adopt(child), child.fileName);
      child.fileName = svgEntry.fileName;
      continue;
    }
    const entry = ctx.file.media.addMedia(child.data, child.type, adopt(child), child.fileName);
    child.fileName = entry.fileName;
  }
}

/** Hash SmartArt data for unique key generation (duplicated from SmartArtRun). */
function hashSmartArtData(options: SmartArtOptions): number {
  // Layout/style/color participate: two diagrams with the same nodes but
  // different definitions are distinct parts. Raw source bytes too: a
  // round-tripped document keeps one part set per source instance even when
  // two instances fold to identical structured options.
  const data = JSON.stringify({
    nodes: options.nodes,
    dataModel: options.dataModel,
    layout: options.layout,
    style: options.style,
    color: options.color,
    raw: rawSmartArtFingerprint(options.raw),
  });
  let hash = 0;
  for (let i = 0; i < data.length; i++) {
    const char = data.charCodeAt(i);
    hash = ((hash << 5) - hash + char) | 0;
  }
  return Math.abs(hash);
}

// ── Drawing-family dispatch ──

/**
 * Stringify the drawing-family children of the JSON dispatch: picture, chart,
 * SmartArt, WPS shape, content part, and WPG group. Each registers its media
 * or part side effects, then emits the `<w:drawing>` run.
 *
 * Returns `undefined` when the child is not a drawing-family wrapper.
 */
export function stringifyDrawingChild(
  child: ParagraphChild,
  ctx: BodyContext,
  { bareInRun = false }: { bareInRun?: boolean } = {},
): string | undefined {
  const wrap = (
    drawingXml: string | undefined,
    opts: Parameters<typeof wrapDrawingRun>[1],
    vmlFallback?: string,
  ): string => wrapDrawingRun(drawingXml, opts, vmlFallback, bareInRun);
  if ("unsupportedDrawing" in child) {
    const vmlFallback = registerVmlFallbackMedia(child.unsupportedDrawing, ctx);
    return wrap(
      stringifyUnsupportedDrawing(child.unsupportedDrawing, ctx),
      child.unsupportedDrawing,
      vmlFallback,
    );
  }
  // Picture — side effect: media registration (content-deduplicated via core Media)
  if ("picture" in child) {
    const opts = child.picture;

    // Linked-only picture (external URL, no bytes) — no media registration;
    // the blip carries r:link alone and the owning part gets one External
    // image relationship.
    if (opts.type !== "svg" && opts.data === undefined && opts.sourceUrl !== undefined) {
      const drawingXml = drawingDesc.stringify(
        {
          mediaData: {
            type: opts.type,
            sourceUrl: opts.sourceUrl,
            transformation: createTransformation(opts.transformation),
            dpi: opts.dpi,
            rotWithShape: opts.rotWithShape,
            sourceRectangle: opts.sourceRectangle,
            fillRectangle: opts.fillRectangle,
            stretch: opts.stretch,
            nonVisualProperties: opts.nonVisualProperties,
            useLocalDpi: opts.useLocalDpi,
            useLocalDpiUri: opts.useLocalDpiUri,
            compression: opts.compression,
            blackWhiteMode: opts.blackWhiteMode,
          },
          docProperties: opts.altText,
          floating: opts.floating,
          outline: opts.outline,
          fill: opts.fill,
          effects: opts.effects,
          geometry: opts.geometry,
          customGeometry: opts.customGeometry,
          scene3d: opts.scene3d,
          shape3d: opts.shape3d,
          blipEffects: opts.blipEffects,
          blipExt: opts.blipExt,
          shapeExtensions: opts.shapeExtensions,
          tile: opts.tile,
          graphicFrameLocks: opts.graphicFrameLocks,
          extensionIds: opts.extensionIds,
          dialect: drawingDialect(ctx),
        },
        ctx,
      );
      return wrap(drawingXml, opts);
    }

    // Data is required past this point: linked-only pictures returned above,
    // and toUint8Array(undefined) throws the clear TypeError for degenerate
    // authoring input (neither bytes nor a linked source).
    const rawData = toUint8Array(opts.data!, { encoding: "base64" }) as Uint8Array;

    // The dedup entry carries media identity only — a byte-identical image
    // reuses the first registrant's entry object, so per-reference placement
    // fields (size, crop, cNvPr, blip hints, the svg raster fallback) are
    // re-applied onto the returned entry instead of baked into the `build`
    // callback.
    let mediaData: MediaData;
    if (opts.type === "svg") {
      const fallbackData = toUint8Array(opts.fallback.data, { encoding: "base64" }) as Uint8Array;
      const fallbackType = opts.fallback.type;
      // Register the raster fallback first so its file name is allocated, then
      // build the svg entry referencing it. Dedup applies to both independently.
      // Media<MediaData> pins addMedia's entry type to the wide union — assert
      // each branch back to the narrow shape its slot declares.
      const fallback = ctx.file.media.addMedia(
        fallbackData,
        fallbackType,
        (fileName) => ({
          type: fallbackType,
          data: fallbackData,
          fileName,
          transformation: { width: 0, height: 0 },
        }),
        opts.fallback.fileName,
      ) as RegularMediaData & CoreMediaData;
      mediaData = {
        ...(ctx.file.media.addMedia(
          rawData,
          "svg",
          (fileName) =>
            ({
              type: "svg" as const,
              data: rawData,
              fileName,
              fallback,
              transformation: { width: 0, height: 0 },
            }) as SvgMediaData & CoreMediaData,
          opts.fileName,
        ) as SvgMediaData & CoreMediaData),
        fallback: fallback as RegularMediaData & CoreMediaData,
        transformation: createTransformation(opts.transformation),
        dpi: opts.dpi,
        rotWithShape: opts.rotWithShape,
        sourceRectangle: opts.sourceRectangle,
        fillRectangle: opts.fillRectangle,
        stretch: opts.stretch,
        nonVisualProperties: opts.nonVisualProperties,
        nonVisualPropertiesPresent: opts.nonVisualPropertiesPresent,
        shapePropertiesPresent: opts.shapePropertiesPresent,
        useLocalDpi: opts.useLocalDpi,
        useLocalDpiUri: opts.useLocalDpiUri,
        compression: opts.compression,
        blackWhiteMode: opts.blackWhiteMode,
      };
    } else {
      const type = opts.type;
      mediaData = {
        ...ctx.file.media.addMedia(
          rawData,
          type,
          (fileName) =>
            ({
              type,
              data: rawData,
              fileName,
              ...(opts.partPath ? { partPath: opts.partPath } : {}),
              transformation: { width: 0, height: 0 },
            }) as MediaData,
          opts.fileName,
        ),
        transformation: createTransformation(opts.transformation),
        dpi: opts.dpi,
        rotWithShape: opts.rotWithShape,
        sourceRectangle: opts.sourceRectangle,
        fillRectangle: opts.fillRectangle,
        stretch: opts.stretch,
        nonVisualProperties: opts.nonVisualProperties,
        nonVisualPropertiesPresent: opts.nonVisualPropertiesPresent,
        shapePropertiesPresent: opts.shapePropertiesPresent,
        useLocalDpi: opts.useLocalDpi,
        useLocalDpiUri: opts.useLocalDpiUri,
        compression: opts.compression,
        blackWhiteMode: opts.blackWhiteMode,
        sourceUrl: opts.sourceUrl,
      };
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
        geometry: opts.geometry,
        customGeometry: opts.customGeometry,
        scene3d: opts.scene3d,
        shape3d: opts.shape3d,
        blipEffects: opts.blipEffects,
        blipExt: opts.blipExt,
        shapeExtensions: opts.shapeExtensions,
        tile: opts.tile,
        graphicFrameLocks: opts.graphicFrameLocks,
        extensionIds: opts.extensionIds,
        dialect: drawingDialect(ctx),
      },
      ctx,
    );
    return wrap(drawingXml, opts);
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

    // Register chart — strip the DOCX anchor-wrapper fields and pass every
    // ChartSpaceOptions field through so round-tripped charts keep their
    // axes, spPr, dLbls, externalData, …
    const {
      transformation: _t,
      floating: _f,
      altText: _a,
      graphicFrameLocks: _g,
      runProperties: _r,
      additionRsid: _ri,
      deletionRsid: _rd,
      runPropertiesRsid: _rPrRsid,
      lastRenderedPageBreak: _l,
      ...chartSpace
    } = opts;
    const chartXml = chartSpaceDesc.stringify(chartSpace, ctx.file);
    const externalData = chartSpace.externalData;
    ctx.file.charts.addChart(chartKey, {
      key: chartKey,
      chartSpaceXml: chartXml ?? "",
      ...(takeChartSourcePath(opts) ? { sourcePath: takeChartSourcePath(opts) } : {}),
      ...(takeChartSourceRelationships(opts)
        ? { sourceRelationships: takeChartSourceRelationships(opts) }
        : {}),
      ...(externalData?.data !== undefined &&
      externalData.fileName !== undefined &&
      externalData.relationshipId !== undefined
        ? {
            embedding: {
              relationshipId: externalData.relationshipId,
              fileName: externalData.fileName,
              data: externalData.data,
              ...(externalData.relationshipType
                ? { relationshipType: externalData.relationshipType }
                : {}),
            },
          }
        : {}),
      ...(chartSpace.userShapes
        ? {
            userShapes: {
              ...buildUserShapesData(chartSpace.userShapes),
              ...(takeUserShapesSourcePath(chartSpace.userShapes)
                ? { path: takeUserShapesSourcePath(chartSpace.userShapes) }
                : {}),
            },
          }
        : {}),
    });

    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
        graphicFrameLocks: opts.graphicFrameLocks,
        extensionIds: opts.extensionIds,
        dialect: drawingDialect(ctx),
      },
      ctx,
    );
    return wrap(drawingXml, opts);
  }

  // ChartEx — side effect: independent chartEx part registration.
  if ("chartEx" in child) {
    const opts = child.chartEx;
    const chartExKey = registerChartEx(opts, ctx);
    const mediaData: ChartExMediaData = {
      chartExKey,
      transformation: createTransformation(opts.transformation),
      type: "chartEx",
    };
    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
        graphicFrameLocks: opts.graphicFrameLocks,
        extensionIds: opts.extensionIds,
        dialect: drawingDialect(ctx),
      },
      ctx,
    );
    return wrap(drawingXml, opts);
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

    // Register SmartArt — custom definitions embed their own id in the doc
    // point's type ids.
    const layoutId =
      typeof opts.layout === "object" ? definitionId(opts.layout) : (opts.layout ?? "default");
    const styleId =
      typeof opts.style === "object" ? definitionId(opts.style) : (opts.style ?? "simple1");
    const colorId =
      typeof opts.color === "object" ? definitionId(opts.color) : (opts.color ?? "accent1_2");
    const dataModelXml = opts.dataModel
      ? stringifyDataModelOptions(opts.dataModel, ctx)
      : createDataModel(opts.nodes, layoutId, styleId, colorId, ctx.reproducible);

    // Data-part companion images (dgm:pt blipFill art): register through the
    // media collection so name pinning and dedup match the picture path. A
    // pinned name taken by different bytes re-allocates — remap the verbatim
    // data rels target to follow, or the rels would point at the wrong bytes.
    let raw = opts.raw;
    if (raw?.media) {
      const media = [];
      for (const m of raw.media) {
        const data = toUint8Array(m.data);
        const type = imageTypeFromPath(m.fileName);
        const sourceDirectory = m.partPath?.split("/").slice(0, -1);
        const entry = ctx.file.media.addMedia(
          data,
          type,
          (fileName) =>
            ({
              type,
              data,
              fileName,
              ...(sourceDirectory ? { partPath: [...sourceDirectory, fileName].join("/") } : {}),
              transformation: { width: 0, height: 0 },
            }) as MediaData,
          m.fileName,
        );
        media.push({
          ...m,
          fileName: entry.fileName,
          ...(entry.partPath ? { partPath: entry.partPath } : {}),
        });
      }
      raw = { ...raw, media };
    }

    ctx.file.smartArts.addSmartArt(smartArtKey, {
      dataModelXml,
      key: smartArtKey,
      layout: opts.layout ?? "default",
      style: opts.style ?? "simple1",
      color: opts.color ?? "accent1_2",
      ...(takeSmartArtSourcePaths(opts) ? { sourcePaths: takeSmartArtSourcePaths(opts) } : {}),
      ...(takeSmartArtSourceRelationships(opts)
        ? { sourceRelationships: takeSmartArtSourceRelationships(opts) }
        : {}),
      // Store a shallow copy with the remapped rels — never mutate the
      // caller's Options object.
      ...(raw ? { raw } : {}),
    });

    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
        graphicFrameLocks: opts.graphicFrameLocks,
        extensionIds: opts.extensionIds,
        dialect: drawingDialect(ctx),
      },
      ctx,
    );
    return wrap(drawingXml, {
      runProperties: opts.runProperties,
      additionRsid: opts.additionRsid,
      deletionRsid: opts.deletionRsid,
      runPropertiesRsid: opts.runPropertiesRsid,
      lastRenderedPageBreak: opts.lastRenderedPageBreak,
    });
  }

  // WPS Shape (WordProcessing Shape) — side effect: blip fill media registration
  if ("wpsShape" in child) {
    const opts = child.wpsShape;
    const mediaData: ShapeMediaData = {
      data: opts,
      transformation: createTransformation(opts.transformation),
      type: "wps",
      graphicDataUri: opts.graphicDataUri,
    };

    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        docPropertiesPresent: opts.docPropertiesPresent,
        floating: opts.floating,
        outline: opts.outline,
        fill: opts.fill,
        graphicFrameLocks: opts.graphicFrameLocks,
        extensionIds: opts.extensionIds,
        dialect: drawingDialect(ctx),
      },
      ctx,
    );
    const vmlFallback = registerVmlFallbackMedia(opts, ctx);
    return wrap(drawingXml, opts, vmlFallback);
  }

  // Content part (w:contentPart) — run-level EG_RunInnerContent element (CT_Rel).
  // Word references ink and other opaque parts this way; the richer placement
  // fields of ContentPartOptions only apply inside a wpg group child.
  if ("contentPart" in child) {
    return bareInRun
      ? `<w:contentPart r:id="${child.contentPart.referenceId}"/>`
      : `<w:r><w:contentPart r:id="${child.contentPart.referenceId}"/></w:r>`;
  }

  // WPG Group (WordProcessing Group) — group of shapes/pictures
  if ("wpgGroup" in child) {
    const opts = child.wpgGroup;
    // Registration stamps per-run state onto the media data (dedup renames in
    // registerMedia, fresh chart keys), so stringify a shallow copy tree — the
    // caller's GroupOptions must stay untouched (options objects are shared,
    // serializable data; mutating them would leak this run's media numbering
    // into the next run's input, and a persisted chart key would skip the
    // chart registration of a later run entirely).
    const mediaData: GroupMediaData = {
      children: opts.children.map(cloneDrawingChild),
      transformation: createTransformation(opts.transformation),
      groupExtent: opts.groupExtent,
      childOffsetX: opts.childOffsetX,
      childOffsetY: opts.childOffsetY,
      childExtentWidth: opts.childExtentWidth,
      childExtentHeight: opts.childExtentHeight,
      fill: opts.fill,
      effects: opts.effects,
      groupShapeLocks: opts.groupShapeLocks,
      extensions: opts.extensions,
      blackWhiteMode: opts.blackWhiteMode,
      type: "wpg",
    };

    // Register pic children media so {fileName} placeholders resolve, recursing
    // into nested wpg groups. wps children carry shape data, not media.
    registerDrawingChildMedia(mediaData.children, ctx);

    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
        graphicFrameLocks: opts.graphicFrameLocks,
        extensionIds: opts.extensionIds,
        dialect: drawingDialect(ctx),
      },
      ctx,
    );
    const vmlFallback = registerVmlFallbackMedia(opts, ctx);
    return wrap(drawingXml, opts, vmlFallback);
  }

  // WordprocessingCanvas — shares WPG child registration and descriptor output.
  if ("wordprocessingCanvas" in child) {
    const opts = child.wordprocessingCanvas;
    const mediaData: CanvasMediaData = {
      type: "wpc",
      children: opts.children.map(cloneDrawingChild),
      transformation: createTransformation(opts.transformation),
      background: opts.background,
      whole: opts.whole,
      extensions: opts.extensions,
    };
    registerDrawingChildMedia(mediaData.children, ctx);
    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
        graphicFrameLocks: opts.graphicFrameLocks,
        extensionIds: opts.extensionIds,
        dialect: drawingDialect(ctx),
      },
      ctx,
    );
    const vmlFallback = registerVmlFallbackMedia(opts, ctx);
    return wrap(drawingXml, opts, vmlFallback);
  }

  return undefined;
}
