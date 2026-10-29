/**
 * Worksheet drawing compile phase: images, charts, SmartArt, shapes,
 * connectors, and groups on a sheet — the drawing part, its relationships,
 * and the `<drawing>` reference inserted into the sheet XML.
 *
 * @module
 */

import {
  imageTypeFromPath,
  pickNonVisualDrawingProperties,
  Relationships,
  toUint8Array,
} from "@office-open/core";
import { buildUserShapesData, chartExSpaceDesc, chartSpaceDesc } from "@office-open/core/chart";
import {
  partPathToRelsPath,
  RELATIONSHIP_TYPES,
  resolveRelationshipTarget,
  type PassthroughRelationship,
  type RelationshipType,
} from "@office-open/core/opc";
import {
  createDataModel,
  definitionId,
  getColorXml,
  getLayoutXml,
  getStyleXml,
  stringifyColorDefinitionPart,
  stringifyDataModelOptions,
  stringifyLayoutDefinitionPart,
  stringifyStyleDefinitionPart,
} from "@office-open/core/smartart";
import { OOXML_XML_DECLARATION } from "@office-open/xml";
import type {
  DrawingChartOptions,
  DrawingChartExOptions,
  DrawingContentPartOptions,
  DrawingPictureOptions,
  DrawingSmartArtOptions,
  DrawingWebExtensionOptions,
} from "@parts/drawing";
import { pickAnchorOptions, drawingDesc } from "@parts/drawing";
import { editSheetTailMarker, type WorksheetOptions } from "@parts/worksheet";

import type { WorksheetCompileState } from "../compiler";
import { XlsxWriteContext } from "../context";

const XML_DECL = OOXML_XML_DECLARATION;

const IMAGE_REL = RELATIONSHIP_TYPES.image;

export function preserveChartDecimalAttributes(xml: string): string {
  return xml;
}

/**
 * Replace `{fileName}` media placeholders in compiled part XML with
 * relationship ids. Core fill descriptors register blip-fill images through
 * `ctx.addMedia`, which returns a placeholder because the owning part's rels
 * don't exist yet; a placeholder surviving into the part makes Excel refuse
 * the package. Each distinct image registers one relationship — consuming
 * parts (theme, drawings) live one level under `xl/`, so the media target is
 * always `../media/<name>`.
 */
export function bindMediaPlaceholders(
  xml: string,
  media: XlsxWriteContext["media"],
  rels: Relationships,
): string {
  const names = new Set(media.array.map((m) => m.fileName));
  let usesMedia = false;
  for (const name of names) {
    if (xml.includes(`{${name}}`)) {
      usesMedia = true;
      break;
    }
  }
  if (!usesMedia) return xml;
  const ridByName = new Map<string, string>();
  return xml.replace(/\{([^{}]+)\}/g, (whole, name: string) => {
    let rid = ridByName.get(name);
    if (rid === undefined) {
      if (!names.has(name)) return whole;
      rid = `rId${rels.add(IMAGE_REL, `../media/${name}`)}`;
      ridByName.set(name, rid);
    }
    return rid;
  });
}

/** Compile a sheet's drawing: register image/chart/SmartArt media and
 * relationships, serialize the drawing part via the descriptor, bind media
 * and text-hyperlink placeholders, and splice the `<drawing r:id>` element
 * into the sheet XML at its CT_Worksheet position. Returns the updated XML. */
export function compileSheetDrawing(
  wsOpts: WorksheetOptions,
  i: number,
  sheetXml: string,
  ctx: XlsxWriteContext,
  mapping: Record<string, { data: string; path: string }>,
  state: WorksheetCompileState,
  wsRels: Relationships,
  passthroughRelationships: readonly PassthroughRelationship[] | undefined,
  wsPath: string,
): string {
  const imgOpts = wsOpts.images ?? [];
  const chartOpts = wsOpts.charts ?? [];
  const chartExOpts = wsOpts.chartExs ?? [];
  const smartArtOpts = wsOpts.smartArts ?? [];
  const webExtensionOpts = wsOpts.webExtensions ?? [];
  const shapeOpts = wsOpts.shapes ?? [];
  const connectorOpts = wsOpts.connectors ?? [];
  const groupOpts = wsOpts.groups ?? [];
  const contentPartOpts = wsOpts.contentParts ?? [];

  const drawingImages: DrawingPictureOptions[] = [];
  const drawingCharts: DrawingChartOptions[] = [];
  const drawingChartExs: DrawingChartExOptions[] = [];
  const drawingSmartArts: DrawingSmartArtOptions[] = [];
  const drawingWebExtensions: DrawingWebExtensionOptions[] = [];
  const drawingContentParts: DrawingContentPartOptions[] = [];
  const drawingRels = new Relationships();
  let rid = 1;
  const sourceWorksheetRels = (passthroughRelationships ?? []).filter(
    (rel) => rel.source === wsPath,
  );
  const drawingRel = sourceWorksheetRels.find((rel) => rel.relationshipType.endsWith("/drawing"));
  const drawingPath = drawingRel
    ? resolveRelationshipTarget(wsPath, drawingRel.target)
    : `xl/drawings/drawing${i + 1}.xml`;
  const sourceDrawingRels = (passthroughRelationships ?? []).filter(
    (rel) => rel.source === drawingPath,
  );
  const addPreservedDrawingRel = (
    sourceRel: { rId: string } | undefined,
    type: RelationshipType,
    target: string,
    targetMode?: "External",
  ): string => {
    const existing = drawingRels.idOf(type, target);
    if (existing !== undefined) return existing;
    const preferred = /^rId\d+$/.exec(sourceRel?.rId ?? "")?.[0];
    if (preferred && !drawingRels.hasId(preferred)) {
      drawingRels.addRelationship(preferred, type, target, targetMode);
      rid = drawingRels.nextRelationshipId;
      return preferred;
    }
    drawingRels.addRelationship(rid, type, target, targetMode);
    const assigned = `rId${rid}`;
    rid = drawingRels.nextRelationshipId;
    return assigned;
  };
  const remapBlipExtensionRelationships = (extensions: string): string => {
    const replacements = [...extensions.matchAll(/r:embed="([^"]+)"/g)]
      .map((match) => match[1]!)
      .filter((sourceRid, index, ids) => ids.indexOf(sourceRid) === index)
      .map((sourceRid) => {
        const sourceRel = sourceDrawingRels.find((rel) => rel.rId === sourceRid);
        if (!sourceRel) return undefined;
        const targetRid = addPreservedDrawingRel(
          sourceRel,
          sourceRel.relationshipType as RelationshipType,
          sourceRel.target,
          sourceRel.targetMode === "External" ? "External" : undefined,
        );
        return targetRid === sourceRid ? undefined : { sourceRid, targetRid };
      })
      .filter((replacement) => replacement !== undefined);
    return replacements.reduce(
      (xml, replacement) =>
        xml.replaceAll(
          `r:embed="${replacement!.sourceRid}"`,
          `r:embed="${replacement!.targetRid}"`,
        ),
      extensions,
    );
  };

  // Process images
  for (const img of imgOpts) {
    let embedRid: string | undefined;
    let linkRid: string | undefined;

    if (img.data !== undefined) {
      // Media-store extension (jpg → jpeg); vector formats pass through.
      const ext = img.type === "jpg" ? "jpeg" : img.type;
      const rawBytes = toUint8Array(img.data, { encoding: "base64" });
      const sourceImageRel =
        (img.sourcePath
          ? sourceDrawingRels.find(
              (rel) =>
                rel.relationshipType.endsWith("/image") &&
                resolveRelationshipTarget(drawingPath, rel.target) === img.sourcePath,
            )
          : undefined) ??
        sourceDrawingRels.filter((rel) => rel.relationshipType.endsWith("/image"))[
          drawingImages.length
        ];
      const sourceImagePath = sourceImageRel
        ? resolveRelationshipTarget(drawingPath, sourceImageRel.target)
        : undefined;
      const entry = ctx.media.addMedia(
        rawBytes,
        ext,
        (fileName) => ({
          fileName,
          type: ext,
          data: rawBytes,
          width: 0,
          height: 0,
        }),
        (sourceImagePath ?? img.sourcePath)?.split("/").pop(),
      );

      // Anchors sharing one picture share the relationship too — the source
      // writes a single image rel that every a:blip references.
      const target = sourceImageRel?.target ?? `../media/${entry.fileName}`;
      embedRid = addPreservedDrawingRel(sourceImageRel, IMAGE_REL, target);
      state.globalMediaIdx++;
    }

    // Linked source (a:blip @r:link): one External image relationship per
    // URL — no media part, no bytes.
    if (img.sourceUrl !== undefined) {
      drawingRels.addRelationship(rid, IMAGE_REL, img.sourceUrl, "External");
      linkRid = `rId${rid}`;
      rid++;
    }
    const blipExt =
      img.blipExt === undefined ? undefined : remapBlipExtensionRelationships(img.blipExt);

    drawingImages.push({
      ...pickAnchorOptions(img),
      rId: embedRid ?? "",
      ...(linkRid ? { linkRId: linkRid } : {}),
      ...pickNonVisualDrawingProperties(img),
      ...(img.properties ? { properties: img.properties } : {}),
      ...(img.blackWhiteMode ? { blackWhiteMode: img.blackWhiteMode } : {}),
      ...(img.macro !== undefined ? { macro: img.macro } : {}),
      ...(img.compression !== undefined ? { compression: img.compression } : {}),
      ...(img.sourceRectangle ? { sourceRectangle: img.sourceRectangle } : {}),
      ...(img.fillRectangle !== undefined ? { fillRectangle: img.fillRectangle } : {}),
      ...(img.blipFillPresent !== undefined ? { blipFillPresent: img.blipFillPresent } : {}),
      ...(img.dpi !== undefined ? { dpi: img.dpi } : {}),
      ...(img.rotWithShape !== undefined ? { rotWithShape: img.rotWithShape } : {}),
      ...(img.preferRelativeResize !== undefined
        ? { preferRelativeResize: img.preferRelativeResize }
        : {}),
      ...(img.blipEffects ? { blipEffects: img.blipEffects } : {}),
      ...(img.useLocalDpi !== undefined ? { useLocalDpi: img.useLocalDpi } : {}),
      ...(blipExt !== undefined ? { blipExt } : {}),
      ...(img.cNvPicPrExt !== undefined ? { cNvPicPrExt: img.cNvPicPrExt } : {}),
      ...(img.locking ? { locking: img.locking } : {}),
      ...(img.hyperlink ? { hyperlink: img.hyperlink } : {}),
      ...(img.zOrder !== undefined ? { zOrder: img.zOrder } : {}),
      ...(img.shapeId !== undefined ? { shapeId: img.shapeId } : {}),
    });
  }

  // Process charts
  for (const chart of chartOpts) {
    const chartKey = `chart_${state.globalChartIdx}`;
    const sourceChartRel = chart.sourcePath
      ? sourceDrawingRels.find(
          (rel) =>
            rel.relationshipType.endsWith("/chart") &&
            resolveRelationshipTarget(drawingPath, rel.target) === chart.sourcePath,
        )
      : undefined;
    const chartPath = sourceChartRel
      ? resolveRelationshipTarget(drawingPath, sourceChartRel.target)
      : `xl/charts/chart${state.globalChartIdx + 1}.xml`;
    state.chartPaths.set(chartKey, chartPath);
    ctx.chartOptions.set(chartKey, chart);
    const userShapes = chart.userShapes
      ? {
          ...buildUserShapesData(chart.userShapes),
          ...(chart.userShapes.path ? { path: chart.userShapes.path } : {}),
        }
      : undefined;
    ctx.charts.addChart(chartKey, {
      key: chartKey,
      chartSpaceXml: preserveChartDecimalAttributes(chartSpaceDesc.stringify(chart, ctx) ?? ""),
      ...(userShapes ? { userShapes } : {}),
    });

    const chartTarget = sourceChartRel?.target ?? `../charts/chart${state.globalChartIdx + 1}.xml`;
    const chartRid = addPreservedDrawingRel(sourceChartRel, RELATIONSHIP_TYPES.chart, chartTarget);

    // cNvPr @title/@ext stay unbridged: WorksheetChartOptions.title is the
    // chart title (c:title) and its ext is the chart-space c:extLst —
    // neither belongs on the graphicFrame's cNvPr.
    const chartCnvPr = pickNonVisualDrawingProperties({
      ...chart,
      title: undefined,
      ext: undefined,
    });
    drawingCharts.push({
      ...pickAnchorOptions(chart),
      ...chartCnvPr,
      frameTransform: chart.frameTransform !== false,
      ...(chart.frameExtentCx !== undefined ? { frameExtentCx: chart.frameExtentCx } : {}),
      ...(chart.frameExtentCy !== undefined ? { frameExtentCy: chart.frameExtentCy } : {}),
      ...(chart.frameOffsetX !== undefined ? { frameOffsetX: chart.frameOffsetX } : {}),
      ...(chart.frameOffsetY !== undefined ? { frameOffsetY: chart.frameOffsetY } : {}),
      ...(chart.frameRotation !== undefined ? { frameRotation: chart.frameRotation } : {}),
      ...(chart.frameFlipHorizontal !== undefined
        ? { frameFlipHorizontal: chart.frameFlipHorizontal }
        : {}),
      ...(chart.frameFlipVertical !== undefined
        ? { frameFlipVertical: chart.frameFlipVertical }
        : {}),
      ...(chart.frameTitle !== undefined ? { title: chart.frameTitle } : {}),
      ...(chart.frameHidden !== undefined ? { hidden: chart.frameHidden } : {}),
      rId: chartRid,
      ...(chart.frameLocks ? { frameLocks: chart.frameLocks } : {}),
      ...(chart.macro !== undefined ? { macro: chart.macro } : {}),
      ...(chart.fPublished !== undefined ? { fPublished: chart.fPublished } : {}),
      ...(chart.hyperlink ? { hyperlink: chart.hyperlink } : {}),
      ...(chart.zOrder !== undefined ? { zOrder: chart.zOrder } : {}),
      ...(chart.shapeId !== undefined ? { shapeId: chart.shapeId } : {}),
    });
    state.globalChartIdx++;
  }

  // ChartEx parts are typed canonical parts; the frame only carries their
  // package relationship, while companion rels keep source rIds verbatim.
  for (const chartEx of chartExOpts) {
    const sourceChartExRel = chartEx.sourcePath
      ? sourceDrawingRels.find(
          (rel) =>
            rel.relationshipType.endsWith("/chartEx") &&
            resolveRelationshipTarget(drawingPath, rel.target) === chartEx.sourcePath,
        )
      : sourceDrawingRels.find((rel) => rel.relationshipType.endsWith("/chartEx"));
    const chartExIndex = state.globalChartExIdx;
    const chartExPath =
      sourceChartExRel !== undefined && chartEx.sourcePath !== undefined
        ? chartEx.sourcePath
        : `xl/charts/chartEx${chartExIndex + 1}.xml`;
    const chartExTarget = sourceChartExRel?.target ?? `../charts/chartEx${chartExIndex + 1}.xml`;
    const chartExRid = addPreservedDrawingRel(
      sourceChartExRel,
      RELATIONSHIP_TYPES.chartEx,
      chartExTarget,
    );

    const chartExRels = new Relationships(chartExPath);
    for (const sourceRel of chartEx.sourceRelationships ?? []) {
      chartExRels.claimSourceRel(sourceRel);
    }
    mapping[`ChartEx${chartExIndex}`] = {
      data: XML_DECL + (chartExSpaceDesc.stringify(chartEx, ctx) ?? ""),
      path: chartExPath,
    };
    if (chartExRels.relationshipCount > 0) {
      mapping[`ChartExRels${chartExIndex}`] = {
        data: XML_DECL + chartExRels.serialize(),
        path: partPathToRelsPath(chartExPath),
      };
    }

    drawingChartExs.push({
      ...pickAnchorOptions(chartEx),
      ...pickNonVisualDrawingProperties(chartEx),
      rId: chartExRid,
      ...(chartEx.frameExtentCx !== undefined ? { frameExtentCx: chartEx.frameExtentCx } : {}),
      ...(chartEx.frameExtentCy !== undefined ? { frameExtentCy: chartEx.frameExtentCy } : {}),
      ...(chartEx.frameOffsetX !== undefined ? { frameOffsetX: chartEx.frameOffsetX } : {}),
      ...(chartEx.frameOffsetY !== undefined ? { frameOffsetY: chartEx.frameOffsetY } : {}),
      ...(chartEx.frameRotation !== undefined ? { frameRotation: chartEx.frameRotation } : {}),
      ...(chartEx.frameFlipHorizontal !== undefined
        ? { frameFlipHorizontal: chartEx.frameFlipHorizontal }
        : {}),
      ...(chartEx.frameFlipVertical !== undefined
        ? { frameFlipVertical: chartEx.frameFlipVertical }
        : {}),
      ...(chartEx.frameLocks ? { frameLocks: chartEx.frameLocks } : {}),
      ...(chartEx.macro !== undefined ? { macro: chartEx.macro } : {}),
      ...(chartEx.fPublished !== undefined ? { fPublished: chartEx.fPublished } : {}),
      ...(chartEx.hyperlink ? { hyperlink: chartEx.hyperlink } : {}),
      ...(chartEx.zOrder !== undefined ? { zOrder: chartEx.zOrder } : {}),
      ...(chartEx.shapeId !== undefined ? { shapeId: chartEx.shapeId } : {}),
    });
    state.globalChartExIdx++;
  }

  // Process SmartArt. The modeled XML parts are always rebuilt from canonical
  // data; source paths remain round-trip metadata and never select raw output.
  for (const sa of smartArtOpts) {
    const relTarget = (path: string): string =>
      path.startsWith("xl/") ? `../${path.slice(3)}` : path;
    // One relationship per (kind, target) — anchors sharing a diagram set
    // share its relationships too.
    const addDiagRel = (relType: RelationshipType, path: string): string => {
      const target = relTarget(path);
      const existing = drawingRels.idOf(relType, target);
      if (existing !== undefined) return existing;
      return addPreservedDrawingRel(undefined, relType, target);
    };
    const smartArtIndex = state.globalSmartArtIdx;
    const canonicalPaths = {
      data: `xl/diagrams/data${smartArtIndex}.xml`,
      layout: `xl/diagrams/layout${smartArtIndex}.xml`,
      quickStyle: `xl/diagrams/quickStyle${smartArtIndex}.xml`,
      colors: `xl/diagrams/colors${smartArtIndex}.xml`,
    };
    if (sa.nodes) {
      state.globalSmartArtIdx++;
      const layoutId =
        typeof sa.layout === "string" || sa.layout === undefined
          ? (sa.layout ?? "default")
          : definitionId(sa.layout);
      const styleId =
        typeof sa.style === "string" || sa.style === undefined
          ? (sa.style ?? "simple1")
          : definitionId(sa.style);
      const colorId =
        typeof sa.color === "string" || sa.color === undefined
          ? (sa.color ?? "accent1_2")
          : definitionId(sa.color);
      const dataRelationships = new Relationships(canonicalPaths.data);
      const dataXml = sa.dataModel
        ? bindMediaPlaceholders(
            stringifyDataModelOptions(sa.dataModel, ctx),
            ctx.media,
            dataRelationships,
          ).replace(
            "<dgm:dataModel>",
            '<dgm:dataModel xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">',
          )
        : createDataModel(sa.nodes, layoutId, styleId, colorId, ctx.reproducible);
      mapping[`SmartArtData${smartArtIndex}`] = {
        data: XML_DECL + dataXml,
        path: canonicalPaths.data,
      };
      if (dataRelationships.relationshipCount > 0) {
        mapping[`SmartArtDataRels${smartArtIndex}`] = {
          data: XML_DECL + dataRelationships.serialize(),
          path: partPathToRelsPath(canonicalPaths.data),
        };
      }
      let layoutXml =
        typeof sa.layout === "string" || sa.layout === undefined
          ? getLayoutXml(layoutId)
          : stringifyLayoutDefinitionPart(sa.layout);
      if (typeof sa.layout === "object" && sa.raw?.layoutMedia?.length) {
        const layoutRels = new Relationships(canonicalPaths.layout);
        for (const media of sa.raw.layoutMedia) {
          const data = toUint8Array(media.data);
          const type = imageTypeFromPath(media.fileName);
          const entry = ctx.media.addMedia(
            data,
            type,
            (fileName) => ({
              type,
              data,
              fileName,
              width: 0,
              height: 0,
            }),
            media.fileName,
          );
          const relationshipId = layoutRels.add(IMAGE_REL, `../media/${entry.fileName}`);
          const reference = media.partPath ?? `xl/media/${media.fileName}`;
          layoutXml = layoutXml.replaceAll(
            `{smartart-layout-image:${reference}}`,
            `rId${relationshipId}`,
          );
        }
        if (layoutRels.relationshipCount > 0) {
          mapping[`SmartArtLayoutRels${smartArtIndex}`] = {
            data: XML_DECL + layoutRels.serialize(),
            path: partPathToRelsPath(canonicalPaths.layout),
          };
        }
      }
      mapping[`SmartArtLayout${smartArtIndex}`] = {
        data: layoutXml,
        path: canonicalPaths.layout,
      };
      mapping[`SmartArtQuickStyle${smartArtIndex}`] = {
        data:
          typeof sa.style === "string" || sa.style === undefined
            ? getStyleXml(styleId)
            : stringifyStyleDefinitionPart(sa.style),
        path: canonicalPaths.quickStyle,
      };
      mapping[`SmartArtColors${smartArtIndex}`] = {
        data:
          typeof sa.color === "string" || sa.color === undefined
            ? getColorXml(colorId)
            : stringifyColorDefinitionPart(sa.color),
        path: canonicalPaths.colors,
      };
      state.generatedSmartArtParts.push(
        {
          path: canonicalPaths.data,
          contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
        },
        {
          path: canonicalPaths.layout,
          contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
        },
        {
          path: canonicalPaths.quickStyle,
          contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
        },
        {
          path: canonicalPaths.colors,
          contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
        },
      );
    }
    const usePaths = canonicalPaths;
    if (!usePaths.data || !usePaths.layout || !usePaths.quickStyle || !usePaths.colors) {
      throw new Error("SmartArt requires canonical content or the four source paths.");
    }
    const relBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    drawingSmartArts.push({
      ...pickAnchorOptions(sa),
      ...pickNonVisualDrawingProperties(sa),
      dataRId: addDiagRel(`${relBase}/diagramData`, usePaths.data),
      layoutRId: addDiagRel(`${relBase}/diagramLayout`, usePaths.layout),
      quickStyleRId: addDiagRel(`${relBase}/diagramQuickStyle`, usePaths.quickStyle),
      colorsRId: addDiagRel(`${relBase}/diagramColors`, usePaths.colors),
      ...(sa.frameLocks ? { frameLocks: sa.frameLocks } : {}),
      ...(sa.macro !== undefined ? { macro: sa.macro } : {}),
      ...(sa.zOrder !== undefined ? { zOrder: sa.zOrder } : {}),
      ...(sa.shapeId !== undefined ? { shapeId: sa.shapeId } : {}),
    });
  }

  // Content parts are opaque companions; only their drawing relationship is
  // rebuilt. The companion bytes stay in rawParts at the source path.
  for (const contentPart of contentPartOpts) {
    const sourceContentRel = sourceDrawingRels.find(
      (rel) =>
        rel.rId === contentPart.rId ||
        resolveRelationshipTarget(drawingPath, rel.target) === contentPart.sourcePath,
    );
    const rId = addPreservedDrawingRel(
      sourceContentRel,
      contentPart.relationshipType as RelationshipType,
      contentPart.relationshipTarget,
    );
    drawingContentParts.push({
      ...pickAnchorOptions(contentPart),
      rId,
      ...(contentPart.zOrder !== undefined ? { zOrder: contentPart.zOrder } : {}),
      ...(contentPart.shapeId !== undefined ? { shapeId: contentPart.shapeId } : {}),
      ...(contentPart.alternateContent !== undefined
        ? { alternateContent: contentPart.alternateContent }
        : {}),
    });
  }

  // WebExtension parts stay passthrough; the rebuilt drawing only needs their
  // source relationships and the fallback snapshot image relationship.
  for (const webExtension of webExtensionOpts) {
    const sourceWebExtensionRel = webExtension.sourcePath
      ? sourceDrawingRels.find(
          (rel) =>
            rel.relationshipType.endsWith("/webextension") &&
            resolveRelationshipTarget(drawingPath, rel.target) === webExtension.sourcePath,
        )
      : undefined;
    const webExtensionTarget =
      sourceWebExtensionRel?.target ?? `../${webExtension.sourcePath.replace(/^xl\//, "")}`;
    const rId = addPreservedDrawingRel(
      sourceWebExtensionRel,
      RELATIONSHIP_TYPES.webExtensionMs,
      webExtensionTarget,
    );
    let fallback = webExtension.fallback;
    if (fallback?.rId && webExtension.snapshotSourcePath) {
      const sourceImageRel = sourceDrawingRels.find(
        (rel) =>
          rel.relationshipType.endsWith("/image") &&
          resolveRelationshipTarget(drawingPath, rel.target) === webExtension.snapshotSourcePath,
      );
      const imageTarget =
        sourceImageRel?.target ?? `../${webExtension.snapshotSourcePath.replace(/^xl\//, "")}`;
      fallback = {
        ...fallback,
        rId: addPreservedDrawingRel(sourceImageRel, RELATIONSHIP_TYPES.image, imageTarget),
      };
    }
    drawingWebExtensions.push({
      ...pickAnchorOptions(webExtension),
      ...pickNonVisualDrawingProperties(webExtension),
      rId,
      ...(webExtension.elementName ? { elementName: webExtension.elementName } : {}),
      ...(fallback ? { fallback } : {}),
      ...(webExtension.frameLocks ? { frameLocks: webExtension.frameLocks } : {}),
      ...(webExtension.macro !== undefined ? { macro: webExtension.macro } : {}),
      ...(webExtension.zOrder !== undefined ? { zOrder: webExtension.zOrder } : {}),
      ...(webExtension.shapeId !== undefined ? { shapeId: webExtension.shapeId } : {}),
    });
  }

  // Generate drawing XML (via descriptor). Snapshot the hyperlink map
  // first so only runs stringified for this sheet's drawing resolve here.
  const hyperlinkBase = ctx.hyperlinks.length;
  const drawingXml = drawingDesc.stringify(
    {
      images: drawingImages,
      charts: drawingCharts,
      chartExs: drawingChartExs,
      smartArts: drawingSmartArts,
      webExtensions: drawingWebExtensions,
      shapes: shapeOpts,
      connectors: connectorOpts,
      groups: groupOpts,
      contentParts: drawingContentParts,
    },
    ctx,
  );
  // Resolve drawing shape text-hyperlink placeholders ({hlink:key} → real
  // rId) and register each as an External hyperlink relationship.
  let resolvedDrawingXml = drawingXml!;
  // One External relationship per distinct URL — several objects/runs
  // pointing at the same target share it (matches how Excel writes rels).
  const hlinkRidByUrl = new Map<string, number>();
  for (const h of ctx.hyperlinks.slice(hyperlinkBase)) {
    let hlinkRid = hlinkRidByUrl.get(h.url);
    if (hlinkRid === undefined) {
      const sourceHlinkRel = sourceDrawingRels.find(
        (rel) => rel.relationshipType.endsWith("/hyperlink") && rel.target === h.url,
      );
      hlinkRid = Number(
        addPreservedDrawingRel(
          sourceHlinkRel,
          RELATIONSHIP_TYPES.hyperlink,
          h.url,
          sourceHlinkRel
            ? sourceHlinkRel.targetMode === "External"
              ? "External"
              : undefined
            : "External",
        ).replace(/^rId/, ""),
      );
      hlinkRidByUrl.set(h.url, hlinkRid);
    }
    resolvedDrawingXml = resolvedDrawingXml
      .split(`r:id="{hlink:${h.key}}"`)
      .join(`r:id="rId${hlinkRid}"`);
  }
  // Shape blip fills inside the drawing register `{fileName}` media
  // placeholders — bind them the same way the theme does.
  resolvedDrawingXml = bindMediaPlaceholders(resolvedDrawingXml, ctx.media, drawingRels);
  // Typed anchors absorb their known relationships; unmodeled companions
  // (for example diagram preview drawings) still belong to the drawing part.
  for (const sourceRel of sourceDrawingRels) {
    if (drawingRels.hasRelationship(sourceRel.relationshipType, sourceRel.target)) continue;
    const preferred = /^rId\d+$/.exec(sourceRel.rId)?.[0];
    if (preferred && !drawingRels.hasId(preferred)) {
      drawingRels.addRelationship(
        preferred,
        sourceRel.relationshipType as RelationshipType,
        sourceRel.target,
        sourceRel.targetMode,
      );
      rid = drawingRels.nextRelationshipId;
      continue;
    }
    drawingRels.addRelationship(
      rid,
      sourceRel.relationshipType as RelationshipType,
      sourceRel.target,
      sourceRel.targetMode,
    );
    rid = drawingRels.nextRelationshipId;
  }
  const sourceDrawingIndex = Number(drawingPath.match(/drawing(\d+)\.xml$/)?.[1]);
  const drawingIdx = sourceDrawingIndex ?? state.nextDrawingIndex++;
  if (sourceDrawingIndex !== undefined) {
    state.nextDrawingIndex = Math.max(state.nextDrawingIndex, sourceDrawingIndex + 1);
  }
  mapping[`Drawing${i}`] = {
    data: XML_DECL + resolvedDrawingXml,
    path: drawingPath,
  };

  // Drawing relationships — omitted when empty (Office strips empty rels shells).
  if (drawingRels.relationshipCount > 0) {
    mapping[`DrawingRels${i}`] = {
      data: XML_DECL + drawingRels.serialize(),
      path: drawingPath.replace(/([^/]+)$/, "_rels/$1.rels"),
    };
  }

  // Insert drawing reference at its CT_Worksheet sequence position.
  if (wsOpts.drawingReferencePresent === false) return sheetXml;
  const drawingTarget = drawingRel?.target ?? `../drawings/drawing${drawingIdx}.xml`;
  const existingDrawingRid = wsRels.idOf(RELATIONSHIP_TYPES.drawing, drawingTarget);
  const drawingRid =
    existingDrawingRid ?? `rId${wsRels.add(RELATIONSHIP_TYPES.drawing, drawingTarget)}`;
  return editSheetTailMarker(sheetXml, "<!--DRAWING-->", `<drawing r:id="${drawingRid}"/>`);
}
