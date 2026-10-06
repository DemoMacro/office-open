/**
 * Worksheet drawing compile phase: images, charts, SmartArt, shapes,
 * connectors, and groups on a sheet — the drawing part, its relationships,
 * and the `<drawing>` reference inserted into the sheet XML.
 *
 * @module
 */

import {
  RELATIONSHIP_TYPES,
  Relationships,
  type PassthroughRelationship,
  resolveRelationshipTarget,
  type RelationshipType,
  pickNonVisualDrawingProperties,
  toUint8Array,
} from "@office-open/core";
import { buildUserShapesData, chartSpaceDesc } from "@office-open/core/chart";
import { OOXML_XML_DECLARATION } from "@office-open/xml";
import type {
  DrawingChartOptions,
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

function decimalAttr(value: string): string {
  const number = Number(value);
  if (!Number.isFinite(number)) return value;
  const shortest = String(number);
  if (shortest.length < 17) return shortest;
  return number.toPrecision(17).replace(/0+$/, "").replace(/\.$/, "");
}

function layoutDecimal(value: string): string {
  const number = Number(value);
  if (!Number.isFinite(number)) return value;
  if (number !== 0 && Math.abs(number) < 0.1) {
    return number.toExponential(16).toUpperCase();
  }
  return decimalAttr(value);
}

export function preserveChartDecimalAttributes(xml: string): string {
  return xml
    .replace(/<c:pageMargins([^>]*)\/>/g, (_, attributes: string) => {
      const normalized = attributes.replace(
        /\b(l|r|t|b|header|footer)="([^"]+)"/g,
        (_match, name: string, value: string) => `${name}="${decimalAttr(value)}"`,
      );
      return `<c:pageMargins${normalized}/>`;
    })
    .replace(
      /<c:(x|y|w|h) val="([^"]+)"\/>/g,
      (_match, name: string, value: string) => `<c:${name} val="${layoutDecimal(value)}"/>`,
    );
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
  const smartArtOpts = wsOpts.smartArts ?? [];
  const webExtensionOpts = wsOpts.webExtensions ?? [];
  const shapeOpts = wsOpts.shapes ?? [];
  const connectorOpts = wsOpts.connectors ?? [];
  const groupOpts = wsOpts.groups ?? [];
  const contentPartOpts = wsOpts.contentParts ?? [];

  const drawingImages: DrawingPictureOptions[] = [];
  const drawingCharts: DrawingChartOptions[] = [];
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

    drawingImages.push({
      ...pickAnchorOptions(img),
      rId: embedRid ?? "",
      ...(linkRid ? { linkRId: linkRid } : {}),
      ...pickNonVisualDrawingProperties(img),
      ...(img.properties ? { properties: img.properties } : {}),
      ...(img.blackWhiteMode ? { blackWhiteMode: img.blackWhiteMode } : {}),
      ...(img.compression !== undefined ? { compression: img.compression } : {}),
      ...(img.sourceRectangle ? { sourceRectangle: img.sourceRectangle } : {}),
      ...(img.preferRelativeResize !== undefined
        ? { preferRelativeResize: img.preferRelativeResize }
        : {}),
      ...(img.blipEffects ? { blipEffects: img.blipEffects } : {}),
      ...(img.useLocalDpi !== undefined ? { useLocalDpi: img.useLocalDpi } : {}),
      ...(img.blipExt !== undefined ? { blipExt: img.blipExt } : {}),
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
    const userShapes = chart.userShapes ? buildUserShapesData(chart.userShapes) : undefined;
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
      rId: chartRid,
      ...(chart.frameLocks ? { frameLocks: chart.frameLocks } : {}),
      ...(chart.macro !== undefined ? { macro: chart.macro } : {}),
      ...(chart.hyperlink ? { hyperlink: chart.hyperlink } : {}),
      ...(chart.zOrder !== undefined ? { zOrder: chart.zOrder } : {}),
      ...(chart.shapeId !== undefined ? { shapeId: chart.shapeId } : {}),
    });
    state.globalChartIdx++;
  }

  // Process SmartArt — the diagram parts themselves are passthrough, so the
  // relationship targets stay exactly the source-relative form; only the
  // rIds are renumbered into the rebuilt drawing rels.
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
    const relBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    drawingSmartArts.push({
      ...pickAnchorOptions(sa),
      ...pickNonVisualDrawingProperties(sa),
      dataRId: addDiagRel(`${relBase}/diagramData`, sa.dataPath),
      layoutRId: addDiagRel(`${relBase}/diagramLayout`, sa.layoutPath),
      quickStyleRId: addDiagRel(`${relBase}/diagramQuickStyle`, sa.quickStylePath),
      colorsRId: addDiagRel(`${relBase}/diagramColors`, sa.colorsPath),
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
      ...(contentPart.alternateContent ? { alternateContent: true } : {}),
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
      ...(fallback ? { fallback } : {}),
      ...(webExtension.frameLocks ? { frameLocks: webExtension.frameLocks } : {}),
      ...(webExtension.macro !== undefined ? { macro: webExtension.macro } : {}),
      ...(webExtension.zOrder !== undefined ? { zOrder: webExtension.zOrder } : {}),
      ...(webExtension.shapeId !== undefined ? { shapeId: webExtension.shapeId } : {}),
    });
  }

  // Generate drawing XML (via descriptor). Snapshot the hyperlink registry
  // first so only runs stringified for this sheet's drawing resolve here.
  const hyperlinkBase = ctx.hyperlinks.length;
  const drawingXml = drawingDesc.stringify(
    {
      images: drawingImages,
      charts: drawingCharts,
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
          "External",
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
  const drawingIdx = i + 1;
  mapping[`Drawing${i}`] = {
    data: XML_DECL + resolvedDrawingXml,
    path: drawingPath,
  };

  // Drawing relationships
  mapping[`DrawingRels${i}`] = {
    data: XML_DECL + drawingRels.serialize(),
    path: drawingPath.replace(/([^/]+)$/, "_rels/$1.rels"),
  };

  // Insert drawing reference at its CT_Worksheet sequence position.
  const drawingTarget = drawingRel?.target ?? `../drawings/drawing${drawingIdx}.xml`;
  const existingDrawingRid = wsRels.idOf(RELATIONSHIP_TYPES.drawing, drawingTarget);
  const drawingRid =
    existingDrawingRid ?? `rId${wsRels.add(RELATIONSHIP_TYPES.drawing, drawingTarget)}`;
  return editSheetTailMarker(sheetXml, "<!--DRAWING-->", `<drawing r:id="${drawingRid}"/>`);
}
