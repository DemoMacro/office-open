/**
 * Drawing tail compile phase: the chart parts (chartSpace, userShapes, their
 * rels and embedded workbooks) and the SmartArt diagram part sets
 * (data/layout/quickStyle/colors plus the Office render-cache drawing), read
 * off the collections the document body stringification filled.
 *
 * @module
 */

import {
  Relationships,
  replacePlaceholders,
  toUint8Array,
  withDiagramNamespaces,
} from "@office-open/core";
import { partPathToRelsPath, RELATIONSHIP_TYPES, type XmlifyedFile } from "@office-open/core/opc";
import {
  getColorXml,
  getLayoutXml,
  getStyleXml,
  stringifyColorDefinitionPart,
  stringifyLayoutDefinitionPart,
  stringifyStyleDefinitionPart,
} from "@office-open/core/smartart";

import type { DocxWriteContext } from "../context";
import { XML_DECL, embeddingRelationshipType } from "./shared";

/** Chart part → user-shapes part relationship (c:userShapes bridge). */
const CHART_USER_SHAPES_REL = RELATIONSHIP_TYPES.chartUserShapes;

/** Relative relationship target from one package part to another. */
function relativeTarget(ownerPath: string, targetPath: string): string {
  const owner = ownerPath.split("/").slice(0, -1);
  const target = targetPath.split("/");
  let common = 0;
  while (common < owner.length && common < target.length - 1 && owner[common] === target[common])
    common++;
  return [
    ...Array.from({ length: owner.length - common }, () => ".."),
    ...target.slice(common),
  ].join("/");
}

export function compileChartParts(ctx: DocxWriteContext): {
  Charts?: XmlifyedFile[];
  ChartEmbeddings?: XmlifyedFile[];
} {
  if (ctx.charts.array.length === 0) return {};
  return {
    Charts: ctx.charts.array.flatMap((chartData, i) => {
      const chartPath = chartData.sourcePath ?? `word/charts/chart${i + 1}.xml`;
      const parts: Array<{ data: string; path: string }> = [
        {
          data: XML_DECL + chartData.chartSpaceXml,
          path: chartPath,
        },
      ];
      // User-shapes part behind c:userShapes: the chart's own rels
      // entry plus the body part (chartUserShapes, same directory).
      if (chartData.userShapes) {
        const userShapesPath = chartData.userShapes.path ?? `word/charts/userShapes${i + 1}.xml`;
        parts.push({
          data: XML_DECL + chartData.userShapes.xml,
          path: userShapesPath,
        });
      }
      // Embedded workbook behind c:externalData rides in the same rels
      // part. Relationship ids are carried verbatim so the re-emitted
      // r:ids resolve without rewriting.
      const relationships = new Relationships(chartPath);
      for (const rel of ctx._options.passthroughRelationships ?? []) {
        if (rel.source === chartPath) relationships.claimSourceRel(rel);
      }
      for (const rel of chartData.sourceRelationships ?? []) {
        relationships.claimSourceRel(rel);
      }
      if (chartData.embedding) {
        relationships.addRelationship(
          chartData.embedding.relationshipId,
          embeddingRelationshipType(chartData.embedding.relationshipType),
          relativeTarget(chartPath, `word/embeddings/${chartData.embedding.fileName}`),
        );
      }
      if (chartData.userShapes) {
        relationships.addRelationship(
          chartData.userShapes.relationshipId,
          CHART_USER_SHAPES_REL,
          relativeTarget(
            chartPath,
            chartData.userShapes.path ?? `word/charts/userShapes${i + 1}.xml`,
          ),
        );
      }
      if (relationships.relationshipCount > 0) {
        parts.push({
          data: XML_DECL + relationships.serialize(),
          path: partPathToRelsPath(chartPath),
        });
      }
      return parts;
    }),
    ChartEmbeddings: (() => {
      const seen = new Set<string>();
      const parts: { data: Uint8Array; path: string }[] = [];
      for (const c of ctx.charts.array) {
        const e = c.embedding;
        if (!e || seen.has(e.fileName)) continue;
        seen.add(e.fileName);
        parts.push({
          data: toUint8Array(e.data),
          path: `word/embeddings/${e.fileName}`,
        });
      }
      return parts;
    })(),
  };
}

export function compileChartExParts(ctx: DocxWriteContext): {
  ChartEx?: XmlifyedFile[];
} {
  if (ctx.chartExes.array.length === 0) return {};
  return {
    ChartEx: ctx.chartExes.array
      .map((chartExData, index) => {
        const chartExPath = chartExData.sourcePath ?? `word/charts/chartEx${index + 1}.xml`;
        const relationships = new Relationships(chartExPath);
        for (const rel of chartExData.sourceRelationships ?? []) {
          relationships.claimSourceRel(rel);
        }
        for (const rel of ctx._options.passthroughRelationships ?? []) {
          if (rel.source === chartExPath) relationships.claimSourceRel(rel);
        }
        const parts: Array<{ data: string; path: string }> = [
          { data: XML_DECL + chartExData.chartExXml, path: chartExPath },
        ];
        if (relationships.relationshipCount > 0) {
          parts.push({
            data: XML_DECL + relationships.serialize(),
            path: partPathToRelsPath(chartExPath),
          });
        }
        return parts;
      })
      .flat(),
  };
}

export function compileSmartArtParts(ctx: DocxWriteContext): {
  DiagramData?: XmlifyedFile[];
  DiagramLayout?: XmlifyedFile[];
  DiagramStyle?: XmlifyedFile[];
  DiagramColors?: XmlifyedFile[];
  DiagramDrawing?: XmlifyedFile[];
  SmartArtDataRels?: XmlifyedFile[];
} {
  if (ctx.smartArts.array.length === 0) return {};
  const smartArtParts = ctx.smartArts.array.map((smartArtData, i) => {
    const dataPath = smartArtData.sourcePaths?.data ?? `word/diagrams/data${i + 1}.xml`;
    const drawingFileIndex = ctx.smartArts.array.indexOf(smartArtData) + 1;
    const drawingPath =
      smartArtData.sourcePaths?.drawing ?? `word/diagrams/drawing${drawingFileIndex}.xml`;
    const hasDrawingRelationship =
      smartArtData.raw?.drawing !== undefined && smartArtData.raw.drawingRelationship !== false;
    const relationships = new Relationships(dataPath);
    for (const relationship of smartArtData.sourceRelationships ?? []) {
      relationships.claimSourceRel(relationship);
    }
    const drawingTarget = relativeTarget(dataPath, drawingPath);
    if (
      hasDrawingRelationship &&
      !relationships.hasRelationship(RELATIONSHIP_TYPES.diagramDrawingMs, drawingTarget)
    ) {
      relationships.add(RELATIONSHIP_TYPES.diagramDrawingMs, drawingTarget);
    }
    const media = smartArtData.raw?.media ?? [];
    const mediaIds = new Map<string, string>();
    for (const mediaReference of media) {
      const mediaPath = mediaReference.partPath ?? `word/media/${mediaReference.fileName}`;
      const target = relativeTarget(dataPath, mediaPath);
      const existingId = relationships.idOf(RELATIONSHIP_TYPES.image, target);
      mediaIds.set(
        mediaReference.fileName,
        existingId ?? `rId${relationships.add(RELATIONSHIP_TYPES.image, target)}`,
      );
    }
    return {
      smartArtData,
      dataPath,
      relationships,
      dataXml: replacePlaceholders(
        XML_DECL + withDiagramNamespaces(smartArtData.dataModelXml),
        mediaIds,
      ),
      media,
    };
  });
  return {
    DiagramData: smartArtParts.map(({ dataPath, dataXml }) => ({
      data: dataXml,
      path: dataPath,
    })),
    DiagramLayout: ctx.smartArts.array.map((smartArtData, i) => ({
      data:
        typeof smartArtData.layout === "string"
          ? getLayoutXml(smartArtData.layout)
          : stringifyLayoutDefinitionPart(smartArtData.layout),
      path: smartArtData.sourcePaths?.layout ?? `word/diagrams/layout${i + 1}.xml`,
    })),
    DiagramStyle: ctx.smartArts.array.map((smartArtData, i) => ({
      data:
        typeof smartArtData.style === "string"
          ? getStyleXml(smartArtData.style)
          : stringifyStyleDefinitionPart(smartArtData.style),
      path: smartArtData.sourcePaths?.quickStyle ?? `word/diagrams/quickStyle${i + 1}.xml`,
    })),
    DiagramColors: ctx.smartArts.array.map((smartArtData, i) => ({
      data:
        typeof smartArtData.color === "string"
          ? getColorXml(smartArtData.color)
          : stringifyColorDefinitionPart(smartArtData.color),
      path: smartArtData.sourcePaths?.colors ?? `word/diagrams/colors${i + 1}.xml`,
    })),
    DiagramDrawing: ctx.smartArts.array
      .map((smartArtData, i) => ({ smartArtData, i }))
      .filter(({ smartArtData }) => smartArtData.raw?.drawing !== undefined)
      .flatMap(({ smartArtData, i }) => {
        const drawingPath =
          smartArtData.sourcePaths?.drawing ?? `word/diagrams/drawing${i + 1}.xml`;
        return [
          {
            data: toUint8Array(smartArtData.raw!.drawing!),
            path: drawingPath,
          },
          ...(smartArtData.raw?.drawingRels === undefined
            ? []
            : [
                {
                  data: toUint8Array(smartArtData.raw.drawingRels),
                  path: partPathToRelsPath(drawingPath),
                },
              ]),
        ];
      }),
    SmartArtDataRels: smartArtParts.flatMap(({ dataPath, relationships }) => {
      if (relationships.relationshipCount === 0) return [];
      return [
        {
          data: XML_DECL + relationships.serialize(),
          path: partPathToRelsPath(dataPath),
        },
      ];
    }),
  };
}
