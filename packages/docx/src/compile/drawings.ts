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
  RELATIONSHIP_TYPES,
  type XmlifyedFile,
  toUint8Array,
  partPathToRelsPath,
} from "@office-open/core";
import {
  getColorXml,
  getLayoutXml,
  getStyleXml,
  stringifyColorDefinitionPart,
  stringifyLayoutDefinitionPart,
  stringifyStyleDefinitionPart,
} from "@office-open/core/smartart";

import type { DocxWriteContext } from "../context";
import { PACKAGE_RELATIONSHIP, XML_DECL } from "./shared";

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
      if (chartData.embedding) {
        relationships.addRelationship(
          chartData.embedding.relationshipId,
          PACKAGE_RELATIONSHIP,
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
      for (const rel of ctx._options.passthroughRelationships ?? []) {
        if (rel.source === chartPath) relationships.claimSourceRel(rel);
      }
      for (const rel of chartData.sourceRelationships ?? []) {
        relationships.claimSourceRel(rel);
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

export function compileSmartArtParts(ctx: DocxWriteContext): {
  DiagramData?: XmlifyedFile[];
  DiagramLayout?: XmlifyedFile[];
  DiagramStyle?: XmlifyedFile[];
  DiagramColors?: XmlifyedFile[];
  DiagramDrawing?: XmlifyedFile[];
  SmartArtDataRels?: XmlifyedFile[];
} {
  if (ctx.smartArts.array.length === 0) return {};
  return {
    DiagramData: ctx.smartArts.array.map((smartArtData, i) => ({
      data:
        smartArtData.raw?.data !== undefined
          ? toUint8Array(smartArtData.raw.data)
          : XML_DECL + smartArtData.dataModelXml,
      path: smartArtData.sourcePaths?.data ?? `word/diagrams/data${i + 1}.xml`,
    })),
    DiagramLayout: ctx.smartArts.array.map((smartArtData, i) => ({
      data:
        smartArtData.raw?.layout !== undefined
          ? toUint8Array(smartArtData.raw.layout)
          : typeof smartArtData.layout === "string"
            ? getLayoutXml(smartArtData.layout)
            : stringifyLayoutDefinitionPart(smartArtData.layout),
      path: smartArtData.sourcePaths?.layout ?? `word/diagrams/layout${i + 1}.xml`,
    })),
    DiagramStyle: ctx.smartArts.array.map((smartArtData, i) => ({
      data:
        smartArtData.raw?.style !== undefined
          ? toUint8Array(smartArtData.raw.style)
          : typeof smartArtData.style === "string"
            ? getStyleXml(smartArtData.style)
            : stringifyStyleDefinitionPart(smartArtData.style),
      path: smartArtData.sourcePaths?.quickStyle ?? `word/diagrams/quickStyle${i + 1}.xml`,
    })),
    DiagramColors: ctx.smartArts.array.map((smartArtData, i) => ({
      data:
        smartArtData.raw?.color !== undefined
          ? toUint8Array(smartArtData.raw.color)
          : typeof smartArtData.color === "string"
            ? getColorXml(smartArtData.color)
            : stringifyColorDefinitionPart(smartArtData.color),
      path: smartArtData.sourcePaths?.colors ?? `word/diagrams/colors${i + 1}.xml`,
    })),
    DiagramDrawing: ctx.smartArts.array
      .map((smartArtData, i) => ({ smartArtData, i }))
      .filter(({ smartArtData }) => smartArtData.raw?.drawing !== undefined)
      .map(({ smartArtData, i }) => ({
        data: toUint8Array(smartArtData.raw!.drawing!),
        path: smartArtData.sourcePaths?.drawing ?? `word/diagrams/drawing${i + 1}.xml`,
      })),
    // Data parts with their own rels (blipFill art): re-emit the rels
    // verbatim — its rIds and ../media targets match the pinned media.
    ...(ctx.smartArts.array.some((s) => s.raw?.dataRels !== undefined)
      ? {
          SmartArtDataRels: ctx.smartArts.array.flatMap((smartArtData, i) =>
            smartArtData.raw?.dataRels !== undefined
              ? [
                  {
                    data: toUint8Array(smartArtData.raw.dataRels),
                    path:
                      smartArtData.sourcePaths?.dataRels ??
                      `word/diagrams/_rels/data${i + 1}.xml.rels`,
                  },
                ]
              : [],
          ),
        }
      : {}),
  };
}
