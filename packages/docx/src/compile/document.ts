/**
 * Document part compile phase: word/document.xml (chart/SmartArt/numbering
 * placeholder resolution over the stringified body) and
 * word/_rels/document.xml.rels.
 *
 * The document part is the only one whose media/embedding relationships
 * re-emit the source rIds on round-trip: the body XML references them
 * verbatim, so registration goes through sourceRidFor before falling back to
 * the offset-derived id.
 *
 * @module
 */

import {
  addSmartArtRelationships,
  findAndReplaceImagePlaceholders,
  formatId,
  hasPlaceholders,
  replaceAllPlaceholders,
  replaceNumberingPlaceholders,
} from "@office-open/core";
import {
  partPathToRelsPath,
  RELATIONSHIP_TYPES,
  resolveRelationshipTarget,
  type XmlifyedFile,
} from "@office-open/core/opc";

import type { DocxWriteContext } from "../context";
import { XML_DECL, embeddingRelationship } from "./shared";

/**
 * Look up the source rId for a document relationship by kind+target. Returns
 * the numeric id when the source carried the same rel, undefined otherwise.
 * The body XML's `r:id` references the source rId verbatim on round-trip, so
 * the structured compiler must emit the rel at that exact id — otherwise
 * `r:id="rId4"` would point at a different rel and Word refuses to open.
 */
function sourceRidFor(
  passthroughRelationships:
    | readonly { source: string; relationshipType: string; target: string; rId: string }[]
    | undefined,
  ownerSource: string,
  relationshipType: string,
  target: string,
): number | undefined {
  if (!passthroughRelationships) return undefined;
  for (const rel of passthroughRelationships) {
    const kind = relationshipType.split("/").pop();
    if (
      rel.source === ownerSource &&
      rel.relationshipType.split("/").pop() === kind &&
      resolveRelationshipTarget(ownerSource, rel.target) ===
        resolveRelationshipTarget(ownerSource, target)
    ) {
      const m = /^rId(\d+)$/.exec(rel.rId);
      if (m) return Number(m[1]);
    }
  }
  return undefined;
}

function sourceRidForPath(
  passthroughRelationships: readonly { source: string; target: string; rId: string }[] | undefined,
  ownerSource: string,
  target: string,
): number | undefined {
  const source = passthroughRelationships?.find(
    (rel) =>
      rel.source === ownerSource &&
      resolveRelationshipTarget(ownerSource, rel.target) ===
        resolveRelationshipTarget(ownerSource, target),
  );
  const match = source ? /^rId(\d+)$/.exec(source.rId) : undefined;
  return match ? Number(match[1]) : undefined;
}

function sourceRelationshipFor(
  passthroughRelationships:
    | readonly {
        source: string;
        relationshipType: string;
        target: string;
        rId: string;
        targetMode?: "External";
      }[]
    | undefined,
  ownerSource: string,
  relationshipType: string,
  target: string,
) {
  if (!passthroughRelationships) return undefined;
  const kind = relationshipType.split("/").pop();
  return passthroughRelationships.find(
    (rel) =>
      rel.source === ownerSource &&
      rel.relationshipType.split("/").pop() === kind &&
      resolveRelationshipTarget(ownerSource, rel.target) ===
        resolveRelationshipTarget(ownerSource, target),
  );
}

function sourceRelationshipForPath(
  passthroughRelationships:
    | readonly {
        source: string;
        relationshipType: string;
        target: string;
        rId: string;
        targetMode?: "External";
      }[]
    | undefined,
  ownerSource: string,
  target: string,
) {
  return passthroughRelationships?.find(
    (rel) =>
      rel.source === ownerSource &&
      resolveRelationshipTarget(ownerSource, rel.target) ===
        resolveRelationshipTarget(ownerSource, target),
  );
}

/** Convert a source package path to a target relative to word/document.xml. */
function documentRelationshipTarget(partPath: string, ownerPath: string): string {
  if (partPath.startsWith("/")) return partPath;
  const owner = ownerPath.split("/").slice(0, -1);
  const target = partPath.split("/");
  let common = 0;
  while (common < owner.length && common < target.length - 1 && owner[common] === target[common])
    common++;
  return [
    ...Array.from({ length: owner.length - common }, () => ".."),
    ...target.slice(common),
  ].join("/");
}

/**
 * Build a {fileName → rId} map for source rels targeting `media/` or
 * `embeddings/`. findAndReplaceImagePlaceholders consults this map to
 * substitute the source rId in body XML (so `<a:blip r:embed="rId5"/>` points
 * at the rel registered at rId5 — not the rel registered at an
 * offset-derived id, which dangles against any source body that referenced
 * the source rId). Placeholder keys use the media-collection fileName; a
 * dedup-renamed file misses the map and falls through to the offset id on
 * BOTH the body and the .rels side, keeping them consistent.
 */
function documentSourceRids(
  ctx: DocxWriteContext,
  parts: readonly { fileName: string; partPath?: string }[],
  fallbackDirectory: "media" | "embeddings",
): ReadonlyMap<string, string> {
  const map = new Map<string, string>();
  const targets = new Map(
    parts.map((part) => [
      part.partPath ?? `word/${fallbackDirectory}/${part.fileName}`,
      part.fileName,
    ]),
  );
  for (const rel of ctx._options.passthroughRelationships ?? []) {
    if (rel.source !== ctx.primaryDocumentPath) continue;
    const fileName = targets.get(resolveRelationshipTarget(rel.source, rel.target));
    if (fileName !== undefined && !map.has(fileName)) map.set(fileName, rel.rId);
  }
  return map;
}

/** Resolve the document body's media/embedding placeholders, then emit the
 * document part and its rels. The rels entry registers media, embeddings,
 * charts, SmartArt, the font table, and the claimed passthrough rels — in
 * that order, so the source rId ordering is preserved. */
export function compileDocumentEntries(
  ctx: DocxWriteContext,
  documentXmlData: string,
  documentRelationshipCount: number,
): { Document: XmlifyedFile; Relationships: XmlifyedFile } {
  const documentPath = ctx.primaryDocumentPath;
  const documentMedia = findAndReplaceImagePlaceholders(
    documentXmlData,
    ctx.media.array,
    documentRelationshipCount,
    "rId",
    documentSourceRids(ctx, ctx.media.array, "media"),
  );
  // OLE embeddings reuse the same {fileName} placeholder bridge as images; run
  // after media so {oleObjectN.bin} placeholders resolve against the embedding array.
  const documentEmbeddingOffset = documentRelationshipCount + documentMedia.referenced.length;
  const documentEmbeddings = findAndReplaceImagePlaceholders(
    documentMedia.xml,
    ctx.embeddings.array,
    documentEmbeddingOffset,
    "rId",
    documentSourceRids(ctx, ctx.embeddings.array, "embeddings"),
  );

  return {
    Document: {
      data: (() => {
        let xmlData = documentEmbeddings.xml;
        if (hasPlaceholders(xmlData)) {
          const mediaCount = documentMedia.referenced.length;
          const embeddingCount = documentEmbeddings.referenced.length;
          const chartKeys = ctx.charts.array.map((c) => c.key);
          const chartOffset = documentRelationshipCount + mediaCount + embeddingCount;
          const chartExKeys = ctx.chartExes.array.map((c) => c.key);
          const chartExOffset = chartOffset + chartKeys.length;
          const smartArtOffset = chartExOffset + chartExKeys.length;
          const smartArtKeys = ctx.smartArts.array.map((s) => s.key);

          // Build combined replacement entries for charts, smartart, and numbering
          const entries: Array<{ prefix?: string; key: string; value: string }> = [];
          for (const [i, key] of chartKeys.entries()) {
            const chartTarget = documentRelationshipTarget(
              ctx.charts.array[i]?.sourcePath ?? `word/charts/chart${i + 1}.xml`,
              documentPath,
            );
            const sourceRid = sourceRidFor(
              ctx._options.passthroughRelationships,
              documentPath,
              RELATIONSHIP_TYPES.chart,
              chartTarget,
            );
            entries.push({
              prefix: "chart:",
              key,
              value: sourceRid !== undefined ? `rId${sourceRid}` : formatId(chartOffset, i, "rId"),
            });
          }
          for (const [i, key] of chartExKeys.entries()) {
            const chartExTarget = documentRelationshipTarget(
              ctx.chartExes.array[i]?.sourcePath ?? `word/charts/chartEx${i + 1}.xml`,
              documentPath,
            );
            const sourceRid = sourceRidForPath(
              ctx._options.passthroughRelationships,
              documentPath,
              chartExTarget,
            );
            entries.push({
              prefix: "chartex:",
              key,
              value:
                sourceRid !== undefined ? `rId${sourceRid}` : formatId(chartExOffset, i, "rId"),
            });
          }
          const saPrefixes = ["smartart:", "smartart-lo:", "smartart-qs:", "smartart-cs:"];
          const saRelationshipTypes = [
            RELATIONSHIP_TYPES.diagramData,
            RELATIONSHIP_TYPES.diagramLayout,
            RELATIONSHIP_TYPES.diagramQuickStyle,
            RELATIONSHIP_TYPES.diagramColors,
          ];
          for (const [i, key] of smartArtKeys.entries()) {
            for (let p = 0; p < saPrefixes.length; p++) {
              const smartArt = ctx.smartArts.array[i];
              const sourcePath =
                smartArt?.sourcePaths?.[
                  p === 1 ? "layout" : p === 2 ? "quickStyle" : p === 3 ? "colors" : "data"
                ];
              const sourceRid = sourcePath
                ? sourceRidFor(
                    ctx._options.passthroughRelationships,
                    documentPath,
                    saRelationshipTypes[p]!,
                    documentRelationshipTarget(sourcePath, documentPath),
                  )
                : undefined;
              entries.push({
                prefix: saPrefixes[p],
                key,
                value:
                  sourceRid !== undefined
                    ? `rId${sourceRid}`
                    : formatId(smartArtOffset + p * smartArtKeys.length, i, "rId"),
              });
            }
          }
          for (const { reference, instance, numId } of ctx.numbering.concreteNumbering) {
            entries.push({ key: `${reference}-${instance}`, value: numId.toString() });
          }
          xmlData = replaceAllPlaceholders(xmlData, entries);
        } else {
          xmlData = replaceNumberingPlaceholders(xmlData, ctx.numbering.concreteNumbering);
        }
        return xmlData;
      })(),
      path: documentPath,
    },
    Relationships: {
      data: (() => {
        for (const [i, ref] of documentMedia.referenced.entries()) {
          const partPath =
            ctx.media.array.find((media) => media.fileName === ref.fileName)?.partPath ??
            `word/media/${ref.fileName}`;
          const target = documentRelationshipTarget(partPath, documentPath);
          const sourceRel = sourceRelationshipFor(
            ctx._options.passthroughRelationships,
            documentPath,
            RELATIONSHIP_TYPES.image,
            target,
          );
          ctx.document.relationships.addIfAbsent(
            sourceRel?.rId ?? documentRelationshipCount + i,
            sourceRel?.relationshipType ?? RELATIONSHIP_TYPES.image,
            sourceRel?.target ?? target,
            sourceRel?.targetMode,
          );
        }
        for (const [i, ref] of documentEmbeddings.referenced.entries()) {
          const target = documentRelationshipTarget(
            `word/embeddings/${ref.fileName}`,
            documentPath,
          );
          const sourceRel = sourceRelationshipFor(
            ctx._options.passthroughRelationships,
            documentPath,
            embeddingRelationship(ctx.embeddings, ref.fileName),
            target,
          );
          ctx.document.relationships.addIfAbsent(
            sourceRel?.rId ?? documentEmbeddingOffset + i,
            sourceRel?.relationshipType ?? embeddingRelationship(ctx.embeddings, ref.fileName),
            sourceRel?.target ?? target,
            sourceRel?.targetMode,
          );
        }

        const chartOffset =
          documentRelationshipCount +
          documentMedia.referenced.length +
          documentEmbeddings.referenced.length;
        for (let i = 0; i < ctx.charts.array.length; i++) {
          const target = documentRelationshipTarget(
            ctx.charts.array[i]?.sourcePath ?? `word/charts/chart${i + 1}.xml`,
            documentPath,
          );
          const sourceRel = sourceRelationshipFor(
            ctx._options.passthroughRelationships,
            documentPath,
            RELATIONSHIP_TYPES.chart,
            target,
          );
          ctx.document.relationships.addIfAbsent(
            sourceRel?.rId ?? chartOffset + i,
            sourceRel?.relationshipType ?? RELATIONSHIP_TYPES.chart,
            sourceRel?.target ?? target,
            sourceRel?.targetMode,
          );
        }

        const chartExOffset =
          documentRelationshipCount +
          documentMedia.referenced.length +
          documentEmbeddings.referenced.length +
          ctx.charts.array.length;
        for (let i = 0; i < ctx.chartExes.array.length; i++) {
          const target = documentRelationshipTarget(
            ctx.chartExes.array[i]?.sourcePath ?? `word/charts/chartEx${i + 1}.xml`,
            documentPath,
          );
          const sourceRel = sourceRelationshipForPath(
            ctx._options.passthroughRelationships,
            documentPath,
            target,
          );
          ctx.document.relationships.addIfAbsent(
            sourceRel?.rId ?? chartExOffset + i,
            sourceRel?.relationshipType ?? RELATIONSHIP_TYPES.chartEx,
            sourceRel?.target ?? target,
            sourceRel?.targetMode,
          );
        }

        addSmartArtRelationships(
          ctx.smartArts.array.map((s) => s.key),
          (id, type, target) => {
            const sourceRel = sourceRelationshipFor(
              ctx._options.passthroughRelationships,
              documentPath,
              type,
              target,
            );
            ctx.document.relationships.addIfAbsent(
              sourceRel?.rId ?? id,
              sourceRel?.relationshipType ?? type,
              sourceRel?.target ?? target,
              sourceRel?.targetMode,
            );
          },
          documentRelationshipCount +
            documentMedia.referenced.length +
            documentEmbeddings.referenced.length +
            ctx.charts.array.length +
            ctx.chartExes.array.length,
          0,
          {
            pathPrefix: "",
            styleRelType: RELATIONSHIP_TYPES.diagramQuickStyle,
            sourceTargetFor: (key, kind) => {
              const sourcePath = ctx.smartArts.array.find((s) => s.key === key)?.sourcePaths?.[
                kind === "quickStyle" ? "quickStyle" : kind
              ];
              return sourcePath ? documentRelationshipTarget(sourcePath, documentPath) : undefined;
            },
            // The drawing part is an Office render cache, present only when the
            // source carried it — Word never emits it for a fresh SmartArt.
            hasDrawing: (key) =>
              ctx.smartArts.array.find((s) => s.key === key)?.raw?.drawing !== undefined,
          },
        );

        ctx.registerFontTableRelationship();
        ctx.addPassthroughDocumentRelationships();

        return XML_DECL + ctx.document.relationships.serialize();
      })(),
      path: partPathToRelsPath(documentPath),
    },
  };
}
