/**
 * Relationship wiring shared by the docx compile phases: the XML declaration,
 * the OLE/PACKAGE embedding relationship resolution, and the per-part media +
 * embedding placeholder bridge.
 *
 * @module
 */

import {
  findAndReplaceImagePlaceholders,
  Relationships,
  replaceAllPlaceholders,
  type RelationshipType,
} from "@office-open/core";
import { RELATIONSHIP_TYPES, relativePartTarget } from "@office-open/core/opc";
import { OOXML_XML_DECLARATION } from "@office-open/xml";
import type { EmbeddingCollection } from "@shared/embeddings/embeddings";

import type { DocxWriteContext } from "../context";

/** XML declaration prepended to every OOXML part. */
export const XML_DECL = OOXML_XML_DECLARATION;

/** Source relationships owned by one part, preserved verbatim for a rebuilt
 * part whose rels the compiler re-emits. */
export function sourceOwnedRels(ctx: DocxWriteContext, sourcePath: string): Relationships {
  const rels = new Relationships();
  for (const rel of ctx._options.passthroughRelationships ?? []) {
    if (rel.source === sourcePath) rels.claimSourceRel(rel);
  }
  return rels;
}

/** {fileName → source rId} for media rels owned by one part. */
export function sourceOwnedMediaRids(ctx: DocxWriteContext, partPath: string): Map<string, string> {
  const sourceRids = new Map<string, string>();
  for (const rel of ctx._options.passthroughRelationships ?? []) {
    if (rel.source !== partPath) continue;
    const match = /(?:^|\/)media\/([^/]+)$/.exec(rel.target);
    if (match) sourceRids.set(match[1]!, rel.rId);
  }
  return sourceRids;
}

/** Serialize numbering media while preserving unused source-owned rels. */
export function resolveNumberingMedia(
  xml: string,
  ctx: DocxWriteContext,
  ownerPath = "word/numbering.xml",
): { xml: string; rels: Relationships } {
  const rels = sourceOwnedRels(ctx, ownerPath);
  const media = findAndReplaceImagePlaceholders(
    xml,
    ctx.media.array,
    1,
    "rId",
    sourceOwnedMediaRids(ctx, ownerPath),
  );
  for (const ref of media.referenced) {
    const target = relativePartTarget(ownerPath, `word/media/${ref.fileName}`);
    if (!rels.hasRelationship(RELATIONSHIP_TYPES.image, target)) {
      rels.add(RELATIONSHIP_TYPES.image, target);
    }
  }
  return { xml: media.xml, rels };
}

/** Relationship type for OLE embedding parts (word|ppt/embeddings/*). */
const OLE_OBJECT_RELATIONSHIP = RELATIONSHIP_TYPES.oleObject;

/** Relationship type for native-format embedding parts (embedded xlsx/docx). */
export const PACKAGE_RELATIONSHIP = RELATIONSHIP_TYPES.package;

/** Re-emit the relationship type the source used for an embedding part — a
 *  native OPC package (xlsx/docx) stays a package rel, an OLE compound binary
 *  stays an oleObject rel. */
export const embeddingRelationship = (
  embeddings: EmbeddingCollection,
  fileName: string,
): RelationshipType =>
  embeddings.array.find((e) => e.fileName === fileName)?.relationshipType === "package"
    ? PACKAGE_RELATIONSHIP
    : OLE_OBJECT_RELATIONSHIP;

/** Resolve a source embedding relationship kind while defaulting fresh authoring. */
export const embeddingRelationshipType = (
  kind: "oleObject" | "package" | undefined,
): RelationshipType => (kind === "oleObject" ? OLE_OBJECT_RELATIONSHIP : PACKAGE_RELATIONSHIP);

/** Relative relationship target from one package part to another. */
export { relativePartTarget };

/** Resolved media/embedding placeholders for one part, with the offsets its
 *  relationship registrations must use (ids are per-part numbering). */
export interface PartMediaResolution {
  xml: string;
  relCount: number;
  embeddingOffset: number;
  mediaRefs: { fileName: string }[];
  embeddingRefs: { fileName: string }[];
}

/**
 * Resolve a part's {fileName} media placeholders against the package media
 * store, then its embedding placeholders chained past the media
 * relationships (same ordering as the document part — headers/footers and
 * notes can carry w:object runs of their own).
 */
export function resolvePartMedia(
  xml: string,
  ctx: DocxWriteContext,
  relCount: number,
): PartMediaResolution {
  const media = findAndReplaceImagePlaceholders(xml, ctx.media.array, relCount);
  const embeddingOffset = relCount + media.referenced.length;
  const embeddings = findAndReplaceImagePlaceholders(
    media.xml,
    ctx.embeddings.array,
    embeddingOffset,
  );
  return {
    xml: embeddings.xml,
    relCount,
    embeddingOffset,
    mediaRefs: media.referenced,
    embeddingRefs: embeddings.referenced,
  };
}

/** Register the resolved media + embedding relationships on the part's rels. */
export function registerPartMedia(
  rels: Relationships,
  ctx: DocxWriteContext,
  resolved: PartMediaResolution,
): void {
  for (const [i, ref] of resolved.mediaRefs.entries()) {
    rels.addRelationship(resolved.relCount + i, RELATIONSHIP_TYPES.image, `media/${ref.fileName}`);
  }
  for (const [i, ref] of resolved.embeddingRefs.entries()) {
    rels.addRelationship(
      resolved.embeddingOffset + i,
      embeddingRelationship(ctx.embeddings, ref.fileName),
      `embeddings/${ref.fileName}`,
    );
  }
}

/** Matches any chart/SmartArt placeholder prefix in a part's XML. */
const CHART_SMARTART_PLACEHOLDER = /\{(chart:|smartart(?:-lo|-qs|-cs)?:)/;

/**
 * Resolve a part's {chart:key} / {smartart*:key} placeholders and register the
 * referenced parts on the part's own rels — headers/footers and notes can
 * carry chart and SmartArt drawings of their own. The document part resolves
 * the full collections (compile/document.ts); here only keys this part
 * actually references get relationships, at ids continuing past the part's
 * media/embedding block. Targets keep the package-wide collection order
 * (charts/chartN.xml, diagrams/{data,layout,quickStyle,colors}N.xml), so
 * several parts referencing one chart all point at the same file.
 */
export function resolvePartCharts(
  xml: string,
  ctx: DocxWriteContext,
  rels: Relationships,
  relCount: number,
  ownerPath = "word/document.xml",
): string {
  if (!CHART_SMARTART_PLACEHOLDER.test(xml)) return xml;
  const entries: Array<{ prefix?: string; key: string; value: string }> = [];
  const referencedCharts = ctx.charts.array.filter((c) => xml.includes(`{chart:${c.key}}`));
  referencedCharts.forEach((chart, i) => {
    entries.push({ prefix: "chart:", key: chart.key, value: `rId${relCount + i}` });
    rels.addRelationship(
      relCount + i,
      RELATIONSHIP_TYPES.chart,
      relativePartTarget(
        ownerPath,
        chart.sourcePath ?? `word/charts/chart${ctx.charts.array.indexOf(chart) + 1}.xml`,
      ),
    );
  });
  const referencedSmartArts = ctx.smartArts.array.filter((s) =>
    ["", "-lo", "-qs", "-cs"].some((suffix) => xml.includes(`{smartart${suffix}:${s.key}}`)),
  );
  const chartCount = referencedCharts.length;
  const smartArtCount = referencedSmartArts.length;
  const base = relCount + chartCount;
  const loOffset = base + smartArtCount;
  const qsOffset = loOffset + smartArtCount;
  const csOffset = qsOffset + smartArtCount;
  const drawingOffset = csOffset + smartArtCount;
  referencedSmartArts.forEach((smartArt, i) => {
    const fileIndex = ctx.smartArts.array.indexOf(smartArt) + 1;
    const relsByPrefix: Array<
      [string, number, RelationshipType, "data" | "layout" | "quickStyle" | "colors"]
    > = [
      ["smartart:", base, RELATIONSHIP_TYPES.diagramData, "data"],
      ["smartart-lo:", loOffset, RELATIONSHIP_TYPES.diagramLayout, "layout"],
      ["smartart-qs:", qsOffset, RELATIONSHIP_TYPES.diagramQuickStyle, "quickStyle"],
      ["smartart-cs:", csOffset, RELATIONSHIP_TYPES.diagramColors, "colors"],
    ];
    for (const [prefix, offset, type, file] of relsByPrefix) {
      entries.push({ prefix, key: smartArt.key, value: `rId${offset + i}` });
      const sourcePath = smartArt.sourcePaths?.[file === "quickStyle" ? "quickStyle" : file];
      rels.addRelationship(
        offset + i,
        type,
        relativePartTarget(ownerPath, sourcePath ?? `word/diagrams/${file}${fileIndex}.xml`),
      );
    }
    // The drawing part is an Office render cache, present only when the source
    // carried it — Word never emits it for a fresh SmartArt.
    if (smartArt.raw?.drawing !== undefined) {
      rels.addRelationship(
        drawingOffset + i,
        RELATIONSHIP_TYPES.diagramDrawingMs,
        relativePartTarget(
          ownerPath,
          smartArt.sourcePaths?.drawing ?? `word/diagrams/drawing${fileIndex}.xml`,
        ),
      );
    }
  });
  return replaceAllPlaceholders(xml, entries);
}
