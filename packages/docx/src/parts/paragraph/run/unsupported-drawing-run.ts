import type { LongHexNumber } from "@office-open/core";
import { TargetModeType } from "@office-open/core";
import type { Element } from "@office-open/xml";
import { stringifyElement } from "@office-open/xml";
import type { BackgroundRawMediaOptions } from "@parts/document/document-background/document-background";

import type { BodyContext } from "../../../context";
import type { RunPropertiesOptions } from "./properties";

/** Canonical DrawingML node retained when no format-specific model exists. */
export interface XmlElementData {
  /** Qualified source name (for example, `wps:wsp`). */
  name: string;
  /** Qualified attributes, including namespace declarations when present. */
  attributes?: Record<string, string | number>;
  /** Character content at this node. */
  text?: string | number | boolean;
  /** Child elements in source order. */
  elements?: XmlElementData[];
}

/**
 * A source relationship referenced by unsupported DrawingML content.
 * Round-trip only — do not hand-author.
 */
export interface UnsupportedDrawingRelationshipOptions {
  /** Source relationship id; stringify restores it when available. */
  sourceRelationshipId: string;
  /** Full OPC relationship Type. */
  relationshipType: string;
  /** OPC target URI or package-relative target. */
  target: string;
  /** OPC target mode; absent = internal. */
  targetMode?: "External";
}

/** A group child retained when no format-specific child model exists. */
export interface UnsupportedGroupChildMediaData {
  type: "unsupported";
  /** Exact source child element. */
  element: XmlElementData;
  /** Relationship references from its non-visual properties. */
  relationships?: UnsupportedDrawingRelationshipOptions[];
}

/**
 * Typed fallback for an unrecognized `a:graphicData` payload. Round-trip
 * only — modelable drawing families have dedicated options.
 *
 * @publicApi
 */
export interface UnsupportedDrawingOptions {
  /** Exact `wp:inline` or `wp:anchor` wrapper, without its `a:graphic`. */
  anchor: XmlElementData;
  /** Anchor mode, mirroring the source `wp:inline` or `wp:anchor` element. */
  anchorMode: "inline" | "anchor";
  /** Exact unsupported `a:graphicData` child subtree. */
  graphicData: XmlElementData;
  /** Exact `a:graphicData` URI; empty only for malformed legacy sources. */
  graphicDataUri?: string;
  /** Hyperlink relationships referenced from DrawingML non-visual properties. */
  relationships?: UnsupportedDrawingRelationshipOptions[];
  /**
   * Raw XML of the wrapping mc:AlternateContent's mc:Fallback (VML or picture
   * equivalent) — carried verbatim so the full wrapper round-trips.
   * Round-trip channel: captured from a parsed source document — do not
   * hand-author.
   */
  vmlFallback?: string;
  /** Media referenced by `vmlFallback` `{fileName}` placeholders, registered on generate. */
  vmlFallbackMedia?: BackgroundRawMediaOptions[];
  /** mc:Choice Requires attribute (e.g. "cx1") used to regenerate the AlternateContent wrapper. */
  mcChoiceRequires?: string;
  /** Structured run properties of the wrapping `w:r` (round-trip). */
  runProperties?: RunPropertiesOptions;
  /** Revision save ID of the wrapping `w:r` (`w:rsidR`, round-trip). */
  additionRsid?: LongHexNumber;
  /** Revision save ID of the wrapping `w:r` properties (round-trip). */
  runPropertiesRsid?: LongHexNumber;
  /** A `w:lastRenderedPageBreak` shared with the drawing run (round-trip). */
  lastRenderedPageBreak?: boolean;
}

function toXmlElement(el: XmlElementData): Element {
  return {
    type: "element",
    name: el.name,
    ...(el.attributes ? { attributes: el.attributes } : {}),
    ...(el.elements?.length
      ? {
          elements: el.elements.map((child) => ({
            ...toXmlElement(child),
            type: "element" as const,
          })),
        }
      : el.text !== undefined
        ? { elements: [{ type: "text" as const, text: el.text }] }
        : {}),
  };
}

function remapRelationshipIds(el: XmlElementData, ids: Map<string, string>): void {
  const value = el.attributes?.["r:id"];
  const mapped = typeof value === "string" ? ids.get(value) : undefined;
  if (mapped) el.attributes!["r:id"] = mapped;
  for (const child of el.elements ?? []) remapRelationshipIds(child, ids);
}

export function restoreUnsupportedDrawingRelationships(
  relationships: NonNullable<UnsupportedDrawingOptions["relationships"]>,
  ctx: BodyContext,
): Map<string, string> {
  const ids = new Map<string, string>();
  for (const relationship of relationships) {
    const sourceId = relationship.sourceRelationshipId;
    const targetMode = relationship.targetMode === "External" ? TargetModeType.EXTERNAL : undefined;
    let id: string;
    if (/^rId\d+$/.test(sourceId)) {
      if (
        !ctx.viewWrapper.relationships.hasExactRelationship(
          sourceId,
          relationship.relationshipType,
          relationship.target,
          targetMode,
        )
      ) {
        ctx.viewWrapper.relationships.addRelationship(
          sourceId,
          relationship.relationshipType,
          relationship.target,
          targetMode,
        );
      }
      id = sourceId;
    } else {
      const existingId = ctx.viewWrapper.relationships.idOf(
        relationship.relationshipType,
        relationship.target,
      );
      id =
        existingId ??
        `rId${ctx.viewWrapper.relationships.add(
          relationship.relationshipType,
          relationship.target,
          targetMode,
        )}`;
    }
    ids.set(sourceId, id);
  }
  return ids;
}

export function stringifyUnsupportedXmlElementData(
  element: XmlElementData,
  relationships: NonNullable<UnsupportedDrawingOptions["relationships"]> | undefined,
  ctx: BodyContext,
): string {
  const clone: XmlElementData = structuredClone(element);
  remapRelationshipIds(clone, restoreUnsupportedDrawingRelationships(relationships ?? [], ctx));
  return stringifyElement(toXmlElement(clone));
}

export function stringifyUnsupportedDrawing(
  options: UnsupportedDrawingOptions,
  ctx: BodyContext,
): string {
  const graphicData: XmlElementData = structuredClone(options.graphicData);
  if (options.graphicDataUri !== undefined) graphicData.attributes!.uri = options.graphicDataUri;
  const anchorXml = stringifyUnsupportedXmlElementData(options.anchor, options.relationships, ctx);
  const graphicXml = `<a:graphic>${stringifyUnsupportedXmlElementData(graphicData, options.relationships, ctx)}</a:graphic>`;
  // CT_Anchor ends at a:graphic; Word's mc extensions (wp14:sizeRelH/V) trail
  // it, so the graphic splices in before the first wp14:* sibling — falling
  // back to the anchor close for sources without extensions.
  const anchorOpen = anchorXml.indexOf(">");
  const firstExtension = anchorXml.search(/<wp14:sizeRel[HV]/);
  const spliceAt =
    anchorOpen >= 0 && firstExtension > anchorOpen ? firstExtension : anchorXml.lastIndexOf("</");
  if (spliceAt < 0) return `<w:drawing>${anchorXml}${graphicXml}</w:drawing>`;
  return `<w:drawing>${anchorXml.slice(0, spliceAt)}${graphicXml}${anchorXml.slice(spliceAt)}</w:drawing>`;
}
