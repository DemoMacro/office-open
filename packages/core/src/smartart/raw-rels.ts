/**
 * Target rewriting for verbatim SmartArt data-part relationships.
 *
 * @module
 */

import { attr, parse, stringify } from "@office-open/xml";

import type { DataType } from "../util/data-type";
import { toUint8Array } from "../util/data-type";

function decodeRels(rels: DataType): { isText: boolean; text: string } {
  return {
    isText: typeof rels === "string",
    text: typeof rels === "string" ? rels : new TextDecoder().decode(toUint8Array(rels)),
  };
}

function rewriteMediaTargets(text: string, renames: ReadonlyMap<string, string>): string {
  let output = "";
  let last = 0;
  for (const match of text.matchAll(/\.\.\/media\/([^"']*)["']/g)) {
    const name = match[1] ?? "";
    const renamed = renames.get(name);
    if (renamed === undefined) continue;
    const start = (match.index ?? 0) + "../media/".length;
    output += text.slice(last, start) + renamed;
    last = start + name.length;
  }
  return output + text.slice(last);
}

/** Re-point raw data rels after companion media is deduplicated under a new name. */
export function remapSmartArtMediaTargets(
  rels: DataType,
  renames: ReadonlyMap<string, string>,
): DataType {
  if (renames.size === 0) return rels;
  const { isText, text } = decodeRels(rels);
  const remapped = rewriteMediaTargets(text, renames);
  return isText ? remapped : new TextEncoder().encode(remapped);
}

/** Re-point the diagramDrawing relationship when output SmartArt parts are renumbered. */
export function remapSmartArtDrawingTarget(rels: DataType, drawingFileName: string): DataType {
  const { isText, text } = decodeRels(rels);
  const relsEl = parse(text);
  let changed = false;
  const relationships = relsEl.elements?.[0]?.elements ?? [];
  for (const relationship of relationships) {
    if (relationship.name !== "Relationship") continue;
    if (!(attr(relationship, "Type") ?? "").endsWith("/diagramDrawing")) continue;
    const target = attr(relationship, "Target");
    if (!target) continue;
    const separator = Math.max(target.lastIndexOf("/"), target.lastIndexOf("\\"));
    const remapped =
      separator === -1 ? drawingFileName : target.slice(0, separator + 1) + drawingFileName;
    if (remapped === target) continue;
    relationship.attributes = { ...relationship.attributes, Target: remapped };
    changed = true;
  }
  if (!changed) return rels;
  const remapped = stringify(relsEl);
  return isText ? remapped : new TextEncoder().encode(remapped);
}
