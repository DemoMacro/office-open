/**
 * Placeholder detection and replacement utilities for compiler post-processing.
 *
 * Both DOCX and PPTX compilers use placeholder tokens (e.g. `{chart:key}`)
 * in XML strings to defer relationship ID assignment. This module provides
 * fast pre-check utilities and shared replacement functions.
 */

export type IdFormat = "rId" | "plain";

export const formatId = (offset: number, index: number, style: IdFormat): string =>
  style === "rId" ? `rId${offset + index}` : `${offset + index}`;

/** Single-pass regex replacement for `{key}` placeholders using a Map lookup. */
const PLACEHOLDER_RE = /\{([^{}]+)\}/g;

function replacePlaceholders(xml: string, map: Map<string, string>): string {
  const parts: string[] = [];
  let last = 0;
  for (const m of xml.matchAll(PLACEHOLDER_RE)) {
    const replacement = map.get(m[1]);
    if (replacement !== undefined) {
      parts.push(xml.substring(last, m.index!), replacement);
      last = m.index! + m[0].length;
    }
  }
  if (last === 0) return xml;
  parts.push(xml.substring(last));
  return parts.join("");
}

export function hasPlaceholders(xml: string): boolean {
  return xml.includes("{");
}

export function collectPlaceholderKeys(xml: string, prefix: string): string[] {
  const keys = new Set<string>();
  const search = `{${prefix}`;
  let pos = 0;

  while ((pos = xml.indexOf(search, pos)) !== -1) {
    const keyStart = pos + search.length;
    const keyEnd = xml.indexOf("}", keyStart);
    if (keyEnd === -1) break;

    const key = xml.substring(keyStart, keyEnd);
    if (key.length > 0) keys.add(key);
    pos = keyEnd + 1;
  }

  return [...keys];
}

export function replaceImagePlaceholders(
  xml: string,
  mediaData: { fileName: string }[],
  offset: number,
  idFormat: IdFormat = "rId",
): string {
  if (mediaData.length === 0) return xml;
  const map = new Map<string, string>();
  for (let i = 0; i < mediaData.length; i++) {
    map.set(mediaData[i].fileName, formatId(offset, i, idFormat));
  }
  return replacePlaceholders(xml, map);
}

export function getReferencedMedia(
  xml: string,
  mediaArray: { fileName: string }[],
): { fileName: string }[] {
  return mediaArray.filter((image) => xml.includes(`{${image.fileName}}`));
}

/**
 * Combined find + replace for image placeholders in a single pass.
 *
 * Scans XML for `{fileName}` placeholders, replaces them with formatted IDs,
 * and returns the list of referenced media items for relationship registration.
 * More efficient than calling getReferencedMedia() + replaceImagePlaceholders() separately.
 */
export function findAndReplaceImagePlaceholders(
  xml: string,
  mediaArray: { fileName: string }[],
  offset: number,
  idFormat: IdFormat = "rId",
): { xml: string; referenced: { fileName: string }[] } {
  if (mediaArray.length === 0 || !hasPlaceholders(xml)) {
    return { xml, referenced: [] };
  }
  // Build fileName → { index, mediaItem } lookup
  const indexMap = new Map<string, { idx: number; item: { fileName: string } }>();
  for (let i = 0; i < mediaArray.length; i++) {
    indexMap.set(mediaArray[i].fileName, { idx: i, item: mediaArray[i] });
  }

  const referenced: { fileName: string }[] = [];
  const replaceMap = new Map<string, string>();
  const parts: string[] = [];
  let last = 0;

  for (const m of xml.matchAll(PLACEHOLDER_RE)) {
    const key = m[1];
    const entry = indexMap.get(key);
    if (entry !== undefined) {
      if (!replaceMap.has(key)) {
        replaceMap.set(key, formatId(offset, entry.idx, idFormat));
        referenced.push(entry.item);
      }
      parts.push(xml.substring(last, m.index!), replaceMap.get(key)!);
      last = m.index! + m[0].length;
    }
  }

  if (last === 0) return { xml, referenced: [] };
  parts.push(xml.substring(last));
  return { xml: parts.join(""), referenced };
}

/**
 * Replace all placeholder types in a single pass.
 *
 * Combines image, chart, smartart, and arbitrary key-value replacements
 * into one regex scan of the XML string. Used by compilers to avoid
 * multiple full-string passes on large documents.
 */
export function replaceAllPlaceholders(
  xml: string,
  entries: Array<{ prefix?: string; key: string; value: string }>,
): string {
  if (entries.length === 0 || !hasPlaceholders(xml)) return xml;
  const map = new Map<string, string>();
  for (const { prefix, key, value } of entries) {
    map.set(prefix ? `${prefix}${key}` : key, value);
  }
  return replacePlaceholders(xml, map);
}

export function replaceChartPlaceholders(
  xml: string,
  chartKeys: string[],
  offset: number,
  idFormat: IdFormat = "rId",
): string {
  if (chartKeys.length === 0) return xml;
  const map = new Map<string, string>();
  for (let i = 0; i < chartKeys.length; i++) {
    map.set(`chart:${chartKeys[i]}`, formatId(offset, i, idFormat));
  }
  return replacePlaceholders(xml, map);
}

import type { RelationshipType } from "../opc/relationships";

export interface SmartArtRelOptions {
  pathPrefix: string;
  styleRelType: RelationshipType;
}

export function replaceSmartArtPlaceholders(
  xml: string,
  keys: string[],
  dataOffset: number,
  idFormat: IdFormat = "rId",
): string {
  if (keys.length === 0) return xml;
  const count = keys.length;
  const loOffset = dataOffset + count;
  const qsOffset = loOffset + count;
  const csOffset = qsOffset + count;

  const map = new Map<string, string>();
  for (let i = 0; i < count; i++) {
    map.set(`smartart:${keys[i]}`, formatId(dataOffset, i, idFormat));
    map.set(`smartart-lo:${keys[i]}`, formatId(loOffset, i, idFormat));
    map.set(`smartart-qs:${keys[i]}`, formatId(qsOffset, i, idFormat));
    map.set(`smartart-cs:${keys[i]}`, formatId(csOffset, i, idFormat));
  }
  return replacePlaceholders(xml, map);
}

export function addSmartArtRelationships(
  keys: string[],
  addRel: (id: number, type: RelationshipType, target: string, targetMode?: string) => void,
  baseOffset: number,
  globalStartIndex: number,
  options: SmartArtRelOptions,
): void {
  const { pathPrefix, styleRelType } = options;
  const count = keys.length;
  const loOffset = baseOffset + count;
  const qsOffset = loOffset + count;
  const csOffset = qsOffset + count;

  for (let i = 0; i < keys.length; i++) {
    const gi = globalStartIndex + i;
    addRel(
      baseOffset + i,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
      `${pathPrefix}diagrams/data${gi + 1}.xml`,
    );
    addRel(
      loOffset + i,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
      `${pathPrefix}diagrams/layout${gi + 1}.xml`,
    );
    addRel(qsOffset + i, styleRelType, `${pathPrefix}diagrams/quickStyle${gi + 1}.xml`);
    addRel(
      csOffset + i,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
      `${pathPrefix}diagrams/colors${gi + 1}.xml`,
    );
    const drOffset = csOffset + count;
    addRel(
      drOffset + i,
      "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
      `${pathPrefix}diagrams/drawing${gi + 1}.xml`,
    );
  }
}

// ── Numbering placeholders (DOCX) ──

/** Replace `{reference-instance}` placeholders with numId values. */
export function replaceNumberingPlaceholders(
  xml: string,
  concreteNumberings: readonly { reference: string; instance: number; numId: number }[],
): string {
  if (concreteNumberings.length === 0) return xml;
  const map = new Map<string, string>();
  for (const { reference, instance, numId } of concreteNumberings) {
    map.set(`${reference}-${instance}`, numId.toString());
  }
  return replacePlaceholders(xml, map);
}

// ── Media/Video placeholders (PPTX) ──

/** Replace `{media:fileName}` placeholders with relationship IDs. */
export function replaceMediaPlaceholders(
  xml: string,
  mediaData: { fileName: string }[],
  offset: number,
): string {
  if (mediaData.length === 0) return xml;
  const map = new Map<string, string>();
  for (let i = 0; i < mediaData.length; i++) {
    map.set(`media:${mediaData[i].fileName}`, formatId(offset, i, "rId"));
  }
  return replacePlaceholders(xml, map);
}

/** Replace `{video:fileName}` placeholders with relationship IDs. */
export function replaceVideoPlaceholders(
  xml: string,
  mediaData: { fileName: string }[],
  offset: number,
): string {
  if (mediaData.length === 0) return xml;
  const map = new Map<string, string>();
  for (let i = 0; i < mediaData.length; i++) {
    map.set(`video:${mediaData[i].fileName}`, formatId(offset, i, "rId"));
  }
  return replacePlaceholders(xml, map);
}

/** Collect media items referenced by `{media:...}` placeholders in XML. */
export function getMediaRefs(
  xml: string,
  mediaArray: { fileName: string }[],
): { fileName: string }[] {
  return collectPrefixedRefs(xml, "media:", mediaArray);
}

/** Collect media items referenced by `{video:...}` placeholders in XML. */
export function getVideoRefs(
  xml: string,
  mediaArray: { fileName: string }[],
): { fileName: string }[] {
  return collectPrefixedRefs(xml, "video:", mediaArray);
}

function collectPrefixedRefs(
  xml: string,
  prefix: string,
  mediaArray: { fileName: string }[],
): { fileName: string }[] {
  const keys = new Set<string>();
  const search = `{${prefix}`;
  let pos = 0;
  while ((pos = xml.indexOf(search, pos)) !== -1) {
    const keyStart = pos + search.length;
    const keyEnd = xml.indexOf("}", keyStart);
    if (keyEnd === -1) break;
    keys.add(xml.substring(keyStart, keyEnd));
    pos = keyEnd + 1;
  }
  return mediaArray.filter((m) => keys.has(m.fileName));
}

// ── Hyperlink placeholders (PPTX) ──

/** Replace `{hlink:key}` placeholders with relationship IDs. */
export function replaceHyperlinkPlaceholders(
  xml: string,
  hyperlinks: { key: string }[],
  offset: number,
): string {
  if (hyperlinks.length === 0) return xml;
  const map = new Map<string, string>();
  for (let i = 0; i < hyperlinks.length; i++) {
    map.set(`hlink:${hyperlinks[i].key}`, formatId(offset, i, "rId"));
  }
  return replacePlaceholders(xml, map);
}
