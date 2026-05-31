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

export function hasPlaceholders(xml: string): boolean {
  return xml.includes("{");
}

export function collectPlaceholderKeys(xml: string, prefix: string): string[] {
  const keys: string[] = [];
  const search = `{${prefix}`;
  let pos = 0;

  while ((pos = xml.indexOf(search, pos)) !== -1) {
    const keyStart = pos + search.length;
    const keyEnd = xml.indexOf("}", keyStart);
    if (keyEnd === -1) break;

    const key = xml.substring(keyStart, keyEnd);
    if (key.length > 0 && !keys.includes(key)) {
      keys.push(key);
    }
    pos = keyEnd + 1;
  }

  return keys;
}

export function replaceImagePlaceholders(
  xml: string,
  mediaData: readonly { readonly fileName: string }[],
  offset: number,
  idFormat: IdFormat = "rId",
): string {
  let result = xml;
  for (let i = 0; i < mediaData.length; i++) {
    result = result.replaceAll(`{${mediaData[i].fileName}}`, formatId(offset, i, idFormat));
  }
  return result;
}

export function getReferencedMedia(
  xml: string,
  mediaArray: readonly { readonly fileName: string }[],
): readonly { readonly fileName: string }[] {
  return mediaArray.filter((image) => xml.includes(`{${image.fileName}}`));
}

export function replaceChartPlaceholders(
  xml: string,
  chartKeys: readonly string[],
  offset: number,
  idFormat: IdFormat = "rId",
): string {
  let result = xml;
  for (let i = 0; i < chartKeys.length; i++) {
    result = result.replaceAll(`{chart:${chartKeys[i]}}`, formatId(offset, i, idFormat));
  }
  return result;
}

import type { RelationshipType } from "./opc/relationships";

export interface SmartArtRelOptions {
  readonly pathPrefix: string;
  readonly styleRelType: RelationshipType;
}

export function replaceSmartArtPlaceholders(
  xml: string,
  keys: readonly string[],
  dataOffset: number,
  idFormat: IdFormat = "rId",
): string {
  let result = xml;
  const count = keys.length;
  const loOffset = dataOffset + count;
  const qsOffset = loOffset + count;
  const csOffset = qsOffset + count;

  for (let i = 0; i < keys.length; i++) {
    result = result.replaceAll(`{smartart:${keys[i]}}`, formatId(dataOffset, i, idFormat));
    result = result.replaceAll(`{smartart-lo:${keys[i]}}`, formatId(loOffset, i, idFormat));
    result = result.replaceAll(`{smartart-qs:${keys[i]}}`, formatId(qsOffset, i, idFormat));
    result = result.replaceAll(`{smartart-cs:${keys[i]}}`, formatId(csOffset, i, idFormat));
  }

  return result;
}

export function addSmartArtRelationships(
  keys: readonly string[],
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
