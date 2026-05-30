import { formatId } from "@office-open/core";

export function replaceMediaPlaceholders(
  xml: string,
  mediaData: readonly { readonly fileName: string }[],
  offset: number,
): string {
  let result = xml;
  mediaData.forEach((m, i) => {
    result = result.replaceAll(`{media:${m.fileName}}`, formatId(offset, i, "rId"));
  });
  return result;
}

export function replaceVideoPlaceholders(
  xml: string,
  mediaData: readonly { readonly fileName: string }[],
  offset: number,
): string {
  let result = xml;
  mediaData.forEach((m, i) => {
    result = result.replaceAll(`{video:${m.fileName}}`, formatId(offset, i, "rId"));
  });
  return result;
}

export function getMediaRefs(
  xml: string,
  mediaArray: readonly { readonly fileName: string }[],
): readonly { readonly fileName: string }[] {
  return collectRefs(xml, "{media:", mediaArray);
}

export function getVideoRefs(
  xml: string,
  mediaArray: readonly { readonly fileName: string }[],
): readonly { readonly fileName: string }[] {
  return collectRefs(xml, "{video:", mediaArray);
}

function collectRefs(
  xml: string,
  search: string,
  mediaArray: readonly { readonly fileName: string }[],
): readonly { readonly fileName: string }[] {
  const keys = new Set<string>();
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
