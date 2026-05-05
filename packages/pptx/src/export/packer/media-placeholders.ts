import { escapeRegex, formatId } from "@office-open/core";

export function replaceMediaPlaceholders(
    xml: string,
    mediaData: readonly { readonly fileName: string }[],
    offset: number,
): string {
    let result = xml;
    mediaData.forEach((m, i) => {
        result = result.replace(
            new RegExp(`\\{media:${escapeRegex(m.fileName)}\\}`, "g"),
            formatId(offset, i, "rId"),
        );
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
        result = result.replace(
            new RegExp(`\\{video:${escapeRegex(m.fileName)}\\}`, "g"),
            formatId(offset, i, "rId"),
        );
    });
    return result;
}

export function getMediaRefs(
    xml: string,
    mediaArray: readonly { readonly fileName: string }[],
): readonly { readonly fileName: string }[] {
    return collectRefs(xml, "media:", mediaArray);
}

export function getVideoRefs(
    xml: string,
    mediaArray: readonly { readonly fileName: string }[],
): readonly { readonly fileName: string }[] {
    return collectRefs(xml, "video:", mediaArray);
}

function collectRefs(
    xml: string,
    prefix: string,
    mediaArray: readonly { readonly fileName: string }[],
): readonly { readonly fileName: string }[] {
    const pattern = new RegExp(`\\{${escapeRegex(prefix)}([^}]+)\\}`, "g");
    const keys = new Set<string>();
    for (const match of xml.matchAll(pattern)) {
        keys.add(match[1]);
    }
    return mediaArray.filter((m) => keys.has(m.fileName));
}
