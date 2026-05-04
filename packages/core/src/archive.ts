import { xml2js, attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { unzipSync, strFromU8 } from "fflate";

const XML_PARSE_OPTIONS = {
    nativeTypeAttributes: true,
    captureSpacesBetweenElements: true,
};

/**
 * Unzip an OOXML file (.docx, .pptx) into a Map of path → Uint8Array.
 */
export function unzipToMap(data: Uint8Array): Map<string, Uint8Array> {
    const entries = unzipSync(data);
    const map = new Map<string, Uint8Array>();
    for (const [path, bytes] of Object.entries(entries)) {
        map.set(path, bytes);
    }
    return map;
}

/**
 * Read a file from the zip as a UTF-8 string.
 */
export function readTextFromZip(zip: Map<string, Uint8Array>, path: string): string | undefined {
    const data = zip.get(path);
    if (data === undefined) return undefined;
    return strFromU8(data);
}

/**
 * Parse an XML file from the zip into an Element tree.
 */
export function readXmlFromZip(zip: Map<string, Uint8Array>, path: string): Element | undefined {
    const text = readTextFromZip(zip, path);
    if (text === undefined) return undefined;
    const wrapper = xml2js(text, XML_PARSE_OPTIONS) as Element;
    return wrapper.elements?.find((e) => e.type === "element");
}

/**
 * Read a binary file from the zip.
 */
export function readBinaryFromZip(
    zip: Map<string, Uint8Array>,
    path: string,
): Uint8Array | undefined {
    return zip.get(path);
}

/**
 * List all files in the zip matching a prefix.
 */
export function listFiles(zip: Map<string, Uint8Array>, prefix: string): string[] {
    const result: string[] = [];
    for (const path of zip.keys()) {
        if (path.startsWith(prefix)) {
            result.push(path);
        }
    }
    return result;
}

/**
 * Convert Uint8Array to base64 string.
 */
export function uint8ToBase64(data: Uint8Array): string {
    let binary = "";
    for (let i = 0; i < data.length; i++) {
        binary += String.fromCharCode(data[i]);
    }
    return btoa(binary);
}

/**
 * Determine image type from file extension.
 */
export function getImageType(fileName: string): string {
    const ext = fileName.split(".").pop()?.toLowerCase() ?? "";
    if (
        ["png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "ico", "emf", "wmf", "svg"].includes(
            ext,
        )
    ) {
        return ext === "jpeg" ? "jpg" : ext;
    }
    return "png";
}

// ── Relationship parsing ──

export interface Relationship {
    id: string;
    target: string;
    type: string;
    targetMode?: string;
}

export function parseRels(zip: Map<string, Uint8Array>, path: string): Relationship[] {
    const xml = readXmlFromZip(zip, path);
    if (!xml) return [];

    const result: Relationship[] = [];
    for (const rel of xml.elements ?? []) {
        if (rel.name !== "Relationship") continue;
        const id = attr(rel, "Id");
        const target = attr(rel, "Target");
        const type = attr(rel, "Type");
        const targetMode = attr(rel, "TargetMode");
        if (id && target) {
            result.push({ id, target, type: type ?? "", ...(targetMode ? { targetMode } : {}) });
        }
    }
    return result;
}

export function findRel(rels: Relationship[], id: string): Relationship | undefined {
    return rels.find((r) => r.id === id);
}

export function findRelsByType(rels: Relationship[], typeSubstring: string): Relationship[] {
    return rels.filter((r) => r.type.includes(typeSubstring));
}
