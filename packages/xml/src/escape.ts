const XML_CHAR_MAP: Record<string, string> = {
    "&": "&amp;",
    '"': "&quot;",
    "'": "&apos;",
    "<": "&lt;",
    ">": "&gt;",
};

const XML_CHAR_PATTERN = /([&"<>'])/g;

/** Escape text content for XML */
export function escapeXml(str: string): string {
    return str.replace(XML_CHAR_PATTERN, (ch) => XML_CHAR_MAP[ch]);
}

/**
 * Escape attribute value matching xml-js's js2xml behavior.
 * Handles already-escaped entities to prevent double-escaping.
 */
export function escapeAttributeValue(str: string): string {
    return String(str)
        .replace(/&(?!amp;|lt;|gt;|quot;|apos;)/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}
