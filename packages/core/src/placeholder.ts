/**
 * Placeholder detection utilities for compiler post-processing optimization.
 *
 * Both DOCX and PPTX compilers use placeholder tokens (e.g. `{chart:key}`)
 * in XML strings to defer relationship ID assignment. This module provides
 * fast pre-check utilities to avoid unnecessary regex scans when no placeholders exist.
 */

/**
 * Check if an XML string contains any placeholder tokens.
 * Uses a simple `{` character check which is O(n) but much cheaper
 * than running multiple regex scans.
 */
export function hasPlaceholders(xml: string): boolean {
    return xml.includes("{");
}

/**
 * Collect all unique placeholder keys with a given prefix from an XML string.
 * Single-pass scan using indexOf, avoiding RegExp overhead.
 *
 * @param xml - The XML string to scan
 * @param prefix - The placeholder prefix (e.g. "chart:", "smartart:", "hlink:")
 * @returns Array of unique keys found after the prefix
 */
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
