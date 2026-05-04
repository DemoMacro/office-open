import type { Element } from "./types";

/**
 * Find the first direct child element with the given name.
 */
export function findChild(parent: Element | undefined, name: string): Element | undefined {
    return parent?.elements?.find((e) => e.name === name);
}

/**
 * Get all direct child elements matching the given name.
 */
export function children(parent: Element | undefined, name: string): Element[] {
    return parent?.elements?.filter((e) => e.name === name) ?? [];
}

/**
 * Get all direct child elements.
 */
export function allChildren(parent: Element | undefined): Element[] {
    return parent?.elements ?? [];
}

/**
 * Get text content of the first child element with the given name.
 */
export function childText(parent: Element | undefined, name: string): string {
    const child = findChild(parent, name);
    return textOf(child);
}

/**
 * Get text content of an element.
 * Handles cases where text may be directly on .text or in a child element.
 */
export function textOf(element: Element | undefined): string {
    if (!element) return "";
    if (element.text !== undefined && typeof element.text === "string") return element.text;
    if (element.elements && element.elements.length > 0) {
        return element.elements.map((e) => (typeof e.text === "string" ? e.text : "")).join("");
    }
    return "";
}

/**
 * Collect text from all direct text nodes within an element.
 */
export function collectText(element: Element | undefined): string {
    if (!element) return "";
    const parts: string[] = [];
    collectTextRecursive(element, parts);
    return parts.join("");
}

function collectTextRecursive(element: Element | undefined, parts: string[]): void {
    if (!element) return;
    if (element.text !== undefined && typeof element.text === "string") {
        parts.push(element.text);
    }
    if (element.elements) {
        for (const child of element.elements) {
            collectTextRecursive(child, parts);
        }
    }
}

/**
 * Get an attribute value as a string.
 */
export function attr(element: Element | undefined, name: string): string | undefined {
    const v = element?.attributes?.[name];
    return v !== undefined ? String(v) : undefined;
}

/**
 * Get an attribute value as a number.
 */
export function attrNum(element: Element | undefined, name: string): number | undefined {
    const v = element?.attributes?.[name];
    if (v === undefined) return undefined;
    const n = Number(v);
    return isNaN(n) ? undefined : n;
}

/**
 * Get an attribute value as a boolean.
 */
export function attrBool(element: Element | undefined, name: string): boolean | undefined {
    const v = element?.attributes?.[name];
    if (v === undefined) return undefined;
    if (typeof v === "boolean") return v;
    const lower = String(v).toLowerCase();
    if (lower === "true" || lower === "1") return true;
    if (lower === "false" || lower === "0") return false;
    return undefined;
}

/**
 * Check if an element has a specific child element.
 */
export function hasChild(parent: Element | undefined, name: string): boolean {
    return parent?.elements?.some((e) => e.name === name) ?? false;
}

/**
 * Find deep descendant elements matching the given name.
 */
export function findDeep(parent: Element | undefined, name: string): Element[] {
    const result: Element[] = [];
    if (!parent) return result;
    for (const child of parent.elements ?? []) {
        if (child.name === name) result.push(child);
        result.push(...findDeep(child, name));
    }
    return result;
}

/**
 * Get the number of direct child elements.
 */
export function childCount(parent: Element | undefined): number {
    return parent?.elements?.length ?? 0;
}
