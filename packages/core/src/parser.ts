import type { Element } from "@office-open/xml";

/**
 * Wrapper for unknown/unhandled XML elements in parsed output.
 * Enables lossless round-trip: unknown elements are preserved as Element trees
 * and can be serialized back via js2xml() during reconstruction.
 */
export interface RawElement {
    $raw: true;
    element: Element;
}

/** Type guard for RawElement */
export function isRaw(el: unknown): el is RawElement {
    return typeof el === "object" && el !== null && (el as RawElement).$raw === true;
}

/**
 * Mixed content array: typed children interleaved with unknown raw elements,
 * preserving original document order for lossless round-trip.
 */
export type MixedChildren<T> = Array<T | RawElement>;
