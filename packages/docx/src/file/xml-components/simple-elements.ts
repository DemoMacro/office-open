/**
 * Docx-specific simple element helpers.
 *
 * @module
 */
import { BuilderElement } from "@office-open/core";
import type { XmlComponent } from "@office-open/core";

/**
 * Creates a WordprocessingML string value element with a `w:val` attribute.
 *
 * @param name - Element name (e.g., "w:pStyle")
 * @param value - String value for the `w:val` attribute
 */
export const createStringElement = (name: string, value: string): XmlComponent =>
    new BuilderElement({
        attributes: {
            value: { key: "w:val", value },
        },
        name,
    });
