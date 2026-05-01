/**
 * Field module for WordprocessingML documents.
 *
 * This module provides support for complex fields, which are regions of text
 * that can contain dynamic content such as page numbers, dates, or mail merge fields.
 * Fields are delimited by field character elements (begin, separate, end).
 *
 * Reference: http://officeopenxml.com/WPrun.php
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import type { FormFieldOptions } from "./form-field";
import { createFormFieldData } from "./form-field";

/**
 * Field character types that delimit field regions.
 *
 * @internal
 */
const FieldCharacterType = {
    BEGIN: "begin",
    END: "end",
    SEPARATE: "separate",
} as const;

interface IFieldCharAttributes {
    readonly type: (typeof FieldCharacterType)[keyof typeof FieldCharacterType];
    readonly dirty?: boolean;
}

/**
 * Creates a field character element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_FldChar">
 *   <xsd:sequence>
 *     <xsd:element name="fldData" type="CT_Text" minOccurs="0"/>
 *     <xsd:element name="ffData" type="CT_FFData" minOccurs="0"/>
 *     <xsd:element name="numberingChange" type="CT_TrackChangeNumbering" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="fldCharType" type="ST_FldCharType" use="required"/>
 *   <xsd:attribute name="fldLock" type="s:ST_OnOff"/>
 *   <xsd:attribute name="dirty" type="s:ST_OnOff"/>
 * </xsd:complexType>
 * ```
 * @internal
 */
const createFieldChar = (
    type: (typeof FieldCharacterType)[keyof typeof FieldCharacterType],
    dirty?: boolean,
    ffData?: XmlComponent,
): XmlComponent => {
    const children: XmlComponent[] = [];
    if (ffData) {
        children.push(ffData);
    }
    return new BuilderElement<IFieldCharAttributes>({
        attributes: {
            dirty: { key: "w:dirty", value: dirty },
            type: { key: "w:fldCharType", value: type },
        },
        children: children.length > 0 ? children : undefined,
        name: "w:fldChar",
    });
};

/**
 * Creates the beginning of a complex field.
 *
 * The Begin element marks the start of a field. A field consists of a begin character,
 * field instructions, an optional separate character, field result, and an end character.
 *
 * For form fields, pass `formField` to embed `w:ffData` within the begin `w:fldChar`.
 *
 * @param dirty - Whether the field should be recalculated
 * @param formField - Optional form field data to embed in the begin character
 *
 * @example
 * ```typescript
 * // Simple field begin
 * createBegin();
 *
 * // Form field (checkbox)
 * createBegin(false, {
 *   name: "Check1",
 *   checkBox: { checked: true, sizeAuto: true },
 * });
 * ```
 */
export const createBegin = (dirty?: boolean, formField?: FormFieldOptions): XmlComponent =>
    createFieldChar(
        FieldCharacterType.BEGIN,
        dirty,
        formField ? createFormFieldData(formField) : undefined,
    );

/**
 * Creates the separator between field code and field result in a complex field.
 *
 * The Separate element divides the field code (instructions) from the field result
 * (the computed value).
 */
export const createSeparate = (dirty?: boolean): XmlComponent =>
    createFieldChar(FieldCharacterType.SEPARATE, dirty);

/**
 * Creates the end of a complex field.
 *
 * The End element marks the end of a field. Every field that begins with a Begin
 * element must be terminated with an End element.
 */
export const createEnd = (dirty?: boolean): XmlComponent =>
    createFieldChar(FieldCharacterType.END, dirty);
