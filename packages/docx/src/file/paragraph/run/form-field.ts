/**
 * Form field module for WordprocessingML documents.
 *
 * Form fields allow creating interactive controls (checkboxes, dropdown lists,
 * text inputs) within a document. Each form field is wrapped in a field code
 * region delimited by fldChar elements (begin/separate/end), with ffData
 * attached to the begin fldChar.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_FFData, CT_FFCheckBox, CT_FFDDList, CT_FFTextInput
 *
 * @module
 */
import { BuilderElement, OnOffElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Text input field type (ST_FFTextType).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_FFTextType">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="regular"/>
 *     <xsd:enumeration value="number"/>
 *     <xsd:enumeration value="date"/>
 *     <xsd:enumeration value="currentTime"/>
 *     <xsd:enumeration value="currentDate"/>
 *     <xsd:enumeration value="calculated"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const FormFieldTextType = {
    /** Regular text input */
    REGULAR: "regular",
    /** Numeric input */
    NUMBER: "number",
    /** Date input */
    DATE: "date",
    /** Current time */
    CURRENT_TIME: "currentTime",
    /** Current date */
    CURRENT_DATE: "currentDate",
    /** Calculated value */
    CALCULATED: "calculated",
} as const;

/**
 * Options for a checkbox form field.
 */
export interface CheckBoxOptions {
    /**
     * Checkbox size in half-points.
     *
     * Mutually exclusive with `sizeAuto`.
     */
    readonly size?: number;
    /**
     * Whether to auto-size the checkbox to match surrounding text.
     *
     * Mutually exclusive with `size`.
     */
    readonly sizeAuto?: boolean;
    /** Default checked state */
    readonly default?: boolean;
    /** Current checked state */
    readonly checked?: boolean;
}

/**
 * Options for a dropdown list form field.
 */
export interface DropDownListOptions {
    /** Items in the dropdown list */
    readonly entries: readonly string[];
    /** Index of the currently selected entry */
    readonly result?: number;
    /** Index of the default entry */
    readonly default?: number;
}

/**
 * Options for a text input form field.
 */
export interface TextInputOptions {
    /** Text input type */
    readonly type?: (typeof FormFieldTextType)[keyof typeof FormFieldTextType];
    /** Default text value */
    readonly default?: string;
    /** Maximum character length */
    readonly maxLength?: number;
    /** Format string */
    readonly format?: string;
}

/**
 * Help or status text options.
 */
export interface FormFieldTextOptions {
    /** Text type: "text" for custom text, "autoText" for auto text reference */
    readonly type: "text" | "autoText";
    /** The text value */
    readonly value: string;
}

/**
 * Common options for all form field types.
 */
export interface FormFieldCommonOptions {
    /** Field name (max 65 characters) */
    readonly name?: string;
    /** Numeric label */
    readonly label?: number;
    /** Tab order index */
    readonly tabIndex?: number;
    /** Whether the field is enabled */
    readonly enabled?: boolean;
    /** Recalculate fields when exiting */
    readonly calcOnExit?: boolean;
    /** Entry macro name */
    readonly entryMacro?: string;
    /** Exit macro name */
    readonly exitMacro?: string;
    /** Help text shown when the user presses F1 */
    readonly helpText?: FormFieldTextOptions;
    /** Status bar text */
    readonly statusText?: FormFieldTextOptions;
}

/**
 * Options for a form field definition (CT_FFData).
 *
 * Exactly one of `checkBox`, `dropDownList`, or `textInput` must be specified.
 */
export interface FormFieldOptions extends FormFieldCommonOptions {
    /** Checkbox form field */
    readonly checkBox?: CheckBoxOptions;
    /** Dropdown list form field */
    readonly dropDownList?: DropDownListOptions;
    /** Text input form field */
    readonly textInput?: TextInputOptions;
}

/**
 * Creates a help or status text element.
 */
const createFormFieldText = (name: string, options: FormFieldTextOptions): XmlComponent =>
    new BuilderElement<{ readonly type: string; readonly val: string }>({
        attributes: {
            type: { key: "w:type", value: options.type },
            val: { key: "w:val", value: options.value },
        },
        name,
    });

/**
 * Creates a checkbox form field element (w:checkBox).
 */
const createCheckBox = (options: CheckBoxOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.size !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: number }>({
                attributes: { val: { key: "w:val", value: options.size } },
                name: "w:size",
            }),
        );
    } else if (options.sizeAuto !== undefined) {
        children.push(new OnOffElement("w:sizeAuto", options.sizeAuto));
    }

    if (options.default !== undefined) {
        children.push(new OnOffElement("w:default", options.default));
    }
    if (options.checked !== undefined) {
        children.push(new OnOffElement("w:checked", options.checked));
    }

    return new BuilderElement({
        children,
        name: "w:checkBox",
    });
};

/**
 * Creates a dropdown list form field element (w:ddList).
 */
const createDropDownList = (options: DropDownListOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.result !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: number }>({
                attributes: { val: { key: "w:val", value: options.result } },
                name: "w:result",
            }),
        );
    }
    if (options.default !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: number }>({
                attributes: { val: { key: "w:val", value: options.default } },
                name: "w:default",
            }),
        );
    }
    for (const entry of options.entries) {
        children.push(
            new BuilderElement<{ readonly val: string }>({
                attributes: { val: { key: "w:val", value: entry } },
                name: "w:listEntry",
            }),
        );
    }

    return new BuilderElement({
        children,
        name: "w:ddList",
    });
};

/**
 * Creates a text input form field element (w:textInput).
 */
const createTextInput = (options: TextInputOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.type !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: string }>({
                attributes: { val: { key: "w:val", value: options.type } },
                name: "w:type",
            }),
        );
    }
    if (options.default !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: string }>({
                attributes: { val: { key: "w:val", value: options.default } },
                name: "w:default",
            }),
        );
    }
    if (options.maxLength !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: number }>({
                attributes: { val: { key: "w:val", value: options.maxLength } },
                name: "w:maxLength",
            }),
        );
    }
    if (options.format !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: string }>({
                attributes: { val: { key: "w:val", value: options.format } },
                name: "w:format",
            }),
        );
    }

    return new BuilderElement({
        children,
        name: "w:textInput",
    });
};

/**
 * Creates a form field data element (w:ffData).
 *
 * This element contains the definition for an interactive form field
 * (checkbox, dropdown list, or text input) within a document.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_FFData">
 *   <xsd:choice maxOccurs="unbounded">
 *     <xsd:element name="name" type="CT_FFName"/>
 *     <xsd:element name="label" type="CT_DecimalNumber"/>
 *     <xsd:element name="tabIndex" type="CT_UnsignedDecimalNumber"/>
 *     <xsd:element name="enabled" type="CT_OnOff"/>
 *     <xsd:element name="calcOnExit" type="CT_OnOff"/>
 *     <xsd:element name="entryMacro" type="CT_MacroName"/>
 *     <xsd:element name="exitMacro" type="CT_MacroName"/>
 *     <xsd:element name="helpText" type="CT_FFHelpText"/>
 *     <xsd:element name="statusText" type="CT_FFStatusText"/>
 *     <xsd:choice>
 *       <xsd:element name="checkBox" type="CT_FFCheckBox"/>
 *       <xsd:element name="ddList" type="CT_FFDDList"/>
 *       <xsd:element name="textInput" type="CT_FFTextInput"/>
 *     </xsd:choice>
 *   </xsd:choice>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Checkbox
 * createFormFieldData({
 *   name: "Check1",
 *   checkBox: { checked: true, sizeAuto: true },
 * });
 *
 * // Dropdown
 * createFormFieldData({
 *   name: "DropDown1",
 *   dropDownList: { entries: ["Option A", "Option B", "Option C"], result: 0 },
 * });
 *
 * // Text input
 * createFormFieldData({
 *   name: "Text1",
 *   textInput: { type: "regular", default: "Enter text here" },
 * });
 * ```
 */
export const createFormFieldData = (options: FormFieldOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.name !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: string }>({
                attributes: { val: { key: "w:val", value: options.name } },
                name: "w:name",
            }),
        );
    }
    if (options.label !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: number }>({
                attributes: { val: { key: "w:val", value: options.label } },
                name: "w:label",
            }),
        );
    }
    if (options.tabIndex !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: number }>({
                attributes: { val: { key: "w:val", value: options.tabIndex } },
                name: "w:tabIndex",
            }),
        );
    }
    if (options.enabled !== undefined) {
        children.push(new OnOffElement("w:enabled", options.enabled));
    }
    if (options.calcOnExit !== undefined) {
        children.push(new OnOffElement("w:calcOnExit", options.calcOnExit));
    }
    if (options.entryMacro !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: string }>({
                attributes: { val: { key: "w:val", value: options.entryMacro } },
                name: "w:entryMacro",
            }),
        );
    }
    if (options.exitMacro !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: string }>({
                attributes: { val: { key: "w:val", value: options.exitMacro } },
                name: "w:exitMacro",
            }),
        );
    }
    if (options.helpText) {
        children.push(createFormFieldText("w:helpText", options.helpText));
    }
    if (options.statusText) {
        children.push(createFormFieldText("w:statusText", options.statusText));
    }

    // Exactly one of the three mutually exclusive form field types
    if (options.checkBox) {
        children.push(createCheckBox(options.checkBox));
    } else if (options.dropDownList) {
        children.push(createDropDownList(options.dropDownList));
    } else if (options.textInput) {
        children.push(createTextInput(options.textInput));
    }

    return new BuilderElement({
        children,
        name: "w:ffData",
    });
};
