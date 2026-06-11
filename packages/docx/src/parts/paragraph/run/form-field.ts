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
import { element } from "@office-open/xml";

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
  size?: number;
  /**
   * Whether to auto-size the checkbox to match surrounding text.
   *
   * Mutually exclusive with `size`.
   */
  sizeAuto?: boolean;
  /** Default checked state */
  default?: boolean;
  /** Current checked state */
  checked?: boolean;
}

/**
 * Options for a dropdown list form field.
 */
export interface DropDownListOptions {
  /** Items in the dropdown list */
  entries: string[];
  /** Index of the currently selected entry */
  result?: number;
  /** Index of the default entry */
  default?: number;
}

/**
 * Options for a text input form field.
 */
export interface TextInputOptions {
  /** Text input type */
  type?: (typeof FormFieldTextType)[keyof typeof FormFieldTextType];
  /** Default text value */
  default?: string;
  /** Maximum character length */
  maxLength?: number;
  /** Format string */
  format?: string;
}

/**
 * Help or status text options.
 */
export interface FormFieldTextOptions {
  /** Text type: "text" for custom text, "autoText" for auto text reference */
  type: "text" | "autoText";
  /** The text value */
  value: string;
}

/**
 * Common options for all form field types.
 */
export interface FormFieldCommonOptions {
  /** Field name (max 65 characters) */
  name?: string;
  /** Numeric label */
  label?: number;
  /** Tab order index */
  tabIndex?: number;
  /** Whether the field is enabled */
  enabled?: boolean;
  /** Recalculate fields when exiting */
  calcOnExit?: boolean;
  /** Entry macro name */
  entryMacro?: string;
  /** Exit macro name */
  exitMacro?: string;
  /** Help text shown when the user presses F1 */
  helpText?: FormFieldTextOptions;
  /** Status bar text */
  statusText?: FormFieldTextOptions;
}

/**
 * Options for a form field definition (CT_FFData).
 *
 * Exactly one of `checkBox`, `dropDownList`, or `textInput` must be specified.
 */
export interface FormFieldOptions extends FormFieldCommonOptions {
  /** Checkbox form field */
  checkBox?: CheckBoxOptions;
  /** Dropdown list form field */
  dropDownList?: DropDownListOptions;
  /** Text input form field */
  textInput?: TextInputOptions;
}

/** Build a w:val-style element with a single attribute. */
const valElement = (name: string, val: string | number): string => `<${name} w:val="${val}"/>`;

/**
 * Creates a help or status text element.
 */
const createFormFieldText = (name: string, options: FormFieldTextOptions): string =>
  `<${name} w:type="${options.type}" w:val="${options.value}"/>`;

/**
 * Creates a checkbox form field element (w:checkBox).
 */
const createCheckBox = (options: CheckBoxOptions): string => {
  const children: string[] = [];

  if (options.size !== undefined) {
    children.push(valElement("w:size", options.size));
  } else if (options.sizeAuto !== undefined) {
    children.push(`<w:sizeAuto/>`);
  }

  if (options.default !== undefined) {
    children.push(`<w:default/>`);
  }
  if (options.checked !== undefined) {
    children.push(`<w:checked/>`);
  }

  return element("w:checkBox", undefined, children);
};

/**
 * Creates a dropdown list form field element (w:ddList).
 */
const createDropDownList = (options: DropDownListOptions): string => {
  const children: string[] = [];

  if (options.result !== undefined) {
    children.push(valElement("w:result", options.result));
  }
  if (options.default !== undefined) {
    children.push(valElement("w:default", options.default));
  }
  for (const entry of options.entries) {
    children.push(valElement("w:listEntry", entry));
  }

  return element("w:ddList", undefined, children);
};

/**
 * Creates a text input form field element (w:textInput).
 */
const createTextInput = (options: TextInputOptions): string => {
  const children: string[] = [];

  if (options.type !== undefined) {
    children.push(valElement("w:type", options.type));
  }
  if (options.default !== undefined) {
    children.push(valElement("w:default", options.default));
  }
  if (options.maxLength !== undefined) {
    children.push(valElement("w:maxLength", options.maxLength));
  }
  if (options.format !== undefined) {
    children.push(valElement("w:format", options.format));
  }

  return element("w:textInput", undefined, children);
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
export const createFormFieldData = (options: FormFieldOptions): string => {
  const children: string[] = [];

  if (options.name !== undefined) {
    children.push(valElement("w:name", options.name));
  }
  if (options.label !== undefined) {
    children.push(valElement("w:label", options.label));
  }
  if (options.tabIndex !== undefined) {
    children.push(valElement("w:tabIndex", options.tabIndex));
  }
  if (options.enabled !== undefined) {
    children.push(`<w:enabled/>`);
  }
  if (options.calcOnExit !== undefined) {
    children.push(`<w:calcOnExit/>`);
  }
  if (options.entryMacro !== undefined) {
    children.push(valElement("w:entryMacro", options.entryMacro));
  }
  if (options.exitMacro !== undefined) {
    children.push(valElement("w:exitMacro", options.exitMacro));
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

  return element("w:ffData", undefined, children);
};
