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
import {
  attr,
  attrBool,
  attrNum,
  children as xmlChildren,
  element,
  findChild,
} from "@office-open/xml";
import type { Element } from "@office-open/xml";

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

  // <w:default> is the reset/initial state Word renders the box in. Word will
  // not render the box without it, so derive it from `checked` when omitted.
  const defaultVal = options.default ?? options.checked;
  if (defaultVal !== undefined) {
    children.push(defaultVal ? `<w:default/>` : `<w:default w:val="0"/>`);
  }
  if (options.checked !== undefined) {
    children.push(options.checked ? `<w:checked/>` : `<w:checked w:val="0"/>`);
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
  // Word requires an explicit <w:enabled/> to render the field; default true.
  // (Previously emitted a val-less tag even for enabled:false, which is "true".)
  const enabled = options.enabled ?? true;
  children.push(enabled ? `<w:enabled/>` : `<w:enabled w:val="0"/>`);
  // <w:calcOnExit> defaults to false (Word's standard).
  const calcOnExit = options.calcOnExit ?? false;
  children.push(calcOnExit ? `<w:calcOnExit/>` : `<w:calcOnExit w:val="0"/>`);
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

/**
 * Parse a w:ffData element back into FormFieldOptions.
 *
 * Inverse of {@link createFormFieldData}. Reads the common form-field metadata
 * (name, label, tabIndex, enabled, calcOnExit) and exactly one of
 * checkBox / dropDownList / textInput.
 */
export function parseFormFieldData(el: Element): FormFieldOptions {
  const opts: Partial<FormFieldOptions> = {};

  const name = findChild(el, "w:name");
  if (name) opts.name = attr(name, "w:val");
  const label = findChild(el, "w:label");
  if (label) {
    const v = attrNum(label, "w:val");
    if (v !== undefined) opts.label = v;
  }
  const tabIndex = findChild(el, "w:tabIndex");
  if (tabIndex) {
    const v = attrNum(tabIndex, "w:val");
    if (v !== undefined) opts.tabIndex = v;
  }
  const enabled = findChild(el, "w:enabled");
  if (enabled) opts.enabled = attrBool(enabled, "w:val") ?? true;
  const calcOnExit = findChild(el, "w:calcOnExit");
  if (calcOnExit) opts.calcOnExit = attrBool(calcOnExit, "w:val") ?? true;

  const checkBox = findChild(el, "w:checkBox");
  if (checkBox) {
    const cb: Partial<CheckBoxOptions> = {};
    if (findChild(checkBox, "w:sizeAuto")) cb.sizeAuto = true;
    const size = findChild(checkBox, "w:size");
    if (size) {
      const v = attrNum(size, "w:val");
      if (v !== undefined) cb.size = v;
    }
    const def = findChild(checkBox, "w:default");
    if (def) cb.default = attrBool(def, "w:val") ?? true;
    const checked = findChild(checkBox, "w:checked");
    if (checked) cb.checked = attrBool(checked, "w:val") ?? true;
    opts.checkBox = cb as CheckBoxOptions;
  } else {
    const ddList = findChild(el, "w:ddList");
    if (ddList) {
      const entries: string[] = [];
      for (const li of xmlChildren(ddList, "w:listEntry")) {
        entries.push(attr(li, "w:val") ?? "");
      }
      const ddl: Partial<DropDownListOptions> = { entries };
      const result = findChild(ddList, "w:result");
      if (result) {
        const v = attrNum(result, "w:val");
        if (v !== undefined) ddl.result = v;
      }
      const def = findChild(ddList, "w:default");
      if (def) {
        const v = attrNum(def, "w:val");
        if (v !== undefined) ddl.default = v;
      }
      opts.dropDownList = ddl as DropDownListOptions;
    } else {
      const textInput = findChild(el, "w:textInput");
      if (textInput) {
        const ti: Partial<TextInputOptions> = {};
        const type = findChild(textInput, "w:type");
        if (type) ti.type = attr(type, "w:val") as TextInputOptions["type"];
        const def = findChild(textInput, "w:default");
        if (def) ti.default = attr(def, "w:val");
        const maxLength = findChild(textInput, "w:maxLength");
        if (maxLength) {
          const v = attrNum(maxLength, "w:val");
          if (v !== undefined) ti.maxLength = v;
        }
        const format = findChild(textInput, "w:format");
        if (format) ti.format = attr(format, "w:val");
        opts.textInput = ti as TextInputOptions;
      }
    }
  }

  return opts as FormFieldOptions;
}
