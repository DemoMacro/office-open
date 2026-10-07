/**
 * x14:formControlPr — Excel form-control property part.
 *
 * @module
 */
import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, escapeXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/** Form-control kind (x14:formControlPr `@objectType`). */
export type FormControlObjectType =
  | "Button"
  | "Checkbox"
  | "Radio"
  | "Drop"
  | "Edit"
  | "GBox"
  | "Label"
  | "List"
  | "Scroll"
  | "Spin"
  | "Group";

/** Check state for checkbox and radio controls (`@checked`). */
export type FormControlChecked = "Checked" | "Unchecked" | "Mixed";

/** Drop-down presentation (`@dropStyle`). */
export type FormControlDropStyle = "Combo" | "ComboEdit" | "Simple";

/** Horizontal text alignment (`@textAlign`), e.g. textAlign="Center". */
export type FormControlTextAlign = "Left" | "Center" | "Right";

/** Vertical text alignment (`@vertical`), e.g. vertical="Top". */
export type FormControlVerticalAlign = "Top" | "Center" | "Bottom";

/** Excel form-control settings (x14:formControlPr). */
export interface FormControlPropertiesOptions {
  /** Control kind, e.g. objectType="Checkbox". */
  objectType: FormControlObjectType;
  /** Check state, e.g. checked="Checked". */
  checked?: FormControlChecked;
  /** Drop-down presentation, e.g. dropStyle="Combo". */
  dropStyle?: FormControlDropStyle;
  /** Linked-cell formula, e.g. fmlaLink="Sheet1!$A$1". */
  fmlaLink?: string;
  /** Source range formula, e.g. fmlaRange="Sheet1!$A$1:$A$4". */
  fmlaRange?: string;
  /** Macro formula, e.g. fmlaMacro="[0]!Module1.Run". */
  fmlaMacro?: string;
  /** Minimum scroll or spinner value, e.g. min="0". */
  min?: number;
  /** Maximum scroll or spinner value, e.g. max="100". */
  max?: number;
  /** Change per click, e.g. increment="1" (XML inc="1"). */
  increment?: number;
  /** Change per page click, e.g. page="10". */
  page?: number;
  /** Current scroll, spinner, or selection value, e.g. val="3". */
  val?: number;
  /** Horizontal text alignment, e.g. textAlign="Center". */
  textAlign?: FormControlTextAlign;
  /** Vertical text alignment, e.g. vertical="Top". */
  vertical?: FormControlVerticalAlign;
  /** True prevents user text edits (lockText="1"). */
  lockText?: boolean;
  /** True omits the first 3D effect (noThreeD="1"). */
  noThreeD?: boolean;
  /** True omits the second 3D effect (noThreeD2="1"). */
  noThreeD2?: boolean;
  /** List or choice labels in XML order, e.g. items=["Yes", "No"]. */
  items?: string[];
}

/** Structured parse failure with the exact control-part location. */
export class ControlPropertiesParseError extends Error {
  readonly part: string;
  readonly path: string;
  readonly name: string;
  readonly reason: string;

  constructor(part: string, path: string, name: string, reason: string) {
    super(`${part}${path}: ${name}: ${reason}`);
    this.part = part;
    this.path = path;
    this.name = name;
    this.reason = reason;
  }
}

const FORM_CONTROL_CHILDREN = new Set(["fmlaLink", "fmlaRange", "fmlaMacro", "item"]);
const FORM_CONTROL_OBJECT_TYPES = new Set([
  "Button",
  "Checkbox",
  "Radio",
  "Drop",
  "Edit",
  "GBox",
  "Label",
  "List",
  "Scroll",
  "Spin",
  "Group",
]);
const FORM_CONTROL_CHECKED = new Set(["Checked", "Unchecked", "Mixed"]);
const FORM_CONTROL_DROP_STYLES = new Set(["Combo", "ComboEdit", "Simple"]);
const FORM_CONTROL_TEXT_ALIGN = new Set(["Left", "Center", "Right"]);
const FORM_CONTROL_VERTICAL_ALIGN = new Set(["Top", "Center", "Bottom"]);
const FORM_CONTROL_OPTIONS = new Set([
  "objectType",
  "checked",
  "dropStyle",
  "fmlaLink",
  "fmlaRange",
  "fmlaMacro",
  "min",
  "max",
  "increment",
  "page",
  "val",
  "textAlign",
  "vertical",
  "lockText",
  "noThreeD",
  "noThreeD2",
  "items",
]);

function textOfElement(el: Element): string {
  return (el.elements ?? [])
    .map((child) => (child.type === "text" ? String(child.text ?? "") : ""))
    .join("");
}

function unsupportedOption(options: FormControlPropertiesOptions): string | undefined {
  return Object.keys(options).find((key) => !FORM_CONTROL_OPTIONS.has(key));
}

function readFormControlProperties(el: Element): FormControlPropertiesOptions {
  const objectType = attr(el, "objectType");
  if (!objectType) {
    throw new ControlPropertiesParseError(
      "xl/ctrlProps",
      "/formControlPr",
      "objectType",
      "missing required attribute",
    );
  }
  if (!FORM_CONTROL_OBJECT_TYPES.has(objectType)) {
    throw new ControlPropertiesParseError(
      "xl/ctrlProps",
      "/formControlPr",
      objectType,
      "unsupported objectType",
    );
  }
  const result: FormControlPropertiesOptions = {
    objectType: objectType as FormControlObjectType,
  };
  const checked = attr(el, "checked");
  if (checked) {
    if (!FORM_CONTROL_CHECKED.has(checked)) {
      throw new ControlPropertiesParseError(
        "xl/ctrlProps",
        "/formControlPr",
        checked,
        "unsupported checked value",
      );
    }
    result.checked = checked as FormControlChecked;
  }
  const dropStyle = attr(el, "dropStyle");
  if (dropStyle) {
    if (!FORM_CONTROL_DROP_STYLES.has(dropStyle)) {
      throw new ControlPropertiesParseError(
        "xl/ctrlProps",
        "/formControlPr",
        dropStyle,
        "unsupported dropStyle value",
      );
    }
    result.dropStyle = dropStyle as FormControlDropStyle;
  }
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    const childName = String(child.name ?? "");
    if (!FORM_CONTROL_CHILDREN.has(childName)) {
      throw new ControlPropertiesParseError(
        "xl/ctrlProps",
        "/formControlPr",
        childName,
        "unsupported element",
      );
    }
    if (childName === "fmlaLink") result.fmlaLink = textOfElement(child);
    else if (childName === "fmlaRange") result.fmlaRange = textOfElement(child);
    else if (childName === "fmlaMacro") result.fmlaMacro = textOfElement(child);
    else result.items = [...(result.items ?? []), textOfElement(child)];
  }
  const min = attrNum(el, "min");
  if (min !== undefined) result.min = min;
  const max = attrNum(el, "max");
  if (max !== undefined) result.max = max;
  const increment = attrNum(el, "inc");
  if (increment !== undefined) result.increment = increment;
  const page = attrNum(el, "page");
  if (page !== undefined) result.page = page;
  const val = attrNum(el, "val");
  if (val !== undefined) result.val = val;
  const textAlign = attr(el, "textAlign");
  if (textAlign) {
    if (!FORM_CONTROL_TEXT_ALIGN.has(textAlign)) {
      throw new ControlPropertiesParseError(
        "xl/ctrlProps",
        "/formControlPr",
        textAlign,
        "unsupported textAlign value",
      );
    }
    result.textAlign = textAlign as FormControlTextAlign;
  }
  const vertical = attr(el, "vertical");
  if (vertical) {
    if (!FORM_CONTROL_VERTICAL_ALIGN.has(vertical)) {
      throw new ControlPropertiesParseError(
        "xl/ctrlProps",
        "/formControlPr",
        vertical,
        "unsupported vertical value",
      );
    }
    result.vertical = vertical as FormControlVerticalAlign;
  }
  const lockText = parseOnOff(attr(el, "lockText"));
  if (lockText !== undefined) result.lockText = lockText;
  const noThreeD = parseOnOff(attr(el, "noThreeD"));
  if (noThreeD !== undefined) result.noThreeD = noThreeD;
  const noThreeD2 = parseOnOff(attr(el, "noThreeD2"));
  if (noThreeD2 !== undefined) result.noThreeD2 = noThreeD2;
  return result;
}

function writeFormControlProperties(options: FormControlPropertiesOptions): string {
  const unsupported = unsupportedOption(options);
  if (unsupported !== undefined) {
    throw new ControlPropertiesParseError(
      "xl/ctrlProps",
      "/formControlPr",
      unsupported,
      "unsupported option",
    );
  }
  const attributes = [`objectType="${options.objectType}"`];
  if (options.checked) attributes.push(`checked="${options.checked}"`);
  if (options.dropStyle) attributes.push(`dropStyle="${options.dropStyle}"`);
  if (options.min !== undefined) attributes.push(`min="${options.min}"`);
  if (options.max !== undefined) attributes.push(`max="${options.max}"`);
  if (options.increment !== undefined) attributes.push(`inc="${options.increment}"`);
  if (options.page !== undefined) attributes.push(`page="${options.page}"`);
  if (options.val !== undefined) attributes.push(`val="${options.val}"`);
  if (options.textAlign) attributes.push(`textAlign="${options.textAlign}"`);
  if (options.vertical) attributes.push(`vertical="${options.vertical}"`);
  if (options.lockText !== undefined) attributes.push(`lockText="${options.lockText ? 1 : 0}"`);
  if (options.noThreeD !== undefined) attributes.push(`noThreeD="${options.noThreeD ? 1 : 0}"`);
  if (options.noThreeD2 !== undefined) attributes.push(`noThreeD2="${options.noThreeD2 ? 1 : 0}"`);
  const children: string[] = [];
  if (options.fmlaLink !== undefined)
    children.push(`<fmlaLink>${escapeXml(options.fmlaLink)}</fmlaLink>`);
  if (options.fmlaRange !== undefined)
    children.push(`<fmlaRange>${escapeXml(options.fmlaRange)}</fmlaRange>`);
  if (options.fmlaMacro !== undefined)
    children.push(`<fmlaMacro>${escapeXml(options.fmlaMacro)}</fmlaMacro>`);
  for (const item of options.items ?? []) children.push(`<item>${escapeXml(item)}</item>`);
  const childXml = children.join("");
  return `<x14:formControlPr xmlns:x14="${FORM_CONTROL_NS}" ${attributes.join(" ")}>${
    childXml ? childXml : ""
  }</x14:formControlPr>`;
}

const FORM_CONTROL_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

export const controlPropertiesDesc: CustomDescriptor<FormControlPropertiesOptions> = {
  kind: "custom",
  stringify(options) {
    return writeFormControlProperties(options);
  },
  parse(el) {
    return readFormControlProperties(el);
  },
};
