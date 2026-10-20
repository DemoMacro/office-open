import type {
  ComboBoxControlOptions,
  FormContainerOptions,
  FormControlOptions,
  FormEventListenerOptions,
  FormListPropertyOptions,
  FormPropertyOptions,
  GridControlOptions,
  GridColumnOptions,
  GridFormControlOptions,
  ListBoxControlOptions,
  OfficeFormsOptions,
  TextAreaControlOptions,
  XFormsElementOptions,
  XFormsModelOptions,
} from "@office-open/core";
import type { Element } from "@office-open/xml";

import { OdfSchemaError } from "./errors";
import { attributeString, childrenNamed, escapeText, textOf, xmlElement } from "./runtime/xml";

type ControlKind =
  | "text"
  | "textArea"
  | "formattedText"
  | "number"
  | "date"
  | "time"
  | "comboBox"
  | "listBox"
  | "checkbox"
  | "password"
  | "file"
  | "fixedText"
  | "button"
  | "image"
  | "radio"
  | "frame"
  | "imageFrame"
  | "hidden"
  | "grid"
  | "valueRange"
  | "genericControl";

type AttributeType =
  | "string"
  | "xmlId"
  | "formId"
  | "boolean"
  | "number"
  | "integer"
  | "positiveInteger"
  | "nonNegativeInteger"
  | "date"
  | "time"
  | "duration"
  | "triState"
  | "buttonType"
  | "commandType"
  | "listSourceType"
  | "listLinkageType"
  | "visualEffect"
  | "imagePosition"
  | "imageAlign"
  | "navigationMode"
  | "tabCycle"
  | "orientation";

interface AttributeSpec {
  field: string;
  type: AttributeType;
}

const BASE_ATTRIBUTES: Record<string, AttributeSpec> = {
  "xml:id": { field: "xmlId", type: "xmlId" },
  "form:id": { field: "id", type: "formId" },
  "form:name": { field: "name", type: "string" },
  "form:control-implementation": { field: "controlImplementation", type: "string" },
  "xforms:bind": { field: "bind", type: "string" },
  "form:input-required": { field: "inputRequired", type: "boolean" },
};

const INPUT_ATTRIBUTES: Record<string, AttributeSpec> = {
  ...BASE_ATTRIBUTES,
  "form:current-value": { field: "currentValue", type: "string" },
  "form:disabled": { field: "disabled", type: "boolean" },
  "form:max-length": { field: "maxLength", type: "nonNegativeInteger" },
  "form:printable": { field: "printable", type: "boolean" },
  "form:readonly": { field: "readonly", type: "boolean" },
  "form:tab-index": { field: "tabIndex", type: "nonNegativeInteger" },
  "form:tab-stop": { field: "tabStop", type: "boolean" },
  "form:title": { field: "title", type: "string" },
  "form:value": { field: "value", type: "string" },
  "form:convert-empty-to-null": { field: "convertEmptyToNull", type: "boolean" },
  "form:data-field": { field: "dataField", type: "string" },
  "form:linked-cell": { field: "linkedCell", type: "string" },
};

const NUMERIC_CONTROL_ATTRIBUTES: Record<string, AttributeSpec> = {
  ...BASE_ATTRIBUTES,
  "form:disabled": { field: "disabled", type: "boolean" },
  "form:max-length": { field: "maxLength", type: "nonNegativeInteger" },
  "form:printable": { field: "printable", type: "boolean" },
  "form:readonly": { field: "readonly", type: "boolean" },
  "form:tab-index": { field: "tabIndex", type: "nonNegativeInteger" },
  "form:tab-stop": { field: "tabStop", type: "boolean" },
  "form:title": { field: "title", type: "string" },
  "form:convert-empty-to-null": { field: "convertEmptyToNull", type: "boolean" },
  "form:data-field": { field: "dataField", type: "string" },
  "form:linked-cell": { field: "linkedCell", type: "string" },
  "form:spin-button": { field: "spinButton", type: "boolean" },
  "form:repeat": { field: "repeat", type: "boolean" },
  "form:delay-for-repeat": { field: "delayForRepeat", type: "duration" },
  "form:validation": { field: "validation", type: "boolean" },
};

const CONTROL_ATTRIBUTES: Record<ControlKind, Record<string, AttributeSpec>> = {
  text: INPUT_ATTRIBUTES,
  textArea: INPUT_ATTRIBUTES,
  formattedText: {
    ...INPUT_ATTRIBUTES,
    "form:spin-button": { field: "spinButton", type: "boolean" },
    "form:repeat": { field: "repeat", type: "boolean" },
    "form:delay-for-repeat": { field: "delayForRepeat", type: "duration" },
    "form:max-value": { field: "maxValue", type: "string" },
    "form:min-value": { field: "minValue", type: "string" },
    "form:validation": { field: "validation", type: "boolean" },
  },
  number: {
    ...NUMERIC_CONTROL_ATTRIBUTES,
    "form:value": { field: "value", type: "number" },
    "form:current-value": { field: "currentValue", type: "number" },
    "form:min-value": { field: "minValue", type: "number" },
    "form:max-value": { field: "maxValue", type: "number" },
  },
  date: {
    ...NUMERIC_CONTROL_ATTRIBUTES,
    "form:value": { field: "value", type: "date" },
    "form:current-value": { field: "currentValue", type: "date" },
    "form:min-value": { field: "minValue", type: "date" },
    "form:max-value": { field: "maxValue", type: "date" },
  },
  time: {
    ...NUMERIC_CONTROL_ATTRIBUTES,
    "form:value": { field: "value", type: "time" },
    "form:current-value": { field: "currentValue", type: "time" },
    "form:min-value": { field: "minValue", type: "time" },
    "form:max-value": { field: "maxValue", type: "time" },
  },
  comboBox: {
    ...INPUT_ATTRIBUTES,
    "form:dropdown": { field: "dropdown", type: "boolean" },
    "form:size": { field: "size", type: "nonNegativeInteger" },
    "form:list-source": { field: "listSource", type: "string" },
    "form:list-source-type": { field: "listSourceType", type: "listSourceType" },
    "form:source-cell-range": { field: "sourceCellRange", type: "string" },
    "form:auto-complete": { field: "autoComplete", type: "boolean" },
    "form:xforms-list-source": { field: "xformsListSource", type: "string" },
  },
  listBox: {
    ...BASE_ATTRIBUTES,
    "form:disabled": { field: "disabled", type: "boolean" },
    "form:dropdown": { field: "dropdown", type: "boolean" },
    "form:printable": { field: "printable", type: "boolean" },
    "form:size": { field: "size", type: "nonNegativeInteger" },
    "form:tab-index": { field: "tabIndex", type: "nonNegativeInteger" },
    "form:tab-stop": { field: "tabStop", type: "boolean" },
    "form:title": { field: "title", type: "string" },
    "form:bound-column": { field: "boundColumn", type: "string" },
    "form:data-field": { field: "dataField", type: "string" },
    "form:list-source": { field: "listSource", type: "string" },
    "form:list-source-type": { field: "listSourceType", type: "listSourceType" },
    "form:linked-cell": { field: "linkedCell", type: "string" },
    "form:list-linkage-type": { field: "listLinkageType", type: "listLinkageType" },
    "form:source-cell-range": { field: "sourceCellRange", type: "string" },
    "form:multiple": { field: "multiple", type: "boolean" },
    "form:xforms-list-source": { field: "xformsListSource", type: "string" },
  },
  checkbox: {
    ...BASE_ATTRIBUTES,
    "form:disabled": { field: "disabled", type: "boolean" },
    "form:label": { field: "label", type: "string" },
    "form:printable": { field: "printable", type: "boolean" },
    "form:tab-index": { field: "tabIndex", type: "nonNegativeInteger" },
    "form:tab-stop": { field: "tabStop", type: "boolean" },
    "form:title": { field: "title", type: "string" },
    "form:value": { field: "value", type: "string" },
    "form:data-field": { field: "dataField", type: "string" },
    "form:visual-effect": { field: "visualEffect", type: "visualEffect" },
    "form:image-position": { field: "imagePosition", type: "imagePosition" },
    "form:image-align": { field: "imageAlign", type: "imageAlign" },
    "form:linked-cell": { field: "linkedCell", type: "string" },
    "form:current-state": { field: "currentState", type: "triState" },
    "form:is-tristate": { field: "triState", type: "boolean" },
    "form:state": { field: "state", type: "triState" },
  },
  password: {
    ...BASE_ATTRIBUTES,
    "form:disabled": { field: "disabled", type: "boolean" },
    "form:max-length": { field: "maxLength", type: "nonNegativeInteger" },
    "form:printable": { field: "printable", type: "boolean" },
    "form:tab-index": { field: "tabIndex", type: "nonNegativeInteger" },
    "form:tab-stop": { field: "tabStop", type: "boolean" },
    "form:title": { field: "title", type: "string" },
    "form:value": { field: "value", type: "string" },
    "form:convert-empty-to-null": { field: "convertEmptyToNull", type: "boolean" },
    "form:linked-cell": { field: "linkedCell", type: "string" },
    "form:echo-char": { field: "echoChar", type: "string" },
  },
  file: INPUT_ATTRIBUTES,
  fixedText: {
    ...BASE_ATTRIBUTES,
    "form:for": { field: "for", type: "string" },
    "form:disabled": { field: "disabled", type: "boolean" },
    "form:label": { field: "label", type: "string" },
    "form:printable": { field: "printable", type: "boolean" },
    "form:title": { field: "title", type: "string" },
    "form:multi-line": { field: "multiLine", type: "boolean" },
  },
  button: {
    ...BASE_ATTRIBUTES,
    "form:button-type": { field: "buttonType", type: "buttonType" },
    "form:disabled": { field: "disabled", type: "boolean" },
    "form:label": { field: "label", type: "string" },
    "form:image-data": { field: "imageData", type: "string" },
    "form:printable": { field: "printable", type: "boolean" },
    "form:tab-index": { field: "tabIndex", type: "nonNegativeInteger" },
    "form:tab-stop": { field: "tabStop", type: "boolean" },
    "office:target-frame": { field: "targetFrame", type: "string" },
    "xlink:href": { field: "href", type: "string" },
    "form:title": { field: "title", type: "string" },
    "form:value": { field: "value", type: "string" },
    "form:image-position": { field: "imagePosition", type: "imagePosition" },
    "form:image-align": { field: "imageAlign", type: "imageAlign" },
    "form:repeat": { field: "repeat", type: "boolean" },
    "form:delay-for-repeat": { field: "delayForRepeat", type: "duration" },
    "form:default-button": { field: "defaultButton", type: "boolean" },
    "form:toggle": { field: "toggle", type: "boolean" },
    "form:focus-on-click": { field: "focusOnClick", type: "boolean" },
    "form:xforms-submission": { field: "xformsSubmission", type: "string" },
  },
  image: {
    ...BASE_ATTRIBUTES,
    "form:button-type": { field: "buttonType", type: "buttonType" },
    "form:disabled": { field: "disabled", type: "boolean" },
    "form:image-data": { field: "imageData", type: "string" },
    "form:printable": { field: "printable", type: "boolean" },
    "form:tab-index": { field: "tabIndex", type: "nonNegativeInteger" },
    "form:tab-stop": { field: "tabStop", type: "boolean" },
    "office:target-frame": { field: "targetFrame", type: "string" },
    "xlink:href": { field: "href", type: "string" },
    "form:title": { field: "title", type: "string" },
    "form:value": { field: "value", type: "string" },
  },
  radio: {
    ...BASE_ATTRIBUTES,
    "formx:group-name": { field: "groupName", type: "string" },
    "form:current-selected": { field: "currentSelected", type: "boolean" },
    "form:disabled": { field: "disabled", type: "boolean" },
    "form:label": { field: "label", type: "string" },
    "form:printable": { field: "printable", type: "boolean" },
    "form:selected": { field: "selected", type: "boolean" },
    "form:tab-index": { field: "tabIndex", type: "nonNegativeInteger" },
    "form:tab-stop": { field: "tabStop", type: "boolean" },
    "form:title": { field: "title", type: "string" },
    "form:value": { field: "value", type: "string" },
    "form:data-field": { field: "dataField", type: "string" },
    "form:visual-effect": { field: "visualEffect", type: "visualEffect" },
    "form:image-position": { field: "imagePosition", type: "imagePosition" },
    "form:image-align": { field: "imageAlign", type: "imageAlign" },
    "form:linked-cell": { field: "linkedCell", type: "string" },
  },
  frame: {
    ...BASE_ATTRIBUTES,
    "form:disabled": { field: "disabled", type: "boolean" },
    "form:for": { field: "for", type: "string" },
    "form:label": { field: "label", type: "string" },
    "form:printable": { field: "printable", type: "boolean" },
    "form:title": { field: "title", type: "string" },
  },
  imageFrame: {
    ...BASE_ATTRIBUTES,
    "form:label": { field: "label", type: "string" },
    "form:disabled": { field: "disabled", type: "boolean" },
    "form:image-data": { field: "imageData", type: "string" },
    "form:printable": { field: "printable", type: "boolean" },
    "form:readonly": { field: "readonly", type: "boolean" },
    "form:title": { field: "title", type: "string" },
    "form:data-field": { field: "dataField", type: "string" },
  },
  hidden: {
    ...BASE_ATTRIBUTES,
    "form:value": { field: "value", type: "string" },
  },
  grid: {
    ...BASE_ATTRIBUTES,
    "form:disabled": { field: "disabled", type: "boolean" },
    "form:printable": { field: "printable", type: "boolean" },
    "form:tab-index": { field: "tabIndex", type: "nonNegativeInteger" },
    "form:tab-stop": { field: "tabStop", type: "boolean" },
    "form:title": { field: "title", type: "string" },
  },
  valueRange: {
    ...BASE_ATTRIBUTES,
    "form:disabled": { field: "disabled", type: "boolean" },
    "form:printable": { field: "printable", type: "boolean" },
    "form:tab-index": { field: "tabIndex", type: "nonNegativeInteger" },
    "form:tab-stop": { field: "tabStop", type: "boolean" },
    "form:title": { field: "title", type: "string" },
    "form:value": { field: "value", type: "string" },
    "form:linked-cell": { field: "linkedCell", type: "string" },
    "form:repeat": { field: "repeat", type: "boolean" },
    "form:delay-for-repeat": { field: "delayForRepeat", type: "duration" },
    "form:max-value": { field: "maxValue", type: "integer" },
    "form:min-value": { field: "minValue", type: "integer" },
    "form:step-size": { field: "stepSize", type: "positiveInteger" },
    "form:page-step-size": { field: "pageStepSize", type: "positiveInteger" },
    "form:orientation": { field: "orientation", type: "orientation" },
  },
  genericControl: BASE_ATTRIBUTES,
};

const CONTROL_ELEMENT_KINDS: Record<string, ControlKind> = {
  "form:text": "text",
  "form:textarea": "textArea",
  "form:formatted-text": "formattedText",
  "form:number": "number",
  "form:date": "date",
  "form:time": "time",
  "form:combobox": "comboBox",
  "form:listbox": "listBox",
  "form:checkbox": "checkbox",
  "form:password": "password",
  "form:file": "file",
  "form:fixed-text": "fixedText",
  "form:button": "button",
  "form:image": "image",
  "form:radio": "radio",
  "form:frame": "frame",
  "form:image-frame": "imageFrame",
  "form:hidden": "hidden",
  "form:grid": "grid",
  "form:value-range": "valueRange",
  "form:generic-control": "genericControl",
};

const GRID_CONTROL_KINDS = new Set<ControlKind>([
  "text",
  "textArea",
  "formattedText",
  "number",
  "date",
  "time",
  "comboBox",
  "listBox",
  "checkbox",
]);

const KIND_ELEMENTS = Object.fromEntries(
  Object.entries(CONTROL_ELEMENT_KINDS).map(([element, kind]) => [kind, element]),
) as Record<ControlKind, string>;

const LIST_SOURCE_TYPES: Record<string, string> = {
  table: "table",
  query: "query",
  sql: "sql",
  "sql-pass-through": "sqlPassThrough",
  "value-list": "valueList",
  "table-fields": "tableFields",
};
const LIST_LINKAGE_TYPES = { selection: "selection", "selection-indices": "selectionIndices" };
const BUTTON_TYPES = { submit: "submit", reset: "reset", push: "push", url: "url" };
const COMMAND_TYPES = { table: "table", query: "query", command: "command" };
const VISUAL_EFFECTS = { flat: "flat", "3d": "3d" };
const IMAGE_POSITIONS = {
  start: "start",
  center: "center",
  end: "end",
  top: "top",
  bottom: "bottom",
};
const IMAGE_ALIGNS = { start: "start", center: "center", end: "end" };
const NAVIGATION_MODES = { none: "none", current: "current", parent: "parent" };
const TAB_CYCLES = { records: "records", current: "current", page: "page" };
const ORIENTATIONS = { horizontal: "horizontal", vertical: "vertical" };
const TRI_STATES = { unchecked: "unchecked", checked: "checked", unknown: "unknown" };

const FORM_ATTRIBUTES: Record<string, AttributeSpec> = {
  "form:name": { field: "name", type: "string" },
  "form:control-implementation": { field: "controlImplementation", type: "string" },
  "xlink:href": { field: "href", type: "string" },
  "office:target-frame": { field: "targetFrame", type: "string" },
  "form:method": { field: "method", type: "string" },
  "form:enctype": { field: "enctype", type: "string" },
  "form:allow-deletes": { field: "allowDeletes", type: "boolean" },
  "form:allow-inserts": { field: "allowInserts", type: "boolean" },
  "form:allow-updates": { field: "allowUpdates", type: "boolean" },
  "form:apply-filter": { field: "applyFilter", type: "boolean" },
  "form:command-type": { field: "commandType", type: "commandType" },
  "form:command": { field: "command", type: "string" },
  "form:datasource": { field: "datasource", type: "string" },
  "form:master-fields": { field: "masterFields", type: "string" },
  "form:detail-fields": { field: "detailFields", type: "string" },
  "form:escape-processing": { field: "escapeProcessing", type: "boolean" },
  "form:filter": { field: "filter", type: "string" },
  "form:ignore-result": { field: "ignoreResult", type: "boolean" },
  "form:navigation-mode": { field: "navigationMode", type: "navigationMode" },
  "form:order": { field: "order", type: "string" },
  "form:tab-cycle": { field: "tabCycle", type: "tabCycle" },
};

export const FORM_SCHEMA_ELEMENTS: readonly { name: string; attributes: readonly string[] }[] = [
  { name: "office:forms", attributes: ["form:automatic-focus", "form:apply-design-mode"] },
  { name: "form:form", attributes: Object.keys(FORM_ATTRIBUTES) },
  { name: "form:connection-resource", attributes: ["xlink:href"] },
  ...Object.entries(CONTROL_ELEMENT_KINDS).map(([name, kind]) => ({
    name,
    attributes: Object.keys(CONTROL_ATTRIBUTES[kind]),
  })),
  {
    name: "form:column",
    attributes: ["form:name", "form:control-implementation", "form:label", "form:text-style-name"],
  },
  { name: "form:properties", attributes: [] },
  {
    name: "form:property",
    attributes: [
      "form:property-name",
      "office:value-type",
      "office:value",
      "office:currency",
      "office:date-value",
      "office:time-value",
      "office:boolean-value",
      "office:string-value",
    ],
  },
  { name: "form:list-property", attributes: ["form:property-name", "office:value-type"] },
  {
    name: "form:list-value",
    attributes: [
      "office:value",
      "office:date-value",
      "office:time-value",
      "office:boolean-value",
      "office:string-value",
    ],
  },
  {
    name: "form:option",
    attributes: ["form:current-selected", "form:selected", "form:label", "form:value"],
  },
  { name: "form:item", attributes: ["form:label"] },
  {
    name: "script:event-listener",
    attributes: [
      "script:event-name",
      "script:language",
      "script:macro-name",
      "xlink:href",
      "xlink:actuate",
    ],
  },
  { name: "xforms:model", attributes: [] },
];

function formError(part: string, path: string, name: string, reason: string): OdfSchemaError {
  return new OdfSchemaError(
    `${part}: ${path}/${name}: ${reason}`,
    part,
    `${path}/${name}`,
    name,
    reason,
  );
}

function assertNoUnknownAttributes(
  element: Element,
  allowed: Iterable<string>,
  part: string,
  path: string,
): void {
  const names = new Set(allowed);
  for (const name of Object.keys(element.attributes ?? {})) {
    if (name === "xmlns" || name.startsWith("xmlns:")) continue;
    if (!names.has(name))
      throw formError(part, path, name, "attribute is not allowed by the form model");
  }
}

function parseDate(value: string): boolean {
  const match = /^(-?\d{4,})-(\d{2})-(\d{2})(?:Z|[+-]\d{2}:\d{2})?$/.exec(value);
  if (!match) return false;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (month < 1 || month > 12 || day < 1 || year === 0) return false;
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const maximumDay = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1];
  return day <= (maximumDay ?? 0);
}

function parseTime(value: string): boolean {
  const match =
    /^(\d{2}):(\d{2}):(\d{2}(?:\.\d+)?)(?:Z|[+-]\d{2}:\d{2})?$/.exec(value) ??
    /^(\d{2})(\d{2})(\d{2,8}(?:\.\d+)?)(?:Z|[+-]\d{2}:\d{2})?$/.exec(value);
  if (!match) return false;
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = Number(match[3]);
  if (hour > 24 || minute > 59 || Number(second) >= 60) return false;
  return hour < 24 || (minute === 0 && second === 0);
}

function parseDuration(value: string): boolean {
  return (
    /^-?P(?:\d+Y)?(?:\d+M)?(?:\d+D)?(?:T(?:\d+H)?(?:\d+M)?(?:\d+(?:\.\d+)?S)?)?$/.test(value) &&
    /\d/.test(value)
  );
}

function parseEnum(value: string, mapping: Record<string, string>): string {
  return mapping[value] ?? "";
}

function parseAttributeValue(value: string, type: AttributeType): unknown {
  if (type === "boolean") {
    if (value === "true") return true;
    if (value === "false") return false;
    throw new Error("invalid boolean");
  }
  if (type === "integer" || type === "positiveInteger" || type === "nonNegativeInteger") {
    if (!/^[+-]?\d+$/.test(value)) throw new Error("invalid integer");
    const parsed = Number(value);
    if (type === "positiveInteger" && parsed <= 0) throw new Error("invalid positive integer");
    if (type === "nonNegativeInteger" && parsed < 0)
      throw new Error("invalid non-negative integer");
    return parsed;
  }
  if (type === "number") {
    const parsed = Number(value);
    if (!Number.isFinite(parsed)) throw new Error("invalid number");
    return parsed;
  }
  if (type === "xmlId" || type === "formId") return value;
  if (type === "date" && !parseDate(value)) throw new Error("invalid date");
  if (type === "time" && !parseTime(value)) throw new Error("invalid time");
  if (type === "duration" && !parseDuration(value)) throw new Error("invalid duration");
  if (type === "triState") {
    const parsed = parseEnum(value, TRI_STATES);
    if (!parsed) throw new Error("invalid checkbox state");
    return parsed;
  }
  const mappings: Record<string, Record<string, string>> = {
    buttonType: BUTTON_TYPES,
    commandType: COMMAND_TYPES,
    listSourceType: LIST_SOURCE_TYPES,
    listLinkageType: LIST_LINKAGE_TYPES,
    visualEffect: VISUAL_EFFECTS,
    imagePosition: IMAGE_POSITIONS,
    imageAlign: IMAGE_ALIGNS,
    navigationMode: NAVIGATION_MODES,
    tabCycle: TAB_CYCLES,
    orientation: ORIENTATIONS,
  };
  if (mappings[type]) {
    const parsed = parseEnum(value, mappings[type]);
    if (!parsed) throw new Error("invalid enum value");
    return parsed;
  }
  return value;
}

function isValidXmlId(value: string): boolean {
  return /^[\p{L}_][\p{L}\p{N}\p{M}._-]*$/u.test(value);
}

class FormIdAllocator {
  readonly used = new Set<string>();
  readonly formIds = new Set<string>();
  duplicateId: string | undefined;
  private next = 1;

  constructor(root?: Element) {
    if (root) this.collect(root);
  }

  private collect(element: Element): void {
    const xmlId = attributeString(element, "xml:id");
    const formId = attributeString(element, "form:id");
    if (xmlId !== undefined) {
      if (this.used.has(xmlId)) this.duplicateId = xmlId;
      this.used.add(xmlId);
    }
    if (formId !== undefined) {
      if (this.formIds.has(formId)) this.duplicateId = formId;
      this.formIds.add(formId);
    }
    for (const child of element.elements ?? []) {
      if (child.type === "element") this.collect(child);
    }
  }

  reserve(value: string): void {
    if (this.used.has(value)) throw new Error("duplicate form control id");
    this.used.add(value);
  }

  allocate(preferred?: string): string {
    if (preferred !== undefined && preferred !== "" && !this.used.has(preferred)) {
      this.used.add(preferred);
      return preferred;
    }
    let value = `form-control-${this.next}`;
    while (this.used.has(value)) {
      this.next += 1;
      value = `form-control-${this.next}`;
    }
    this.used.add(value);
    this.next += 1;
    return value;
  }
}

function parseSpecifiedAttributes(
  element: Element,
  specs: Record<string, AttributeSpec>,
  part: string,
  path: string,
  ignoredAttributes: readonly string[] = [],
): Record<string, unknown> {
  assertNoUnknownAttributes(element, [...Object.keys(specs), ...ignoredAttributes], part, path);
  const result: Record<string, unknown> = {};
  for (const [attributeName, spec] of Object.entries(specs)) {
    const value = attributeString(element, attributeName);
    if (value === undefined) continue;
    try {
      result[spec.field] = parseAttributeValue(value, spec.type);
    } catch {
      throw formError(part, path, attributeName, `invalid ${spec.type} value: ${value}`);
    }
  }
  return result;
}

function requiredAttribute(element: Element, name: string, part: string, path: string): string {
  const value = attributeString(element, name);
  if (value === undefined) throw formError(part, path, name, "attribute is required");
  return value;
}

function parseRequiredBoolean(value: string, name: string, part: string, path: string): boolean {
  try {
    return parseAttributeValue(value, "boolean") === true;
  } catch {
    throw formError(part, path, name, `invalid boolean value: ${value}`);
  }
}

function parseDateOrDateTime(value: string): boolean {
  const match =
    /^(-?\d{4,})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:\.\d+)?)(?:Z|[+-]\d{2}:\d{2})?$/.exec(
      value,
    );
  if (!match) return parseDate(value);
  return (
    parseDate(`${match[1]}-${match[2]}-${match[3]}`) &&
    Number(match[4]) <= 24 &&
    Number(match[5]) <= 59 &&
    Number(match[6]) < 60
  );
}

function parsePropertyContainer(
  element: Element,
  part: string,
  path: string,
): Array<FormPropertyOptions | FormListPropertyOptions> {
  const children = (element.elements ?? []).filter(
    (child) => child.type === "element",
  ) as Element[];
  const unknown = children.find(
    (child) => child.name !== "form:property" && child.name !== "form:list-property",
  );
  if (unknown)
    throw formError(part, path, unknown.name ?? "unknown", "element is not a form property");
  if (!children.length)
    throw formError(part, path, element.name ?? "unknown", "properties container is empty");
  return children.map((child, index) => {
    const propertyPath = `${path}/${child.name}[${index + 1}]`;
    const name = requiredAttribute(child, "form:property-name", part, propertyPath);
    const type = attributeString(child, "office:value-type");
    const allowed = [
      "form:property-name",
      "office:value-type",
      "office:value",
      "office:currency",
      "office:date-value",
      "office:time-value",
      "office:boolean-value",
      "office:string-value",
    ];
    assertNoUnknownAttributes(child, allowed, part, propertyPath);
    return child.name === "form:property"
      ? parseSingleProperty(child, name, type, part, propertyPath)
      : parseListProperty(child, name, type, part, propertyPath);
  });
}

function parseSingleProperty(
  element: Element,
  name: string,
  type: string | undefined,
  part: string,
  path: string,
): FormPropertyOptions {
  if (type === "float" || type === "percentage" || type === "currency") {
    const lexical = requiredAttribute(element, "office:value", part, path);
    const parsed = parseAttributeValue(lexical, "number");
    if (typeof parsed !== "number")
      throw formError(part, path, "office:value", "invalid numeric property");
    const currency = type === "currency" ? attributeString(element, "office:currency") : undefined;
    return {
      name,
      value: currency === undefined ? { type, value: parsed } : { type, value: parsed, currency },
    };
  }
  if (type === "date" || type === "time") {
    const attributeName = type === "date" ? "office:date-value" : "office:time-value";
    const lexical = requiredAttribute(element, attributeName, part, path);
    if (
      (type === "date" && !parseDateOrDateTime(lexical)) ||
      (type === "time" && !parseDuration(lexical))
    )
      throw formError(part, path, attributeName, `invalid ${type} property value: ${lexical}`);
    return { name, value: { type, value: lexical } };
  }
  if (type === "boolean") {
    const lexical = requiredAttribute(element, "office:boolean-value", part, path);
    return { name, value: { type, value: parseAttributeValue(lexical, "boolean") === true } };
  }
  if (type === "string") {
    const lexical = attributeString(element, "office:string-value");
    return { name, value: { type, ...(lexical === undefined ? {} : { value: lexical }) } };
  }
  if (type === "void") return { name, value: { type } };
  throw formError(
    part,
    path,
    "office:value-type",
    `invalid form property type: ${type ?? "(missing)"}`,
  );
}

function parseListProperty(
  element: Element,
  name: string,
  type: string | undefined,
  part: string,
  path: string,
): FormListPropertyOptions {
  const children = childrenNamed(element, "form:list-value");
  const unknown = (element.elements ?? []).find(
    (child) => child.type === "element" && child.name !== "form:list-value",
  );
  if (unknown)
    throw formError(part, path, unknown.name ?? "unknown", "element is not a list property value");
  if (
    !["float", "percentage", "currency", "date", "time", "boolean", "string"].includes(type ?? "")
  )
    throw formError(
      part,
      path,
      "office:value-type",
      `invalid form list property type: ${type ?? "(missing)"}`,
    );
  const valueType = type as FormListPropertyOptions["value"]["type"];
  const attributeName =
    valueType === "date"
      ? "office:date-value"
      : valueType === "time"
        ? "office:time-value"
        : valueType === "boolean"
          ? "office:boolean-value"
          : valueType === "string"
            ? "office:string-value"
            : "office:value";
  const values: Array<string | number | boolean> = [];
  for (const [index, child] of children.entries()) {
    const childPath = `${path}/form:list-value[${index + 1}]`;
    const lexical = requiredAttribute(child, attributeName, part, childPath);
    if (valueType === "date" && !parseDateOrDateTime(lexical))
      throw formError(part, path, attributeName, "invalid date list value");
    if (valueType === "time" && !parseDuration(lexical))
      throw formError(part, path, attributeName, "invalid time list value");
    if (valueType === "float" || valueType === "percentage" || valueType === "currency") {
      const parsed = parseAttributeValue(lexical, "number");
      values.push(parsed as number);
      continue;
    }
    if (valueType === "boolean") values.push(parseAttributeValue(lexical, "boolean") === true);
    else values.push(lexical);
  }
  const currency =
    valueType === "currency" ? attributeString(element, "office:currency") : undefined;
  const value: FormListPropertyOptions["value"] =
    valueType === "float" || valueType === "percentage"
      ? { type: valueType, values: values as number[] }
      : valueType === "currency"
        ? {
            type: valueType,
            values: values as number[],
            ...(currency === undefined ? {} : { currency }),
          }
        : valueType === "date" || valueType === "time" || valueType === "string"
          ? { type: valueType, values: values as string[] }
          : { type: valueType, values: values as boolean[] };
  return { name, value };
}

function parseEventContainer(
  element: Element,
  part: string,
  path: string,
): FormEventListenerOptions[] {
  const children = (element.elements ?? []).filter(
    (child) => child.type === "element",
  ) as Element[];
  const unknown = children.find((child) => child.name !== "script:event-listener");
  if (unknown)
    throw formError(
      part,
      path,
      unknown.name ?? "unknown",
      "element is not a script event listener",
    );
  return children.map((child, index) => {
    const eventPath = `${path}/script:event-listener[${index + 1}]`;
    assertNoUnknownAttributes(
      child,
      [
        "script:event-name",
        "script:language",
        "script:macro-name",
        "xlink:type",
        "xlink:href",
        "xlink:actuate",
      ],
      part,
      eventPath,
    );
    const eventName = requiredAttribute(child, "script:event-name", part, eventPath);
    const language = requiredAttribute(child, "script:language", part, eventPath);
    const macroName = attributeString(child, "script:macro-name");
    const href = attributeString(child, "xlink:href");
    if (macroName !== undefined && href === undefined)
      return { target: "macro", eventName, language, macroName };
    if (href !== undefined && macroName === undefined) {
      if ((attributeString(child, "xlink:type") ?? "simple") !== "simple")
        throw formError(part, eventPath, "xlink:type", "event link type must be simple");
      const actuate = attributeString(child, "xlink:actuate");
      if (actuate !== undefined && actuate !== "onRequest")
        throw formError(part, eventPath, "xlink:actuate", "event link actuation must be onRequest");
      return {
        target: "url",
        eventName,
        language,
        href,
        ...(actuate === undefined ? {} : { actuate }),
      };
    }
    throw formError(
      part,
      eventPath,
      "script:macro-name",
      "event must use either macro-name or an href",
    );
  });
}

function parseXFormsElement(element: Element): XFormsElementOptions {
  const children: XFormsElementOptions["children"] = element.elements?.map((child) =>
    child.type === "element" ? parseXFormsElement(child) : String(child.text ?? ""),
  );
  return {
    name: element.name ?? "",
    attributes: Object.fromEntries(
      Object.entries(element.attributes ?? {}).map(([name, value]) => [name, String(value)]),
    ),
    children: children?.length ? children : undefined,
  };
}

export function parseXFormsModel(element: Element, part: string, path: string): XFormsModelOptions {
  if (element.name !== "xforms:model")
    throw formError(part, path, element.name ?? "unknown", "element is not an XForms model");
  const children: XFormsModelOptions["children"] = element.elements
    ?.filter((child) => child.type === "element")
    .map((child) => parseXFormsElement(child as Element));
  return {
    attributes: Object.fromEntries(
      Object.entries(element.attributes ?? {}).map(([name, value]) => [name, String(value)]),
    ),
    children: children?.length ? children : undefined,
  };
}

function parseControl(
  element: Element,
  part: string,
  path: string,
  ids: FormIdAllocator,
  grid = false,
): FormControlOptions {
  const kind = CONTROL_ELEMENT_KINDS[element.name ?? ""];
  if (!kind)
    throw formError(
      part,
      path,
      element.name ?? "unknown",
      "element has no canonical form control mapping",
    );
  const controlPath = `${path}/${element.name}`;
  if (grid && !GRID_CONTROL_KINDS.has(kind))
    throw formError(
      part,
      controlPath,
      element.name ?? "unknown",
      "control is not allowed in a form grid column",
    );
  const attributes = parseSpecifiedAttributes(element, CONTROL_ATTRIBUTES[kind], part, controlPath);
  const id = attributes.id as string | undefined;
  const explicitXmlId = attributes.xmlId as string | undefined;
  if (id !== undefined && !isValidXmlId(id))
    throw formError(part, controlPath, "form:id", "invalid NCName");
  if (explicitXmlId !== undefined && !isValidXmlId(explicitXmlId))
    throw formError(part, controlPath, "xml:id", "invalid ID");
  attributes.xmlId = explicitXmlId ?? ids.allocate(id);
  const control = { control: kind, ...attributes } as FormControlOptions;
  for (const child of element.elements ?? []) {
    if (child.type === "text") continue;
    const childElement = child;
    if (childElement.name === "form:properties") {
      control.properties = parsePropertyContainer(childElement, part, controlPath);
    } else if (childElement.name === "office:event-listeners") {
      control.events = parseEventContainer(childElement, part, controlPath);
    } else if (kind === "textArea" && childElement.name === "text:p") {
      const textArea = control as TextAreaControlOptions;
      textArea.paragraphs = [...(textArea.paragraphs ?? []), textOf(childElement)];
    } else if (kind === "comboBox" && childElement.name === "form:item") {
      const comboBox = control as ComboBoxControlOptions;
      assertNoUnknownAttributes(childElement, ["form:label"], part, `${controlPath}/form:item`);
      const labelAttribute = attributeString(childElement, "form:label");
      comboBox.items = [
        ...(comboBox.items ?? []),
        {
          ...(labelAttribute === undefined ? {} : { labelAttribute }),
          label: textOf(childElement),
        },
      ];
    } else if (kind === "listBox" && childElement.name === "form:option") {
      const listBox = control as ListBoxControlOptions;
      const optionPath = `${controlPath}/form:option`;
      assertNoUnknownAttributes(
        childElement,
        ["form:current-selected", "form:selected", "form:label", "form:value"],
        part,
        optionPath,
      );
      const currentSelected = attributeString(childElement, "form:current-selected");
      const selected = attributeString(childElement, "form:selected");
      const labelAttribute = attributeString(childElement, "form:label");
      const value = attributeString(childElement, "form:value");
      listBox.options = [
        ...(listBox.options ?? []),
        {
          ...(currentSelected === undefined
            ? {}
            : {
                currentSelected: parseRequiredBoolean(
                  currentSelected,
                  "form:current-selected",
                  part,
                  optionPath,
                ),
              }),
          ...(selected === undefined
            ? {}
            : {
                selected: parseRequiredBoolean(selected, "form:selected", part, optionPath),
              }),
          ...(labelAttribute === undefined ? {} : { labelAttribute }),
          ...(value === undefined ? {} : { value }),
          label: textOf(childElement),
        },
      ];
    } else if (kind === "grid" && childElement.name === "form:column") {
      const grid = control as GridControlOptions;
      grid.columns = [...(grid.columns ?? []), parseColumn(childElement, part, controlPath, ids)];
    } else {
      throw formError(
        part,
        controlPath,
        childElement.name ?? "unknown",
        "child has no canonical form control mapping",
      );
    }
  }
  return control;
}

function parseColumn(
  element: Element,
  part: string,
  path: string,
  ids: FormIdAllocator,
): GridColumnOptions {
  const columnPath = `${path}/form:column`;
  assertNoUnknownAttributes(
    element,
    ["form:name", "form:control-implementation", "form:label", "form:text-style-name"],
    part,
    columnPath,
  );
  const name = attributeString(element, "form:name");
  const controlImplementation = attributeString(element, "form:control-implementation");
  const label = attributeString(element, "form:label");
  const textStyleName = attributeString(element, "form:text-style-name");
  const column: GridColumnOptions = {
    ...(name === undefined ? {} : { name }),
    ...(controlImplementation === undefined ? {} : { controlImplementation }),
    ...(label === undefined ? {} : { label }),
    ...(textStyleName === undefined ? {} : { textStyleName }),
    controls: [],
  };
  for (const child of element.elements ?? []) {
    if (child.type === "text") continue;
    if (!CONTROL_ELEMENT_KINDS[child.name ?? ""])
      throw formError(
        part,
        columnPath,
        child.name ?? "unknown",
        "element is not a grid column control",
      );
    column.controls.push(
      parseControl(child, part, columnPath, ids, true) as GridFormControlOptions,
    );
  }
  if (!column.controls.length)
    throw formError(part, columnPath, "form:column", "grid column requires at least one control");
  return column;
}

function parseForm(
  element: Element,
  part: string,
  path: string,
  ids: FormIdAllocator,
): FormContainerOptions {
  const formPath = `${path}/form:form`;
  const attributes = parseSpecifiedAttributes(element, FORM_ATTRIBUTES, part, formPath, [
    "xlink:type",
    "xlink:actuate",
  ]);
  const linkType = attributeString(element, "xlink:type");
  if (linkType !== undefined && linkType !== "simple")
    throw formError(part, formPath, "xlink:type", "form link type must be simple");
  const actuate = attributeString(element, "xlink:actuate");
  if (actuate !== undefined && actuate !== "onRequest")
    throw formError(part, formPath, "xlink:actuate", "form link actuation must be onRequest");
  const form = { ...attributes, controls: [], forms: [] } as FormContainerOptions;
  for (const child of element.elements ?? []) {
    if (child.type === "text") continue;
    if (child.name === "form:properties") {
      form.properties = parsePropertyContainer(child, part, formPath);
    } else if (child.name === "office:event-listeners") {
      form.events = parseEventContainer(child, part, formPath);
    } else if (CONTROL_ELEMENT_KINDS[child.name ?? ""]) {
      form.controls.push(parseControl(child, part, formPath, ids));
    } else if (child.name === "form:form") {
      form.forms = [...(form.forms ?? []), parseForm(child, part, formPath, ids)];
    } else if (child.name === "form:connection-resource") {
      const resourcePath = `${formPath}/form:connection-resource`;
      assertNoUnknownAttributes(child, ["xlink:href"], part, resourcePath);
      form.connectionResource = {
        href: requiredAttribute(child, "xlink:href", part, resourcePath),
      };
    } else {
      throw formError(
        part,
        formPath,
        child.name ?? "unknown",
        "child has no canonical form mapping",
      );
    }
  }
  if (!form.forms?.length) delete form.forms;
  return form;
}

export function parseOfficeForms(
  container: Element | undefined,
  part = "content.xml",
  path = "/office:document-content/office:body/office:forms",
): OfficeFormsOptions {
  if (!container) return { containers: [] };
  if (container.name !== "office:forms")
    throw formError(part, path, container.name ?? "unknown", "element is not office:forms");
  assertNoUnknownAttributes(
    container,
    ["form:automatic-focus", "form:apply-design-mode"],
    part,
    path,
  );
  const automaticFocus = attributeString(container, "form:automatic-focus");
  const designMode = attributeString(container, "form:apply-design-mode");
  const ids = new FormIdAllocator(container);
  if (ids.duplicateId !== undefined)
    throw formError(part, path, ids.duplicateId, "duplicate form control id");
  const options: OfficeFormsOptions = { containers: [] };
  for (const [name, value] of [
    ["form:automatic-focus", automaticFocus],
    ["form:apply-design-mode", designMode],
  ] as const) {
    if (value === undefined) continue;
    try {
      if (name === "form:automatic-focus")
        options.automaticFocus = parseAttributeValue(value, "boolean") === true;
      else options.designMode = parseAttributeValue(value, "boolean") === true;
    } catch {
      throw formError(part, path, name, `invalid boolean value: ${value}`);
    }
  }
  for (const child of container.elements ?? []) {
    if (child.type === "text") continue;
    if (child.name === "form:form") {
      options.containers.push(parseForm(child, part, path, ids));
    } else if (child.name === "xforms:model") {
      options.models = [...(options.models ?? []), parseXFormsModel(child, part, path)];
    } else {
      throw formError(
        part,
        path,
        child.name ?? "unknown",
        "child has no canonical office forms mapping",
      );
    }
  }
  return options;
}

function invert(mapping: Record<string, string>): Record<string, string> {
  return Object.fromEntries(Object.entries(mapping).map(([token, field]) => [field, token]));
}

function xmlAttributeValue(
  value: unknown,
  type: AttributeType,
): string | number | boolean | undefined {
  if (value === undefined) return undefined;
  if (typeof value === "boolean" || typeof value === "number") return value;
  const mappings: Record<string, Record<string, string>> = {
    buttonType: invert(BUTTON_TYPES),
    commandType: invert(COMMAND_TYPES),
    listSourceType: invert(LIST_SOURCE_TYPES),
    listLinkageType: invert(LIST_LINKAGE_TYPES),
    visualEffect: invert(VISUAL_EFFECTS),
    imagePosition: invert(IMAGE_POSITIONS),
    imageAlign: invert(IMAGE_ALIGNS),
    navigationMode: invert(NAVIGATION_MODES),
    tabCycle: invert(TAB_CYCLES),
    orientation: invert(ORIENTATIONS),
    triState: invert(TRI_STATES),
  };
  if (typeof value === "string") return mappings[type]?.[value] ?? value;
  throw new Error(`unsupported form attribute value: ${typeof value}`);
}

function propertyValueAttributes(
  value: FormPropertyOptions["value"],
): Record<string, string | number | boolean | undefined> {
  if (value.type === "float" || value.type === "percentage")
    return { "office:value-type": value.type, "office:value": value.value };
  if (value.type === "currency")
    return {
      "office:value-type": value.type,
      "office:value": value.value,
      "office:currency": value.currency,
    };
  if (value.type === "date")
    return { "office:value-type": value.type, "office:date-value": value.value };
  if (value.type === "time")
    return { "office:value-type": value.type, "office:time-value": value.value };
  if (value.type === "boolean")
    return { "office:value-type": value.type, "office:boolean-value": value.value };
  if (value.type === "string")
    return { "office:value-type": value.type, "office:string-value": value.value };
  return { "office:value-type": "void" };
}

function listValueAttributes(
  value: FormListPropertyOptions["value"],
  item: string | number | boolean,
): Record<string, string | number | boolean | undefined> {
  if (value.type === "date") return { "office:date-value": item };
  if (value.type === "time") return { "office:time-value": item };
  if (value.type === "boolean") return { "office:boolean-value": item };
  if (value.type === "string") return { "office:string-value": item };
  return { "office:value": item };
}

function propertiesXml(
  properties: Array<FormPropertyOptions | FormListPropertyOptions> | undefined,
): string {
  if (!properties?.length) return "";
  return xmlElement(
    "form:properties",
    undefined,
    properties.map((property) => {
      if ("values" in property.value) {
        const value = property.value;
        return xmlElement(
          "form:list-property",
          {
            "form:property-name": property.name,
            "office:value-type": value.type,
            ...(value.type === "currency" && value.currency !== undefined
              ? { "office:currency": value.currency }
              : {}),
          },
          value.values.map((item) =>
            xmlElement("form:list-value", listValueAttributes(value, item)),
          ),
        );
      }
      return xmlElement("form:property", {
        "form:property-name": property.name,
        ...propertyValueAttributes(property.value),
      });
    }),
  );
}

function eventXml(event: FormEventListenerOptions): string {
  return xmlElement("script:event-listener", {
    "script:event-name": event.eventName,
    "script:language": event.language,
    ...(event.target === "macro"
      ? { "script:macro-name": event.macroName }
      : {
          "xlink:type": "simple",
          "xlink:href": event.href,
          "xlink:actuate": event.actuate ?? "onRequest",
        }),
  });
}

function eventsXml(events: FormEventListenerOptions[] | undefined): string {
  return events?.length
    ? xmlElement("office:event-listeners", undefined, events.map(eventXml))
    : "";
}

function xFormsElementXml(element: XFormsElementOptions): string {
  return xmlElement(
    element.name,
    element.attributes,
    (element.children ?? []).map((child) =>
      typeof child === "string" ? escapeText(child) : xFormsElementXml(child),
    ),
  );
}

function xFormsModelXml(model: XFormsModelOptions): string {
  return xmlElement(
    "xforms:model",
    model.attributes,
    (model.children ?? []).map((child) =>
      typeof child === "string" ? escapeText(child) : xFormsElementXml(child),
    ),
  );
}

function controlXml(control: FormControlOptions, ids: FormIdAllocator): string {
  const kind = control.control;
  const attributeValues: Record<string, unknown> = { ...control };
  const attributes = Object.fromEntries(
    Object.entries(CONTROL_ATTRIBUTES[kind]).map(([name, spec]) => [
      name,
      xmlAttributeValue(attributeValues[spec.field], spec.type),
    ]),
  );
  if (control.xmlId === undefined) attributes["xml:id"] = ids.allocate(control.id);
  const children = [propertiesXml(control.properties), eventsXml(control.events)];
  if (control.control === "textArea")
    children.push(
      ...(control.paragraphs ?? []).map((paragraph) =>
        xmlElement("text:p", undefined, [escapeText(paragraph)]),
      ),
    );
  if (control.control === "comboBox")
    children.push(
      ...control.items.map((item) =>
        xmlElement(
          "form:item",
          { "form:label": item.labelAttribute },
          item.label === undefined ? [] : [escapeText(item.label)],
        ),
      ),
    );
  if (control.control === "listBox")
    children.push(
      ...control.options.map((option) =>
        xmlElement(
          "form:option",
          {
            "form:current-selected": option.currentSelected,
            "form:selected": option.selected,
            ...(option.labelAttribute === undefined ? {} : { "form:label": option.labelAttribute }),
            ...(option.value === undefined ? {} : { "form:value": option.value }),
          },
          option.label === undefined ? [] : [escapeText(option.label)],
        ),
      ),
    );
  if (control.control === "grid")
    children.push(...control.columns.map((column) => columnXml(column, ids)));
  return xmlElement(KIND_ELEMENTS[kind], attributes, children.filter(Boolean));
}

function columnXml(column: GridColumnOptions, ids: FormIdAllocator): string {
  return xmlElement(
    "form:column",
    {
      "form:name": column.name,
      "form:control-implementation": column.controlImplementation,
      "form:label": column.label,
      "form:text-style-name": column.textStyleName,
    },
    column.controls.map((control) => controlXml(control, ids)),
  );
}

function formXml(form: FormContainerOptions, ids: FormIdAllocator): string {
  return xmlElement(
    "form:form",
    {
      "form:name": form.name,
      "form:control-implementation": form.controlImplementation,
      ...(form.href === undefined ? {} : { "xlink:type": "simple", "xlink:href": form.href }),
      "office:target-frame": form.targetFrame,
      "form:method": form.method,
      "form:enctype": form.enctype,
      "form:allow-deletes": form.allowDeletes,
      "form:allow-inserts": form.allowInserts,
      "form:allow-updates": form.allowUpdates,
      "form:apply-filter": form.applyFilter,
      "form:command-type": form.commandType,
      "form:command": form.command,
      "form:datasource": form.datasource,
      "form:master-fields": form.masterFields,
      "form:detail-fields": form.detailFields,
      "form:escape-processing": form.escapeProcessing,
      "form:filter": form.filter,
      "form:ignore-result": form.ignoreResult,
      "form:navigation-mode": form.navigationMode,
      "form:order": form.order,
      "form:tab-cycle": form.tabCycle,
    },
    [
      propertiesXml(form.properties),
      eventsXml(form.events),
      ...form.controls.map((control) => controlXml(control, ids)),
      ...(form.forms ?? []).map((nested) => formXml(nested, ids)),
      ...(form.connectionResource
        ? [xmlElement("form:connection-resource", { "xlink:href": form.connectionResource.href })]
        : []),
    ].filter(Boolean),
  );
}

function collectControlIds(control: FormControlOptions, ids: Set<string>): void {
  if (control.xmlId !== undefined) {
    if (ids.has(control.xmlId)) throw new Error("duplicate form control id");
    ids.add(control.xmlId);
  }
  if (control.control === "grid")
    for (const column of control.columns)
      for (const nested of column.controls) collectControlIds(nested, ids);
}

function collectStringifyIds(forms: FormContainerOptions[], ids: Set<string>): void {
  for (const form of forms) {
    for (const control of form.controls) collectControlIds(control, ids);
    collectStringifyIds(form.forms ?? [], ids);
  }
}

class StringifyIdAllocator extends FormIdAllocator {
  constructor(forms: FormContainerOptions[]) {
    super();
    collectStringifyIds(forms, this.used);
  }
}

export function formsXml(containers: FormContainerOptions[] | undefined): string {
  if (!containers?.length) return "";
  const ids = new StringifyIdAllocator(containers);
  return containers.map((form) => formXml(form, ids)).join("");
}

export function officeFormsXml(options: OfficeFormsOptions | undefined): string {
  if (
    options === undefined ||
    (!options.containers.length &&
      !options.models?.length &&
      options.automaticFocus === undefined &&
      options.designMode === undefined)
  )
    return "";
  const ids = new StringifyIdAllocator(options?.containers ?? []);
  return xmlElement(
    "office:forms",
    {
      "form:automatic-focus": options?.automaticFocus,
      "form:apply-design-mode": options?.designMode,
    },
    [
      ...(options?.containers ?? []).map((form) => formXml(form, ids)),
      ...(options?.models ?? []).map(xFormsModelXml),
    ],
  );
}
