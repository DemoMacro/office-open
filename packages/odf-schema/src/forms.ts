import {
  attributeString,
  childNamed,
  childrenNamed,
  escapeText,
  textOf,
  xmlElement,
} from "@office-open/ocf";
import type { Element } from "@office-open/xml";

import { OdfSchemaError } from "./error";

export const FORM_ELEMENT_NAMES = [
  "form:form",
  "form:connection-resource",
  "form:text",
  "form:textarea",
  "form:formatted-text",
  "form:number",
  "form:date",
  "form:time",
  "form:combobox",
  "form:listbox",
  "form:checkbox",
  "form:password",
  "form:file",
  "form:fixed-text",
  "form:button",
  "form:image",
  "form:radio",
  "form:frame",
  "form:image-frame",
  "form:hidden",
  "form:grid",
  "form:value-range",
  "form:generic-control",
  "form:column",
  "form:properties",
  "form:property",
  "form:list-property",
  "form:list-value",
  "form:option",
  "form:item",
  "script:event-listener",
  "xforms:model",
] as const;

export const FORM_ATTRIBUTE_NAMES = [
  "form:id",
  "form:name",
  "form:control-implementation",
  "form:bound-column",
  "form:button-type",
  "form:convert-empty-to-null",
  "form:current-value",
  "form:data-field",
  "form:delay-for-repeat",
  "form:disabled",
  "form:echo-char",
  "form:multi-line",
  "form:default-button",
  "form:toggle",
  "form:focus-on-click",
  "form:xforms-submission",
  "form:max-value",
  "form:min-value",
  "form:step-size",
  "form:page-step-size",
  "form:orientation",
  "form:image-data",
  "form:image-align",
  "form:image-position",
  "form:label",
  "form:list-linkage-type",
  "form:list-source",
  "form:list-source-type",
  "form:automatic-focus",
  "form:apply-design-mode",
  "form:selected",
  "form:size",
  "form:current-state",
  "form:is-tristate",
  "form:state",
  "form:linked-cell",
  "form:spin-button",
  "form:repeat",
  "form:auto-complete",
  "form:multiple",
  "form:xforms-list-source",
  "form:validation",
  "form:property-name",
  "form:value",
  "form:method",
  "form:enctype",
  "form:allow-deletes",
  "form:allow-inserts",
  "form:allow-updates",
  "form:apply-filter",
  "form:command-type",
  "form:command",
  "form:datasource",
  "form:master-fields",
  "form:detail-fields",
  "form:escape-processing",
  "form:filter",
  "form:ignore-result",
  "form:navigation-mode",
  "form:order",
  "form:tab-cycle",
  "form:text-style-name",
  "office:target-frame",
  "office:value-type",
  "office:value",
  "office:currency",
  "office:date-value",
  "office:time-value",
  "office:boolean-value",
  "office:string-value",
  "script:event-name",
  "script:language",
  "script:macro-name",
  "svg:x",
  "svg:y",
  "svg:width",
  "svg:height",
  "xlink:type",
  "xlink:href",
  "xlink:actuate",
  "xforms:bind",
] as const;

export type FormElementName = (typeof FORM_ELEMENT_NAMES)[number];
export type FormAttributeName = (typeof FORM_ATTRIBUTE_NAMES)[number];
export type FormAttributeValue = string | number | boolean | undefined;
export type FormAttributes = Partial<Record<FormAttributeName, FormAttributeValue>>;

const CONTROL_NAMES = FORM_ELEMENT_NAMES.filter(
  (name) =>
    name.startsWith("form:") &&
    ![
      "form:form",
      "form:connection-resource",
      "form:column",
      "form:properties",
      "form:property",
      "form:list-property",
      "form:list-value",
      "form:option",
      "form:item",
    ].includes(name) &&
    name !== "form:image-frame",
) as Extract<FormElementName, `form:${string}`>[];

export type FormControlName = (typeof CONTROL_NAMES)[number];

export interface FormEventOptions {
  eventName: string;
  language: string;
  macroName?: string;
  href?: string;
  actuate?: "onRequest";
}

export interface FormPropertyValueOptions {
  type: "float" | "percentage" | "currency" | "date" | "time" | "boolean" | "string";
  value?: string | number | boolean;
  currency?: string;
  listValues?: Array<string | number | boolean>;
}

export interface FormControlOptions {
  kind: FormControlName;
  id?: string;
  name?: string;
  attributes?: FormAttributes;
  properties?: Array<{ name: string; value: FormPropertyValueOptions; list?: boolean }>;
  events?: FormEventOptions[];
  options?: Array<{ attributes?: FormAttributes; label?: string; items?: FormAttributes[] }>;
  columns?: Array<{
    attributes?: FormAttributes;
    controls: FormControlOptions[];
    properties?: FormControlOptions["properties"];
    events?: FormEventOptions[];
  }>;
  children?: FormControlOptions[];
  xformsModel?: XFormsModelOptions;
}

export interface XFormsModelOptions {
  attributes: Record<string, string | number | boolean>;
  children?: Array<{
    name: string;
    attributes?: Record<string, string | number | boolean>;
    children?: unknown[];
    text?: string;
  }>;
}

export interface OdfFormOptions {
  name?: string;
  attributes?: FormAttributes;
  controls: FormControlOptions[];
  properties?: FormControlOptions["properties"];
  events?: FormEventOptions[];
  connectionResource?: { href: string; actuate?: "onRequest" };
}

export interface OfficeFormsOptions {
  forms: OdfFormOptions[];
}

export const FORM_SCHEMA_ELEMENTS: readonly {
  name: FormElementName;
  attributes: readonly FormAttributeName[];
}[] = [
  {
    name: "form:form",
    attributes: [
      "form:name",
      "form:control-implementation",
      "form:method",
      "form:enctype",
      "form:datasource",
      "form:command",
      "form:filter",
      "form:order",
      "form:tab-cycle",
      "xlink:href",
      "office:target-frame",
    ],
  },
  { name: "form:connection-resource", attributes: ["xlink:href", "xlink:actuate"] },
  ...CONTROL_NAMES.map((name) => ({ name, attributes: FORM_ATTRIBUTE_NAMES })),
  { name: "form:image-frame", attributes: FORM_ATTRIBUTE_NAMES },
  { name: "form:column", attributes: FORM_ATTRIBUTE_NAMES },
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
  { name: "form:option", attributes: FORM_ATTRIBUTE_NAMES },
  { name: "form:item", attributes: FORM_ATTRIBUTE_NAMES },
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

const ATTRIBUTE_SET = new Set<string>(FORM_ATTRIBUTE_NAMES);

function formError(part: string, path: string, name: string, reason: string): OdfSchemaError {
  return new OdfSchemaError(
    `${part}: ${path}/${name}: ${reason}`,
    part,
    `${path}/${name}`,
    name,
    reason,
  );
}

function normalizeAttributeValue(value: string): string | number | boolean {
  if (value === "true") return true;
  if (value === "false") return false;
  if (/^-?\d+(?:\.\d+)?$/.test(value)) return Number(value);
  return value;
}

function assertAttributes(
  element: Element | undefined,
  part: string,
  path: string,
  _name: string,
): FormAttributes {
  const result: FormAttributes = {};
  for (const [qualifiedName, value] of Object.entries(element?.attributes ?? {})) {
    if (!ATTRIBUTE_SET.has(qualifiedName)) {
      throw formError(part, path, qualifiedName, "attribute has no canonical form model mapping");
    }
    result[qualifiedName as FormAttributeName] = normalizeAttributeValue(String(value));
  }
  return result;
}

function parseProperty(
  element: Element,
  part: string,
  path: string,
): NonNullable<FormControlOptions["properties"]>[number] {
  const attributes = assertAttributes(element, part, path, element.name!);
  const value: FormPropertyValueOptions = {
    type: (attributeString(element, "office:value-type") ??
      "string") as FormPropertyValueOptions["type"],
    value:
      attributes["office:value"] ??
      attributes["office:string-value"] ??
      attributes["office:boolean-value"] ??
      attributes["office:date-value"] ??
      attributes["office:time-value"],
    currency: attributes["office:currency"] as string | undefined,
    listValues: childrenNamed(element, "form:list-value").map((child) => {
      const childAttributes = assertAttributes(child, part, path, "form:list-value");
      return (
        childAttributes["office:value"] ??
        childAttributes["office:string-value"] ??
        childAttributes["office:boolean-value"] ??
        childAttributes["office:date-value"] ??
        childAttributes["office:time-value"] ??
        ""
      );
    }),
  };
  return {
    name: attributeString(element, "form:property-name") ?? "",
    value,
    list: element.name === "form:list-property",
  };
}

function parseEvent(element: Element, part: string, path: string): FormEventOptions {
  assertAttributes(element, part, path, element.name!);
  return {
    eventName: attributeString(element, "script:event-name") ?? "",
    language: attributeString(element, "script:language") ?? "",
    macroName: attributeString(element, "script:macro-name"),
    href: attributeString(element, "xlink:href"),
    actuate: attributeString(element, "xlink:actuate") as FormEventOptions["actuate"],
  };
}

function parseProperties(
  element: Element,
  part: string,
  path: string,
): FormControlOptions["properties"] {
  return childrenNamed(element, "form:property")
    .concat(childrenNamed(element, "form:list-property"))
    .map((property) => parseProperty(property, part, path));
}

function parseEvents(element: Element, part: string, path: string): FormEventOptions[] {
  return childrenNamed(element, "script:event-listener").map((event) =>
    parseEvent(event, part, path),
  );
}

function parseControl(element: Element, part: string, path: string): FormControlOptions {
  const controlPath = `${path}/${element.name}`;
  const attributes = assertAttributes(element, part, path, element.name!);
  const control: FormControlOptions = {
    kind: element.name as FormControlName,
    id: attributes["form:id"] as string | undefined,
    name: attributes["form:name"] as string | undefined,
    attributes,
  };
  for (const child of element.elements ?? []) {
    if (child.type === "text") continue;
    if (child.name === "form:properties")
      control.properties = parseProperties(child, part, controlPath);
    else if (child.name === "script:event-listener")
      control.events = [...(control.events ?? []), parseEvent(child, part, controlPath)];
    else if (child.name === "form:option" || child.name === "form:item") {
      const option = {
        attributes: assertAttributes(child, part, controlPath, child.name!),
        label: textOf(child),
      };
      control.options = [...(control.options ?? []), option];
    } else if (child.name === "form:column") {
      const column = {
        attributes: assertAttributes(child, part, controlPath, child.name!),
        controls: (child.elements ?? [])
          .filter((nested) => CONTROL_NAMES.includes(nested.name as FormControlName))
          .map((nested) => parseControl(nested, part, controlPath)),
        properties: childNamed(child, "form:properties")
          ? parseProperties(childNamed(child, "form:properties")!, part, controlPath)
          : undefined,
        events: parseEvents(child, part, controlPath),
      };
      control.columns = [...(control.columns ?? []), column];
    } else if (child.name === "xforms:model") {
      control.xformsModel = parseXFormsModel(child, part, controlPath);
    } else if (CONTROL_NAMES.includes(child.name as FormControlName)) {
      control.children = [...(control.children ?? []), parseControl(child, part, controlPath)];
    } else {
      throw formError(
        part,
        controlPath,
        child.name ?? "text",
        "child has no canonical form control mapping",
      );
    }
  }
  return control;
}

export function parseXFormsModel(
  element: Element,
  _part: string,
  _path: string,
): XFormsModelOptions {
  return {
    attributes: { ...element.attributes } as Record<string, string | number | boolean>,
    children: (element.elements ?? [])
      .filter((child) => child.type === "element")
      .map((child) => ({
        name: child.name!,
        attributes: {
          ...child.attributes,
        } as Record<string, string | number | boolean>,
        children: child.elements ?? [],
        text: textOf(child),
      })),
  };
}

export function parseOfficeForms(
  container: Element | undefined,
  part = "content.xml",
  path = "/office:document-content/office:body/office:forms",
): OfficeFormsOptions {
  const forms: OdfFormOptions[] = childrenNamed(container, "form:form").map((element) => {
    const formPath = `${path}/form:form`;
    const attributes = assertAttributes(element, part, path, "form:form");
    const controls: FormControlOptions[] = [];
    let properties: FormControlOptions["properties"];
    let events: FormEventOptions[] = [];
    let connectionResource: OdfFormOptions["connectionResource"];
    for (const child of element.elements ?? []) {
      if (child.type === "text") continue;
      if (CONTROL_NAMES.includes(child.name as FormControlName))
        controls.push(parseControl(child, part, formPath));
      else if (child.name === "form:properties")
        properties = parseProperties(child, part, formPath);
      else if (child.name === "script:event-listener")
        events.push(parseEvent(child, part, formPath));
      else if (child.name === "form:connection-resource") {
        assertAttributes(child, part, formPath, child.name!);
        connectionResource = {
          href: attributeString(child, "xlink:href") ?? "",
          actuate: attributeString(child, "xlink:actuate") as "onRequest" | undefined,
        };
      } else
        throw formError(
          part,
          formPath,
          child.name ?? "text",
          "child has no canonical form mapping",
        );
    }
    return {
      name: attributes["form:name"] as string | undefined,
      attributes,
      controls,
      properties,
      events,
      connectionResource,
    };
  });
  return { forms };
}

function propertyAttributes(
  property: NonNullable<FormControlOptions["properties"]>[number],
): FormAttributes {
  const value = property.value;
  return {
    "form:property-name": property.name,
    "office:value-type": value.type,
    "office:value":
      value.type === "float" || value.type === "percentage" || value.type === "currency"
        ? value.value
        : undefined,
    "office:currency": value.currency,
    "office:date-value": value.type === "date" ? value.value : undefined,
    "office:time-value": value.type === "time" ? value.value : undefined,
    "office:boolean-value": value.type === "boolean" ? value.value : undefined,
    "office:string-value": value.type === "string" ? value.value : undefined,
  };
}

function eventXml(event: FormEventOptions): string {
  return xmlElement("script:event-listener", {
    "script:event-name": event.eventName,
    "script:language": event.language,
    "script:macro-name": event.macroName,
    "xlink:href": event.href,
    "xlink:actuate": event.href ? (event.actuate ?? "onRequest") : undefined,
  });
}

function propertiesXml(
  properties: NonNullable<FormControlOptions["properties"]> | undefined,
): string[] {
  if (!properties?.length) return [];
  return [
    xmlElement(
      "form:properties",
      undefined,
      properties.flatMap((property) => {
        const element = property.list ? "form:list-property" : "form:property";
        const listValues =
          property.value.listValues?.map((value) =>
            xmlElement("form:list-value", {
              "office:value": value,
              "office:string-value": typeof value === "string" ? value : undefined,
              "office:boolean-value": typeof value === "boolean" ? value : undefined,
              "office:date-value":
                typeof value === "string" && property.value.type === "date" ? value : undefined,
              "office:time-value":
                typeof value === "string" && property.value.type === "time" ? value : undefined,
            }),
          ) ?? [];
        return [xmlElement(element, propertyAttributes(property), listValues)];
      }),
    ),
  ];
}

function xformsXml(model: XFormsModelOptions): string {
  const children = (model.children ?? []).map((child) => {
    const attributes = Object.fromEntries(
      Object.entries(child.attributes ?? {}).map(([key, value]) => [key, String(value)]),
    );
    const descendants =
      (child.children as Element[] | undefined)?.map((descendant) => ({
        name: descendant.name ?? "",
        attributes: Object.fromEntries(
          Object.entries(descendant.attributes ?? {}).map(([key, value]) => [key, String(value)]),
        ),
        children: [],
        text: textOf(descendant),
      })) ?? [];
    return xmlElement(child.name, attributes, descendantXml(descendants, child.text));
  });
  const attributes = Object.fromEntries(
    Object.entries(model.attributes).map(([key, value]) => [key, String(value)]),
  );
  return xmlElement("xforms:model", attributes, children);
}

function descendantXml(children: XFormsModelOptions["children"], text?: string): string[] {
  const elements = (children ?? []).map((child) =>
    xmlElement(
      child.name,
      Object.fromEntries(
        Object.entries(child.attributes ?? {}).map(([key, value]) => [key, String(value)]),
      ),
      descendantXml(child.children as XFormsModelOptions["children"], child.text),
    ),
  );
  return text ? [escapeText(text), ...elements] : elements;
}

function controlXml(control: FormControlOptions): string {
  const children = [
    ...(control.properties ? propertiesXml(control.properties) : []),
    ...(control.events ?? []).map(eventXml),
    ...(control.options ?? []).map((option) =>
      xmlElement("form:option", option.attributes, option.label ? [escapeText(option.label)] : []),
    ),
    ...(control.columns ?? []).map((column) =>
      xmlElement("form:column", column.attributes, [
        ...(column.properties ? propertiesXml(column.properties) : []),
        ...(column.events ?? []).map(eventXml),
        ...column.controls.map(controlXml),
      ]),
    ),
    ...(control.children ?? []).map(controlXml),
    ...(control.xformsModel ? [xformsXml(control.xformsModel)] : []),
  ];
  return xmlElement(
    control.kind,
    { "form:id": control.id, "form:name": control.name, ...control.attributes },
    children,
  );
}

export function formsXml(options: OfficeFormsOptions | undefined): string {
  return (options?.forms ?? [])
    .map((form) => {
      const children = [
        ...(form.properties ? propertiesXml(form.properties) : []),
        ...(form.events ?? []).map(eventXml),
        ...form.controls.map(controlXml),
        ...(form.connectionResource
          ? [
              xmlElement("form:connection-resource", {
                "xlink:href": form.connectionResource.href,
                "xlink:actuate": form.connectionResource.actuate,
              }),
            ]
          : []),
      ];
      return xmlElement("form:form", { "form:name": form.name, ...form.attributes }, children);
    })
    .join("");
}

export function officeFormsXml(options: OfficeFormsOptions | undefined): string {
  if (!options?.forms.length) return "";
  return xmlElement("office:forms", undefined, [formsXml(options)]);
}
