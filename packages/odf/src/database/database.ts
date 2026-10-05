import type { Element } from "@office-open/xml";

import { generateOcf, readOcf, readXml } from "../container/package";
import { OdfSchemaError } from "../errors";
import { formsXml, parseOfficeForms, type OfficeFormsOptions } from "../forms/forms";
import { metaXml, ODF_NAMESPACES, parseMeta } from "../meta/core-properties";
import { childNamed, childrenNamed, textOf, xmlElement } from "../runtime/xml";

const MIME = "application/vnd.oasis.opendocument.database";
const NAMESPACES = `${ODF_NAMESPACES} xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:xforms="http://www.w3.org/2002/xforms"`;

export const DATABASE_ELEMENT_NAMES = [
  "db:application-connection-settings",
  "db:auto-increment",
  "db:character-set",
  "db:column",
  "db:column-definition",
  "db:column-definitions",
  "db:columns",
  "db:component",
  "db:component-collection",
  "db:connection-data",
  "db:connection-resource",
  "db:data-source",
  "db:data-source-setting",
  "db:data-source-setting-value",
  "db:data-source-settings",
  "db:database-description",
  "db:delimiter",
  "db:driver-settings",
  "db:file-based-database",
  "db:filter-statement",
  "db:forms",
  "db:index",
  "db:index-column",
  "db:index-columns",
  "db:indices",
  "db:key",
  "db:key-column",
  "db:key-columns",
  "db:keys",
  "db:login",
  "db:order-statement",
  "db:queries",
  "db:query",
  "db:query-collection",
  "db:reports",
  "db:schema-definition",
  "db:server-database",
  "db:table-definition",
  "db:table-definitions",
  "db:table-exclude-filter",
  "db:table-filter",
  "db:table-filter-pattern",
  "db:table-include-filter",
  "db:table-representation",
  "db:table-representations",
  "db:table-setting",
  "db:table-settings",
  "db:table-type",
  "db:table-type-filter",
  "db:update-table",
] as const;

export const DATABASE_ATTRIBUTE_NAMES = [
  "db:additional-column-statement",
  "db:append-table-alias-name",
  "db:apply-command",
  "db:as-template",
  "db:base-dn",
  "db:boolean-comparison-mode",
  "db:catalog-name",
  "db:command",
  "db:data-source-setting-is-list",
  "db:data-source-setting-name",
  "db:data-source-setting-type",
  "db:data-type",
  "db:database-name",
  "db:decimal",
  "db:default-cell-style-name",
  "db:default-row-style-name",
  "db:delete-rule",
  "db:description",
  "db:enable-sql92-check",
  "db:encoding",
  "db:escape-processing",
  "db:extension",
  "db:field",
  "db:hostname",
  "db:ignore-driver-privileges",
  "db:is-ascending",
  "db:is-autoincrement",
  "db:is-clustered",
  "db:is-empty-allowed",
  "db:is-first-row-header-line",
  "db:is-nullable",
  "db:is-password-required",
  "db:is-table-name-length-limited",
  "db:is-unique",
  "db:local-socket",
  "db:login-timeout",
  "db:max-row-count",
  "db:media-type",
  "db:name",
  "db:parameter-name-substitution",
  "db:port",
  "db:precision",
  "db:referenced-table-name",
  "db:related-column-name",
  "db:row-retrieving-statement",
  "db:scale",
  "db:schema-name",
  "db:show-deleted",
  "db:string",
  "db:style-name",
  "db:suppress-version-columns",
  "db:system-driver-settings",
  "db:thousand",
  "db:title",
  "db:type",
  "db:type-name",
  "db:update-rule",
  "db:use-catalog",
  "db:use-system-user",
  "db:user-name",
  "db:visible",
  "xlink:actuate",
  "xlink:href",
  "xlink:show",
  "xlink:type",
] as const;

export type DatabaseElementName = (typeof DATABASE_ELEMENT_NAMES)[number];
export type DatabaseAttributeName = (typeof DATABASE_ATTRIBUTE_NAMES)[number];
export type DatabaseAttributeValue = string | number | boolean | undefined;
export type DatabaseAttributes = Partial<Record<DatabaseAttributeName, DatabaseAttributeValue>>;

export interface OdfDatabaseElement {
  name: DatabaseElementName;
  attributes?: DatabaseAttributes;
  children?: OdfDatabaseElement[];
  text?: string;
  forms?: OfficeFormsOptions;
}

export interface DatabaseDocumentOptions {
  title?: string;
  dataSource: OdfDatabaseElement;
}

export const DATABASE_SCHEMA_ELEMENTS: readonly {
  name: DatabaseElementName;
  attributes: readonly DatabaseAttributeName[];
}[] = DATABASE_ELEMENT_NAMES.map((name) => ({ name, attributes: DATABASE_ATTRIBUTE_NAMES }));

const DATABASE_ATTRIBUTES = new Set<string>(DATABASE_ATTRIBUTE_NAMES);

function databaseError(part: string, path: string, name: string, reason: string): OdfSchemaError {
  return new OdfSchemaError(
    `${part}: ${path}/${name}: ${reason}`,
    part,
    `${path}/${name}`,
    name,
    reason,
  );
}

function assertDatabaseAttributes(
  element: Element,
  part: string,
  path: string,
): DatabaseAttributes {
  const attributes: DatabaseAttributes = {};
  for (const [name, value] of Object.entries(element.attributes ?? {})) {
    if (!DATABASE_ATTRIBUTES.has(name)) {
      throw databaseError(part, path, name, "attribute has no canonical database model mapping");
    }
    attributes[name as DatabaseAttributeName] = value;
  }
  return attributes;
}

function parseDatabaseElement(element: Element, part: string, path: string): OdfDatabaseElement {
  if (!DATABASE_ELEMENT_NAMES.includes(element.name as DatabaseElementName)) {
    throw databaseError(
      part,
      path,
      element.name ?? "unknown",
      "element has no canonical database model mapping",
    );
  }
  const elementPath = `${path}/${element.name}`;
  const options: OdfDatabaseElement = {
    name: element.name as DatabaseElementName,
    attributes: assertDatabaseAttributes(element, part, path),
    text:
      element.name === "db:database-description" || element.name === "db:data-source-setting-value"
        ? textOf(element)
        : undefined,
  };
  if (element.name === "db:forms") {
    options.forms = parseOfficeForms(element, part, elementPath);
    for (const child of element.elements ?? []) {
      if (child.type === "element" && child.name !== "form:form") {
        throw databaseError(
          part,
          elementPath,
          child.name ?? "unknown",
          "child has no canonical database form mapping",
        );
      }
    }
    return options;
  }
  const children = (element.elements ?? []).filter(
    (child): child is Element => child.type === "element",
  );
  for (const child of children) {
    if (
      element.name === "db:driver-settings" &&
      child.name === "db:table-filter" &&
      child.elements?.some(
        (nested) =>
          nested.type === "element" &&
          !["db:table-include-filter", "db:table-exclude-filter", "db:table-type-filter"].includes(
            nested.name ?? "",
          ),
      )
    ) {
      throw databaseError(
        part,
        elementPath,
        child.name!,
        "table filter contains an unsupported child",
      );
    }
    options.children = [
      ...(options.children ?? []),
      parseDatabaseElement(child, part, elementPath),
    ];
  }
  return options;
}

function databaseElementXml(element: OdfDatabaseElement): string {
  if (element.name === "db:forms")
    return xmlElement(element.name, element.attributes, [formsXml(element.forms)]);
  const children = [
    ...(element.text !== undefined ? [element.text] : []),
    ...(element.children ?? []).map(databaseElementXml),
  ];
  return xmlElement(element.name, element.attributes, children);
}

export function generateDatabaseDocument(options: DatabaseDocumentOptions): Uint8Array {
  const dataSourceXml = databaseElementXml(options.dataSource);
  return generateOcf(MIME, {
    "content.xml": databaseContentXml(dataSourceXml),
    "meta.xml": metaXml({ title: options.title }),
  });
}

export function parseDatabaseDocument(data: Uint8Array): DatabaseDocumentOptions {
  const { files } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:database");
  const dataSourceElement = childrenNamed(body, "db:data-source")[0];
  if (!dataSourceElement) {
    throw databaseError(
      "content.xml",
      "/office:document-content/office:body/office:database",
      "db:data-source",
      "database body requires a data source",
    );
  }
  for (const child of body?.elements ?? []) {
    if (child.type === "element" && child.name !== "db:data-source") {
      throw databaseError(
        "content.xml",
        "/office:document-content/office:body/office:database",
        child.name ?? "unknown",
        "database body has no canonical mapping",
      );
    }
  }
  return {
    ...parseMeta(files),
    dataSource: parseDatabaseElement(
      dataSourceElement,
      "content.xml",
      "/office:document-content/office:body/office:database",
    ),
  };
}

function databaseContentXml(dataSource: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:body><office:database>${dataSource}</office:database></office:body></office:document-content>`;
}
