/**
 * Excel rich-data part types and descriptors.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import {
  attr,
  attrBool,
  attrNum,
  children,
  escapeXml,
  findChild,
  stringifyElement,
  textOf,
} from "@office-open/xml";
import type { Element } from "@office-open/xml";

const RD_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
const RD2_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
const REL_NS = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";

/** Foreign extension carried by an otherwise typed rich-data container. */
export interface RichDataExtensionOptions {
  uri: string;
  innerXml?: string;
}

/** A rich-value fallback (CT_RichValueFallback). */
export interface RichValueFallbackOptions {
  type?: string;
  text?: string;
}

/** One rich value (CT_RichValue). */
export interface RichValueOptions {
  /** Structure index (required). */
  structureIndex: number;
  fallback?: RichValueFallbackOptions;
  /** Ordered primitive values referenced by the structure keys. */
  values: string[];
}

/** xl/richData/rdrichvalue.xml (CT_RichValueData). */
export interface RichValueDataOptions {
  count?: number;
  values?: RichValueOptions[];
  extensions?: RichDataExtensionOptions[];
}

/** Structure key (CT_Key). */
export interface RichValueStructureKeyOptions {
  name: string;
  type?: "d" | "b" | "e" | "s" | "i";
}

/** One rich-value structure (CT_RichValueStructure). */
export interface RichValueStructureOptions {
  type: string;
  keys: RichValueStructureKeyOptions[];
}

/** xl/richData/rdrichvaluestructure.xml (CT_RichValueStructures). */
export interface RichValueStructuresOptions {
  count?: number;
  structures?: RichValueStructureOptions[];
  extensions?: RichDataExtensionOptions[];
}

/** Reserved-key flag (CT_RichValueTypeReservedKeyFlag). */
export interface RichValueTypeFlagOptions {
  name: string;
  value: boolean;
}

/** Reserved key (CT_RichValueTypeReservedKey). */
export interface RichValueTypeKeyOptions {
  name: string;
  flags: RichValueTypeFlagOptions[];
}

/** Key flags on a global or named rich-value type. */
export interface RichValueTypeKeyFlagsOptions {
  keys: RichValueTypeKeyOptions[];
}

/** Named rich-value type (CT_RichValueType). */
export interface RichValueTypeOptions {
  name: string;
  keyFlags?: RichValueTypeKeyFlagsOptions;
  extensions?: RichDataExtensionOptions[];
}

/** Named type collection (CT_RichValueTypes). */
export interface RichValueTypesOptions {
  types: RichValueTypeOptions[];
}

/** Global type metadata (CT_RichValueGlobalType). */
export interface RichValueGlobalTypeOptions {
  keyFlags?: RichValueTypeKeyFlagsOptions;
  extensions?: RichDataExtensionOptions[];
}

/** xl/richData/rdRichValueTypes.xml (CT_RichValueTypesInfo). */
export interface RichValueTypesInfoOptions {
  global?: RichValueGlobalTypeOptions;
  types?: RichValueTypesOptions;
  extensions?: RichDataExtensionOptions[];
}

/** Relationship reference in xl/richData/richValueRel.xml (CT_RichValueRelRelationship). */
export interface RichValueRelOptions {
  /** Relationship id in the richValueRel part's own .rels (required). */
  rId: string;
}

/** xl/richData/richValueRel.xml (CT_RichValueRels). */
export interface RichValueRelsOptions {
  relationships: RichValueRelOptions[];
  extensions?: RichDataExtensionOptions[];
}

/** Four typed parts backing Excel pictures and other rich values. */
export interface RichDataPackageOptions {
  data?: RichValueDataOptions;
  structures?: RichValueStructuresOptions;
  types?: RichValueTypesInfoOptions;
  relationships?: RichValueRelsOptions;
  /** Source package paths; defaults follow Excel's richData layout. Round-trip only. */
  dataPath?: string;
  structuresPath?: string;
  typesPath?: string;
  relationshipsPath?: string;
  /** Source workbook relationship ids. Round-trip only. */
  dataRelationshipId?: string;
  structuresRelationshipId?: string;
  typesRelationshipId?: string;
  relationshipsRelationshipId?: string;
  /** Source relationships owned by richValueRel.xml. Round-trip only. */
  partRelationships?: RichDataPartRelationshipOptions[];
}

/** One package relationship owned by richValueRel.xml. */
export interface RichDataPartRelationshipOptions {
  rId: string;
  relationshipType: string;
  target: string;
  targetMode?: "External";
}

/** Parse the companion .rels for richValueRel.xml with all OPC attributes. */
export function parseRichDataPartRelationships(
  el: Element | undefined,
): RichDataPartRelationshipOptions[] {
  return (el?.elements ?? []).flatMap((element) => {
    if (element.name !== "Relationship") return [];
    const rId = attr(element, "Id");
    const relationshipType = attr(element, "Type");
    const target = attr(element, "Target");
    if (rId === undefined || relationshipType === undefined || target === undefined) {
      throw new RichDataParseError(
        "richData/_rels",
        "/Relationships/Relationship",
        "Relationship",
        "missing Id, Type, or Target",
      );
    }
    return [
      {
        rId,
        relationshipType,
        target,
        ...(attr(element, "TargetMode") === "External" ? { targetMode: "External" as const } : {}),
      },
    ];
  });
}

export class RichDataParseError extends Error {
  constructor(
    readonly part: string,
    readonly path: string,
    readonly name: string,
    readonly reason: string,
  ) {
    super(`${part}${path}: ${name}: ${reason}`);
  }
}

function extensionList(el: Element, part: string, path: string): RichDataExtensionOptions[] {
  const list =
    children(el, "xlrd:extLst").at(0) ??
    children(el, "xlrd2:extLst").at(0) ??
    children(el, "xlrvrel:extLst").at(0);
  if (!list) return [];
  return (list.elements ?? []).map((element) => {
    if (!element.name?.endsWith(":ext") && element.name !== "ext") {
      throw new RichDataParseError(
        part,
        `${path}/extLst`,
        element.name ?? "unknown",
        "unexpected extension",
      );
    }
    const uri = attr(element, "uri");
    if (uri === undefined) {
      throw new RichDataParseError(part, `${path}/extLst/ext`, "@uri", "missing required value");
    }
    return { uri, innerXml: (element.elements ?? []).map(stringifyElement).join("") };
  });
}

function extensionXml(extensions: RichDataExtensionOptions[] | undefined, prefix: string): string {
  if (extensions === undefined) return "";
  return `<${prefix}:extLst>${(extensions ?? [])
    .map(
      (extension) =>
        `<${prefix}:ext uri="${escapeXml(extension.uri)}">${extension.innerXml ?? ""}</${prefix}:ext>`,
    )
    .join("")}</${prefix}:extLst>`;
}

function parseKeyFlags(el: Element, part: string, path: string): RichValueTypeKeyFlagsOptions {
  const element = findChild(el, "xlrd2:keyFlags");
  if (!element) throw new RichDataParseError(part, path, "keyFlags", "missing required element");
  rejectUnknownChildren(element, ["xlrd2:key"], part, `${path}/keyFlags`);
  rejectUnknownAttributes(element, [], part, `${path}/keyFlags`);
  const keys = children(element, "xlrd2:key").map((key) => {
    rejectUnknownChildren(key, ["xlrd2:flag"], part, `${path}/keyFlags/key`);
    rejectUnknownAttributes(key, ["name"], part, `${path}/keyFlags/key`);
    const name = attr(key, "name");
    if (name === undefined) {
      throw new RichDataParseError(part, `${path}/keyFlags/key`, "@name", "missing required value");
    }
    return {
      name,
      flags: children(key, "xlrd2:flag").map((flag) => {
        rejectUnknownChildren(flag, [], part, `${path}/keyFlags/key/flag`);
        rejectUnknownAttributes(flag, ["name", "value"], part, `${path}/keyFlags/key/flag`);
        const flagName = attr(flag, "name");
        const value = attrBool(flag, "value");
        if (flagName === undefined || value === undefined) {
          throw new RichDataParseError(
            part,
            `${path}/keyFlags/key/flag`,
            "flag",
            "invalid name or value",
          );
        }
        return { name: flagName, value };
      }),
    };
  });
  return { keys };
}

function stringifyKeyFlags(flags: RichValueTypeKeyFlagsOptions): string {
  return `<xlrd2:keyFlags>${flags.keys
    .map(
      (key) =>
        `<xlrd2:key name="${escapeXml(key.name)}">${key.flags
          .map(
            (flag) => `<xlrd2:flag name="${escapeXml(flag.name)}" value="${flag.value ? 1 : 0}"/>`,
          )
          .join("")}</xlrd2:key>`,
    )
    .join("")}</xlrd2:keyFlags>`;
}

function rejectUnknownChildren(
  el: Element,
  allowed: readonly string[],
  part: string,
  path: string,
): void {
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    if (!allowed.includes(child.name ?? "")) {
      throw new RichDataParseError(part, path, child.name ?? "unknown", "unexpected child");
    }
  }
}

function rejectUnknownAttributes(
  el: Element,
  allowed: readonly string[],
  part: string,
  path: string,
): void {
  for (const attribute of Object.keys(el.attributes ?? {})) {
    if (attribute === "xmlns" || attribute.startsWith("xmlns:")) continue;
    if (attribute.startsWith("mc:")) continue;
    if (!allowed.includes(attribute)) {
      throw new RichDataParseError(part, path, `@${attribute}`, "unexpected attribute");
    }
  }
}

export const richValueDataDesc: CustomDescriptor<RichValueDataOptions> = {
  kind: "custom",
  stringify(opts) {
    const values = (opts.values ?? [])
      .map(
        (value) =>
          `<xlrd:rv s="${value.structureIndex}">${
            value.fallback
              ? `<xlrd:fb${value.fallback.type !== undefined ? ` t="${escapeXml(value.fallback.type)}"` : ""}>${escapeXml(value.fallback.text ?? "")}</xlrd:fb>`
              : ""
          }${value.values.map((item) => `<xlrd:v>${escapeXml(item)}</xlrd:v>`).join("")}</xlrd:rv>`,
      )
      .join("");
    const count = opts.count !== undefined ? ` count="${opts.count}"` : "";
    return `<xlrd:rvData xmlns:xlrd="${RD_NS}"${count}>${values}${extensionXml(opts.extensions, "xlrd")}</xlrd:rvData>`;
  },
  parse(el) {
    const part = "richData";
    rejectUnknownChildren(el, ["xlrd:rv", "xlrd:extLst"], part, "/rvData");
    rejectUnknownAttributes(el, ["count"], part, "/rvData");
    const count = attrNum(el, "count");
    const values = children(el, "xlrd:rv").map((value) => {
      rejectUnknownChildren(value, ["xlrd:fb", "xlrd:v"], part, "/rvData/rv");
      rejectUnknownAttributes(value, ["s"], part, "/rvData/rv");
      const structureIndex = attrNum(value, "s");
      if (structureIndex === undefined) {
        throw new RichDataParseError(part, "/rvData/rv", "@s", "missing required value");
      }
      const fallbackEl = findChild(value, "xlrd:fb");
      if (fallbackEl) {
        rejectUnknownChildren(fallbackEl, [], part, "/rvData/rv/fb");
        rejectUnknownAttributes(fallbackEl, ["t"], part, "/rvData/rv/fb");
      }
      return {
        structureIndex,
        ...(fallbackEl
          ? {
              fallback: {
                ...(attr(fallbackEl, "t") !== undefined ? { type: attr(fallbackEl, "t") } : {}),
                text: textOf(fallbackEl),
              },
            }
          : {}),
        values: children(value, "xlrd:v").map(textOf),
      };
    });
    const extensions = extensionList(el, part, "/rvData");
    return {
      ...(count !== undefined ? { count } : {}),
      values,
      ...(extensions.length > 0 ? { extensions } : {}),
    };
  },
};

export const richValueStructuresDesc: CustomDescriptor<RichValueStructuresOptions> = {
  kind: "custom",
  stringify(opts) {
    const structures = (opts.structures ?? [])
      .map(
        (structure) =>
          `<xlrd:s t="${escapeXml(structure.type)}">${structure.keys
            .map(
              (key) =>
                `<xlrd:k n="${escapeXml(key.name)}"${key.type !== undefined ? ` t="${key.type}"` : ""}/>`,
            )
            .join("")}</xlrd:s>`,
      )
      .join("");
    const count = opts.count !== undefined ? ` count="${opts.count}"` : "";
    return `<xlrd:rvStructures xmlns:xlrd="${RD_NS}"${count}>${structures}${extensionXml(opts.extensions, "xlrd")}</xlrd:rvStructures>`;
  },
  parse(el) {
    const part = "richData";
    rejectUnknownChildren(el, ["xlrd:s", "xlrd:extLst"], part, "/rvStructures");
    rejectUnknownAttributes(el, ["count"], part, "/rvStructures");
    const count = attrNum(el, "count");
    const structures = children(el, "xlrd:s").map((structure) => {
      rejectUnknownChildren(structure, ["xlrd:k"], part, "/rvStructures/s");
      rejectUnknownAttributes(structure, ["t"], part, "/rvStructures/s");
      const type = attr(structure, "t");
      if (type === undefined) {
        throw new RichDataParseError(part, "/rvStructures/s", "@t", "missing required value");
      }
      return {
        type,
        keys: children(structure, "xlrd:k").map((key) => {
          rejectUnknownChildren(key, [], part, "/rvStructures/s/k");
          rejectUnknownAttributes(key, ["n", "t"], part, "/rvStructures/s/k");
          const name = attr(key, "n");
          if (name === undefined) {
            throw new RichDataParseError(part, "/rvStructures/s/k", "@n", "missing required value");
          }
          const valueType = attr(key, "t");
          return {
            name,
            ...(valueType !== undefined
              ? { type: valueType as RichValueStructureKeyOptions["type"] }
              : {}),
          };
        }),
      };
    });
    const extensions = extensionList(el, part, "/rvStructures");
    return {
      ...(count !== undefined ? { count } : {}),
      structures,
      ...(extensions.length > 0 ? { extensions } : {}),
    };
  },
};

export const richValueTypesInfoDesc: CustomDescriptor<RichValueTypesInfoOptions> = {
  kind: "custom",
  stringify(opts) {
    const global = opts.global
      ? `<xlrd2:global>${opts.global.keyFlags ? stringifyKeyFlags(opts.global.keyFlags) : ""}${extensionXml(opts.global.extensions, "xlrd2")}</xlrd2:global>`
      : "";
    const types = opts.types
      ? `<xlrd2:types>${opts.types.types
          .map(
            (type) =>
              `<xlrd2:type name="${escapeXml(type.name)}">${type.keyFlags ? stringifyKeyFlags(type.keyFlags) : ""}${extensionXml(type.extensions, "xlrd2")}</xlrd2:type>`,
          )
          .join("")}</xlrd2:types>`
      : "";
    return `<xlrd2:rvTypesInfo xmlns:xlrd2="${RD2_NS}">${global}${types}${extensionXml(opts.extensions, "xlrd2")}</xlrd2:rvTypesInfo>`;
  },
  parse(el) {
    const part = "richData";
    rejectUnknownChildren(
      el,
      ["xlrd2:global", "xlrd2:types", "xlrd2:extLst"],
      part,
      "/rvTypesInfo",
    );
    rejectUnknownAttributes(el, [], part, "/rvTypesInfo");
    const globalEl = findChild(el, "xlrd2:global");
    const typesEl = findChild(el, "xlrd2:types");
    if (globalEl) {
      rejectUnknownChildren(
        globalEl,
        ["xlrd2:keyFlags", "xlrd2:extLst"],
        part,
        "/rvTypesInfo/global",
      );
      rejectUnknownAttributes(globalEl, [], part, "/rvTypesInfo/global");
    }
    if (typesEl) {
      rejectUnknownChildren(typesEl, ["xlrd2:type"], part, "/rvTypesInfo/types");
      rejectUnknownAttributes(typesEl, [], part, "/rvTypesInfo/types");
    }
    const globalExtensions = globalEl ? extensionList(globalEl, part, "/rvTypesInfo/global") : [];
    return {
      ...(globalEl
        ? {
            global: {
              ...(findChild(globalEl, "xlrd2:keyFlags")
                ? { keyFlags: parseKeyFlags(globalEl, part, "/rvTypesInfo/global") }
                : {}),
              ...(globalExtensions.length > 0 ? { extensions: globalExtensions } : {}),
            },
          }
        : {}),
      ...(typesEl
        ? {
            types: {
              types: children(typesEl, "xlrd2:type").map((type) => {
                rejectUnknownChildren(
                  type,
                  ["xlrd2:keyFlags", "xlrd2:extLst"],
                  part,
                  "/rvTypesInfo/types/type",
                );
                rejectUnknownAttributes(type, ["name"], part, "/rvTypesInfo/types/type");
                const name = attr(type, "name");
                if (name === undefined) {
                  throw new RichDataParseError(
                    part,
                    "/rvTypesInfo/types/type",
                    "@name",
                    "missing required value",
                  );
                }
                const extensions = extensionList(type, part, "/rvTypesInfo/types/type");
                return {
                  name,
                  ...(findChild(type, "xlrd2:keyFlags")
                    ? { keyFlags: parseKeyFlags(type, part, "/rvTypesInfo/types/type") }
                    : {}),
                  ...(extensions.length > 0 ? { extensions } : {}),
                };
              }),
            },
          }
        : {}),
      ...(extensionList(el, part, "/rvTypesInfo").length > 0
        ? { extensions: extensionList(el, part, "/rvTypesInfo") }
        : {}),
    };
  },
};

export const richValueRelsDesc: CustomDescriptor<RichValueRelsOptions> = {
  kind: "custom",
  stringify(opts) {
    return `<xlrvrel:richValueRels xmlns:xlrvrel="${REL_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${opts.relationships
      .map((relationship) => `<xlrvrel:rel r:id="${escapeXml(relationship.rId)}"/>`)
      .join("")}${extensionXml(opts.extensions, "xlrvrel")}</xlrvrel:richValueRels>`;
  },
  parse(el) {
    const part = "richData";
    rejectUnknownChildren(el, ["xlrvrel:rel", "xlrvrel:extLst"], part, "/richValueRels");
    rejectUnknownAttributes(el, [], part, "/richValueRels");
    return {
      relationships: children(el, "xlrvrel:rel").map((relationship) => {
        rejectUnknownChildren(relationship, [], part, "/richValueRels/rel");
        rejectUnknownAttributes(relationship, ["r:id"], part, "/richValueRels/rel");
        const rId = attr(relationship, "r:id");
        if (rId === undefined) {
          throw new RichDataParseError(
            part,
            "/richValueRels/rel",
            "@r:id",
            "missing required value",
          );
        }
        return { rId };
      }),
      ...(extensionList(el, part, "/richValueRels").length > 0
        ? { extensions: extensionList(el, part, "/richValueRels") }
        : {}),
    };
  },
};
