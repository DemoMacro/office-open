import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrBool, children, escapeXml } from "@office-open/xml";

import type { ActionDefinitionOptions, KeyMapOptions, WordCustomizationsOptions } from "./types";

const WNE_NAMESPACE = "http://schemas.microsoft.com/office/word/2006/wordml";

function stringifyKeyMap(map: KeyMapOptions): string {
  const attrs: string[] = [];
  if (map.mask !== undefined) attrs.push(` wne:mask="${map.mask ? 1 : 0}"`);
  if (map.primaryKeyCode !== undefined)
    attrs.push(` wne:kcmPrimary="${escapeXml(map.primaryKeyCode)}"`);
  if (map.secondaryKeyCode !== undefined)
    attrs.push(` wne:kcmSecondary="${escapeXml(map.secondaryKeyCode)}"`);
  if (map.actionName !== undefined) {
    return `<wne:keymap${attrs.join("")}><wne:acd wne:acdName="${escapeXml(map.actionName)}"/></wne:keymap>`;
  }
  if (map.characterCode !== undefined) {
    return `<wne:keymap${attrs.join("")}><wne:wch wne:val="${escapeXml(map.characterCode)}"/></wne:keymap>`;
  }
  if (map.macroName !== undefined) {
    return `<wne:keymap${attrs.join("")}><wne:macro wne:macroName="${escapeXml(map.macroName)}"/></wne:keymap>`;
  }
  return `<wne:keymap${attrs.join("")}/>`;
}

function stringifyActionDefinition(definition: ActionDefinitionOptions): string {
  const attributes = [
    definition.basedOn !== undefined && ` wne:fciBasedOn="${escapeXml(definition.basedOn)}"`,
    definition.argumentValue !== undefined &&
      ` wne:argValue="${escapeXml(definition.argumentValue)}"`,
    ` wne:acdName="${escapeXml(definition.name)}"`,
    definition.indexBasedOn !== undefined &&
      ` wne:fciIndexBasedOn="${escapeXml(definition.indexBasedOn)}"`,
  ].filter((value): value is string => Boolean(value));
  return `<wne:acd${attributes.join("")}/>`;
}

export const customizationsDesc: CustomDescriptor<WordCustomizationsOptions> = {
  kind: "custom",
  stringify(options) {
    const blocks: string[] = [];
    const toolbarChildren: string[] = [];
    if (options.toolbarDataRelationshipId !== undefined)
      toolbarChildren.push(
        `<wne:toolbarData r:id="${escapeXml(options.toolbarDataRelationshipId)}"/>`,
      );
    if (options.toolbarActionNames?.length) {
      toolbarChildren.push(
        `<wne:acdManifest>${options.toolbarActionNames
          .map((name) => `<wne:acdEntry wne:acdName="${escapeXml(name)}"/>`)
          .join("")}</wne:acdManifest>`,
      );
    }
    if (toolbarChildren.length > 0)
      blocks.push(`<wne:toolbars>${toolbarChildren.join("")}</wne:toolbars>`);
    if (options.keyMaps?.length) {
      blocks.push(`<wne:keymaps>${options.keyMaps.map(stringifyKeyMap).join("")}</wne:keymaps>`);
    }
    if (options.actionDefinitions?.length) {
      const definitions = options.actionDefinitions.map(stringifyActionDefinition).join("");
      blocks.push(`<wne:acds>${definitions}</wne:acds>`);
    }
    if (blocks.length === 0) return undefined;
    const namespaces = `xmlns:wne="${WNE_NAMESPACE}"`;
    return options.toolbarDataRelationshipId === undefined
      ? `<wne:tcg ${namespaces}>${blocks.join("")}</wne:tcg>`
      : `<wne:tcg ${namespaces} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${blocks.join("")}</wne:tcg>`;
  },
  parse(el) {
    const options: WordCustomizationsOptions = {};
    const keyMaps = children(el, "wne:keymaps")
      .flatMap((parent) => children(parent, "wne:keymap"))
      .map((map): KeyMapOptions => ({
        mask: attrBool(map, "wne:mask"),
        primaryKeyCode: attr(map, "wne:kcmPrimary"),
        secondaryKeyCode: attr(map, "wne:kcmSecondary"),
        actionName: attr(children(map, "wne:acd")[0], "wne:acdName"),
        characterCode: attr(children(map, "wne:wch")[0], "wne:val"),
        macroName: attr(children(map, "wne:macro")[0], "wne:macroName"),
      }));
    if (keyMaps.length > 0) options.keyMaps = keyMaps;
    const toolbarActionNames = children(el, "wne:toolbars")
      .flatMap((parent) => children(parent, "wne:acdManifest"))
      .flatMap((parent) => children(parent, "wne:acdEntry"))
      .map((entry) => attr(entry, "wne:acdName"))
      .filter((name): name is string => name !== undefined);
    if (toolbarActionNames.length > 0) options.toolbarActionNames = toolbarActionNames;
    const toolbarData = children(el, "wne:toolbars")
      .flatMap((parent) => children(parent, "wne:toolbarData"))
      .map((data) => attr(data, "r:id"))
      .find((relationshipId): relationshipId is string => relationshipId !== undefined);
    if (toolbarData !== undefined) options.toolbarDataRelationshipId = toolbarData;
    const actionDefinitions = children(el, "wne:acds")
      .flatMap((parent) => children(parent, "wne:acd"))
      .map((definition): ActionDefinitionOptions => ({
        name: attr(definition, "wne:acdName") ?? "",
        basedOn: attr(definition, "wne:fciBasedOn"),
        argumentValue: attr(definition, "wne:argValue"),
        indexBasedOn: attr(definition, "wne:fciIndexBasedOn"),
      }))
      .filter((definition) => definition.name !== "");
    if (actionDefinitions.length > 0) options.actionDefinitions = actionDefinitions;
    return options;
  },
};
