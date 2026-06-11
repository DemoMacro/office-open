/**
 * Relationships descriptor — produces .rels XML files.
 *
 * Reference: OPC, Relationships.xsd
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { escapeXml } from "@office-open/xml";

export interface RelationshipEntry {
  id: number | string;
  type: string;
  target: string;
  targetMode?: string;
}

export interface RelationshipsInput {
  relationships: RelationshipEntry[];
}

export const relationshipsDesc: CustomDescriptor<RelationshipsInput> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ];
    for (const rel of opts.relationships) {
      const targetModeAttr = rel.targetMode ? ` TargetMode="${escapeXml(rel.targetMode)}"` : "";
      p.push(
        `<Relationship Id="rId${rel.id}" Type="${escapeXml(rel.type)}" Target="${escapeXml(rel.target)}"${targetModeAttr}/>`,
      );
    }
    p.push("</Relationships>");
    return p.join("");
  },

  parse(el, _ctx) {
    const relationships: RelationshipEntry[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "Relationship") continue;
      const id = child.attributes?.["Id"];
      const type = child.attributes?.["Type"];
      const target = child.attributes?.["Target"];
      const targetMode = child.attributes?.["TargetMode"];
      if (id && type && target) {
        const numId = String(id).replace("rId", "");
        relationships.push({
          id: Number(numId) || id,
          type: String(type),
          target: String(target),
          targetMode: targetMode ? String(targetMode) : undefined,
        });
      }
    }
    return { relationships } as Record<string, unknown>;
  },
};
