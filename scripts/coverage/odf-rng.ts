import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { parse, type Element } from "@office-open/xml";

export interface SchemaElementDescriptor {
  name: string;
  attributes: readonly string[];
}

export function rngElementDescriptors(
  schemaFiles: readonly string[],
  prefixes: readonly string[],
): readonly SchemaElementDescriptor[] {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
  const capabilities = new Map<string, Set<string>>();
  for (const schemaFile of schemaFiles) {
    const filePath = path.join(root, schemaFile);
    if (!fs.existsSync(filePath)) continue;
    const grammar = parse(fs.readFileSync(filePath, "utf8"), {
      ignoreDeclaration: true,
    }).elements?.find((element): element is Element => element.type === "element");
    const defines = new Map<string, Element>();
    for (const child of grammar?.elements ?? []) {
      if (child.type === "element" && child.name === "rng:define")
        defines.set(String(child.attributes?.name ?? ""), child);
    }
    for (const elementNode of elementNodes(grammar!)) {
      const name = String(elementNode.attributes?.name ?? "");
      if (!name || !prefixes.includes(name.split(":")[0]!)) continue;
      const attributes = capabilities.get(name) ?? new Set<string>();
      collect(elementNode, defines, attributes);
      capabilities.set(name, attributes);
    }
  }
  return [...capabilities].map(([name, attributes]) => ({
    name,
    attributes: [...attributes].sort(),
  }));
}

function elementNodes(node: Element): Element[] {
  return (node.elements ?? []).flatMap((child): Element[] => {
    if (child.type !== "element") return [];
    return child.name === "rng:element" ? [child, ...elementNodes(child)] : elementNodes(child);
  });
}

function collect(node: Element, defines: ReadonlyMap<string, Element>, result: Set<string>): void {
  for (const child of node.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === "rng:attribute" && typeof child.attributes?.name === "string")
      result.add(child.attributes.name);
    else if (child.name === "rng:ref" && typeof child.attributes?.name === "string") {
      const target = defines.get(child.attributes.name);
      if (target) collect(target, defines, result);
    } else if (child.name !== "rng:element") collect(child, defines, result);
  }
}
