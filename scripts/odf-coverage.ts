import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { parse, type Element } from "@office-open/xml";

import { ODF_CODEC_REGISTRY } from "./lib/odf-codec-registry";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SCHEMA_FILES = [
  "odf-schemas/OpenDocument-v1.3-schema.rng",
  "odf-schemas/OpenDocument-v1.3-manifest-schema.rng",
] as const;

interface SchemaCapabilities {
  elements: ReadonlyMap<string, ReadonlySet<string>>;
}

interface RegistryState {
  entry: (typeof ODF_CODEC_REGISTRY)[number];
  available: boolean;
  missingExports: string[];
  capabilities: ReadonlyMap<string, ReadonlySet<string>>;
}

function schemaCapabilities(): SchemaCapabilities {
  const capabilities = new Map<string, Set<string>>();
  for (const schemaFile of SCHEMA_FILES) {
    const filePath = path.join(ROOT, schemaFile);
    if (!fs.existsSync(filePath)) continue;
    const grammar = parse(fs.readFileSync(filePath, "utf8"), {
      ignoreDeclaration: true,
    }).elements?.find((element): element is Element => element.type === "element");
    const defines = new Map<string, Element>();
    for (const child of grammar?.elements ?? []) {
      if (child.type === "element" && child.name === "rng:define") {
        defines.set(String(child.attributes?.name ?? ""), child);
      }
    }
    for (const definition of defines.values()) {
      for (const elementNode of descendants(definition, "rng:element")) {
        const name = String(elementNode.attributes?.name ?? "");
        if (!name || name.includes(":any")) continue;
        const attributes = capabilities.get(name) ?? new Set<string>();
        collectAttributes(elementNode, defines, attributes);
        capabilities.set(name, attributes);
      }
    }
  }
  return { elements: capabilities };
}

function descendants(node: Element, name: string): Element[] {
  return (node.elements ?? []).flatMap((child): Element[] => {
    if (child.type !== "element") return [];
    return child.name === name ? [child, ...descendants(child, name)] : descendants(child, name);
  });
}

function collectAttributes(
  node: Element,
  defines: ReadonlyMap<string, Element>,
  result: Set<string>,
): void {
  for (const child of node.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === "rng:attribute") {
      const name = child.attributes?.name;
      if (typeof name === "string") result.add(name);
    } else if (child.name === "rng:ref" && child.attributes?.name) {
      const target = defines.get(String(child.attributes.name));
      if (target) collectAttributes(target, defines, result);
    } else if (child.name !== "rng:element") {
      collectAttributes(child, defines, result);
    }
  }
}

function descriptorCapabilities(
  entry: (typeof ODF_CODEC_REGISTRY)[number],
): ReadonlyMap<string, ReadonlySet<string>> {
  const result = new Map<string, Set<string>>();
  for (const descriptor of entry.schemaElements ?? []) {
    result.set(descriptor.name, new Set(descriptor.attributes));
  }
  return result;
}

async function registryStates(): Promise<RegistryState[]> {
  return Promise.all(
    ODF_CODEC_REGISTRY.map(async (entry) => {
      const modulePath = path.join(ROOT, "packages", entry.owner.package, entry.owner.module);
      const missingExports: string[] = [];
      try {
        const module = await import(pathToFileURL(modulePath).href);
        for (const name of new Set(
          [entry.owner.export, ...(entry.roundTrip ?? [])].filter((name): name is string =>
            Boolean(name),
          ),
        )) {
          if (!(name in module)) missingExports.push(name);
        }
      } catch (error) {
        console.warn(
          `coverage warning: ${entry.id}: ${error instanceof Error ? error.message.split("\n")[0] : "module load failed"}`,
        );
        missingExports.push(entry.owner.export ?? "<module>");
      }
      return {
        entry,
        available: fs.existsSync(modulePath) && missingExports.length === 0,
        missingExports,
        capabilities: descriptorCapabilities(entry),
      };
    }),
  );
}

function axis(numerator: number, denominator: number): string {
  return denominator === 0 ? "n/a" : `${((numerator / denominator) * 100).toFixed(1)}%`;
}

function capabilityDeltas(
  state: RegistryState,
  schema: SchemaCapabilities,
): {
  missingElements: string[];
  missingAttributes: Array<[string, string]>;
  extraElements: string[];
} {
  const missingElements = [...schema.elements.keys()].filter(
    (name) => !state.capabilities.has(name),
  );
  const missingAttributes: Array<[string, string]> = [];
  for (const [name, required] of schema.elements) {
    const actual = state.capabilities.get(name);
    if (!actual) continue;
    missingAttributes.push(
      ...[...required]
        .filter((attribute) => !actual.has(attribute))
        .map((attribute) => [name, attribute] as [string, string]),
    );
  }
  const extraElements = [...state.capabilities.keys()].filter((name) => !schema.elements.has(name));
  return { missingElements, missingAttributes, extraElements };
}

async function main(): Promise<void> {
  const summary = process.argv.includes("--summary");
  const schema = schemaCapabilities();
  const states = await registryStates();
  const ownedElements = states.flatMap((state) => [...state.capabilities.keys()]);
  const schemaTotal = schema.elements.size;
  const schemaCovered = new Set(ownedElements.filter((name) => schema.elements.has(name))).size;
  const canonical = states.filter((state) => state.entry.classification === "canonical");
  const subdocuments = states.filter((state) => state.entry.classification === "subdocument");
  const semantic = states.filter((state) => state.entry.roundTrip);
  const schemaCoverage =
    schemaCovered === schemaTotal &&
    states.every((state) => capabilityDeltas(state, schema).missingAttributes.length === 0);
  const canonicalCoverage = canonical.every((state) => state.available) && canonical.length > 0;
  const subdocCoverage = subdocuments.every((state) => state.available) && subdocuments.length > 0;
  const semanticRoundTrip = semantic.every((state) => state.available) && semantic.length > 0;

  console.log("======================================================================");
  console.log("ODF Codec Coverage");
  console.log("======================================================================");
  console.log(
    `schemaCoverage: ${axis(schemaCovered, schemaTotal)} (${schemaCovered}/${schemaTotal} elements; attributes ${schemaCoverage ? "complete" : "incomplete"})`,
  );
  console.log(
    `canonicalCoverage: ${axis(canonical.filter((state) => state.available).length, canonical.length)}`,
  );
  console.log(
    `subdocCoverage: ${axis(subdocuments.filter((state) => state.available).length, subdocuments.length)}`,
  );
  console.log(
    `semanticRoundTrip: ${axis(semantic.filter((state) => state.available).length, semantic.length)}`,
  );
  console.log(
    `unsupported/generic-only entries: ${states.filter((state) => ["unsupported", "generic-only"].includes(state.entry.classification)).length}`,
  );

  if (!summary) {
    for (const state of states) {
      const delta = capabilityDeltas(state, schema);
      console.log(
        `\n${state.entry.id}: ${state.available ? "available" : `missing ${state.missingExports.join(", ")}`}`,
      );
      if (delta.missingElements.length)
        console.log(
          `  missing elements: ${delta.missingElements.length}\n    ${delta.missingElements.slice(0, 20).join("\n    ")}`,
        );
      if (delta.missingAttributes.length)
        console.log(
          `  missing attributes: ${delta.missingAttributes.length}\n    ${delta.missingAttributes
            .slice(0, 20)
            .map(([element, attribute]) => `${element} ${attribute}`)
            .join("\n    ")}`,
        );
      if (delta.extraElements.length)
        console.log(`  unknown descriptor elements: ${delta.extraElements.join(", ")}`);
    }
  }

  const gate =
    schemaCoverage &&
    canonicalCoverage &&
    subdocCoverage &&
    semanticRoundTrip &&
    states.every((state) => state.available) &&
    states.every((state) => !["unsupported", "generic-only"].includes(state.entry.classification));
  if (!gate) process.exitCode = 1;
}

await main();
