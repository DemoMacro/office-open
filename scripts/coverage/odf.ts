import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { parse, type Element } from "@office-open/xml";

import { runOdfCoverageFixture, runOdfNegativeFixture } from "./odf-fixtures";
import { ODF_CODEC_REGISTRY } from "./odf-registry";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
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
  let loaded = 0;
  for (const schemaFile of SCHEMA_FILES) {
    const filePath = path.join(ROOT, schemaFile);
    if (!fs.existsSync(filePath)) continue;
    loaded += 1;
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
  if (loaded !== SCHEMA_FILES.length || capabilities.size === 0) {
    throw new Error(`ODF RNG schemas are incomplete (${loaded}/${SCHEMA_FILES.length} files)`);
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
        missingExports.push(
          error instanceof Error
            ? error.message.split("\n")[0]
            : (entry.owner.export ?? "<module>"),
        );
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

function ownershipFailures(schema: SchemaCapabilities, states: readonly RegistryState[]): string[] {
  const owned = new Set<string>();
  const failures: string[] = [];
  for (const state of states) {
    for (const [name, attributes] of state.capabilities) {
      const required = schema.elements.get(name);
      if (!required) continue;
      owned.add(name);
      const missing = [...required].filter((attribute) => !attributes.has(attribute));
      if (missing.length > 0) {
        failures.push(`${name}: missing attributes ${missing.join(", ")}`);
      }
    }
  }
  const missing = [...schema.elements.keys()].filter((name) => !owned.has(name));
  if (missing.length > 0) {
    failures.push(`schema elements without a named mapper: ${missing.length}`);
  }
  const extra = states.flatMap((state) =>
    [...state.capabilities.keys()].filter((name) => !schema.elements.has(name)),
  );
  if (extra.length > 0) {
    failures.push(`registry elements outside the RNG universe: ${extra.join(", ")}`);
  }
  for (const entry of states) {
    for (const capability of entry.entry.negativeCapabilities ?? []) {
      if (owned.has(capability.element)) {
        failures.push(
          `${entry.entry.id}: strict-throw element is marked covered: ${capability.element}`,
        );
      }
    }
  }
  return failures;
}

async function main(): Promise<void> {
  const summary = process.argv.includes("--summary");
  const schema = schemaCapabilities();
  const states = await registryStates();
  const failures: string[] = [];

  for (const state of states) {
    const entry = state.entry;
    if (entry.classification === "unsupported" || entry.classification === "generic-only") {
      failures.push(`${entry.id}: ${entry.classification} ownership cannot pass`);
    }
    if (!entry.schemaElements?.length) failures.push(`${entry.id}: no schema elements`);
    if (!entry.testId || !entry.fixtureKey) {
      failures.push(`${entry.id}: missing executable test id or fixture key`);
    }
    if (!state.available) {
      failures.push(`${entry.id}: unavailable (${state.missingExports.join("; ")})`);
    }
  }

  for (const entry of ODF_CODEC_REGISTRY) {
    try {
      runOdfCoverageFixture(entry.fixtureKey as Parameters<typeof runOdfCoverageFixture>[0]);
    } catch (error) {
      failures.push(
        `${entry.id}/${entry.testId}: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
    for (const negative of entry.negativeCapabilities ?? []) {
      try {
        runOdfNegativeFixture(negative.testId as Parameters<typeof runOdfNegativeFixture>[0]);
      } catch (error) {
        failures.push(
          `${entry.id}/${negative.testId}: ${error instanceof Error ? error.message : String(error)}`,
        );
      }
    }
  }

  failures.push(...ownershipFailures(schema, states));

  const ownedElements = new Set(states.flatMap((state) => [...state.capabilities.keys()]));
  const schemaCovered = [...schema.elements.keys()].filter((name) =>
    ownedElements.has(name),
  ).length;
  const canonical = states.filter((state) => state.entry.classification === "canonical");
  const subdocuments = states.filter((state) => state.entry.classification === "subdocument");
  const negativeCount = states.reduce(
    (total, state) => total + (state.entry.negativeCapabilities?.length ?? 0),
    0,
  );

  console.log("======================================================================");
  console.log("ODF Codec Coverage");
  console.log("======================================================================");
  console.log(
    `schemaCoverage: ${axis(schemaCovered, schema.elements.size)} (${schemaCovered}/${schema.elements.size} elements; owned by named mappers)`,
  );
  console.log(
    `canonicalCoverage: ${axis(canonical.filter((state) => state.available).length, canonical.length)}`,
  );
  console.log(
    `subdocCoverage: ${axis(subdocuments.filter((state) => state.available).length, subdocuments.length)}`,
  );
  console.log(`executable fixture tests: ${states.length}`);
  console.log(`strict-throw negative tests: ${negativeCount}`);
  console.log(
    `unsupported/generic-only entries: ${states.filter((state) => ["unsupported", "generic-only"].includes(state.entry.classification)).length}`,
  );

  if (!summary) {
    for (const failure of failures) console.log(`  FAIL ${failure}`);
  }
  if (failures.length > 0) {
    console.error(`ODF coverage gate: FAILED (${failures.length} failures)`);
    process.exitCode = 1;
    return;
  }
  console.log("ODF coverage gate: OK");
}

await main();
