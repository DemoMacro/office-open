/**
 * ODF codec coverage report backed by an explicit registry.
 *
 * Source text is never scanned. Codec availability is checked through the
 * owning module's public exports; RNG elements are assigned through the
 * registry's schema scopes. Generic OdfXmlNode support is intentionally kept
 * outside canonical and subdocument coverage.
 *
 * Usage:
 *   pnpm tsx scripts/odf-coverage.ts --summary
 *   pnpm tsx scripts/odf-coverage.ts --missing
 *   pnpm tsx scripts/odf-coverage.ts --format odt|ods|odp|chart|database
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { pathToFileURL } from "node:url";
import { fileURLToPath } from "node:url";

import { ODF_CODEC_REGISTRY } from "./lib/odf-codec-registry";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SCHEMA_FILES = [
  "odf-schemas/OpenDocument-v1.3-schema.rng",
  "odf-schemas/OpenDocument-v1.3-manifest-schema.rng",
] as const;

interface RegistryState {
  entry: (typeof ODF_CODEC_REGISTRY)[number];
  packageMissing: boolean;
  moduleMissing: boolean;
  missingExports: string[];
  available: boolean;
  coveredElements?: Set<string>;
}

function parseArguments(arguments_: string[]): { summary: boolean; format?: string } {
  const formatFlag = arguments_.findIndex((argument) => argument === "--format");
  const flagFormat = formatFlag >= 0 ? arguments_[formatFlag + 1] : undefined;
  const positional = arguments_.find((argument, index) => {
    if (argument === "--format") return false;
    if (index === formatFlag + 1) return false;
    return !argument.startsWith("--");
  });
  const format = flagFormat ?? positional;
  const validFormats = new Set(
    ODF_CODEC_REGISTRY.filter(
      (entry) => entry.classification === "canonical" || entry.classification === "subdocument",
    ).map((entry) => entry.id),
  );
  if (format !== undefined && !validFormats.has(format)) {
    console.error(`unknown format: ${format}`);
    console.error(`valid formats: ${[...validFormats].join(", ")}`);
    process.exit(2);
  }
  return { summary: arguments_.includes("--summary"), format };
}

function schemaElements(): Map<string, Set<string>> {
  const elements = new Map<string, Set<string>>();
  for (const schemaFile of SCHEMA_FILES) {
    const filePath = path.join(ROOT, schemaFile);
    if (!fs.existsSync(filePath)) continue;
    const schema = fs.readFileSync(filePath, "utf8");
    for (const match of schema.matchAll(/<rng:element name="([^"]+)"\s*(?:\/>|>)/g)) {
      const element = match[1]!;
      const prefix = element.split(":")[0]!;
      const scoped = elements.get(prefix) ?? new Set<string>();
      scoped.add(element);
      elements.set(prefix, scoped);
    }
  }
  return elements;
}

async function registryState(): Promise<RegistryState[]> {
  return Promise.all(
    ODF_CODEC_REGISTRY.map(async (entry) => {
      const packageDirectory = path.join(ROOT, "packages", entry.owner.package);
      const modulePath = path.join(packageDirectory, entry.owner.module);
      const packageMissing = !fs.existsSync(path.join(packageDirectory, "package.json"));
      const moduleMissing = !fs.existsSync(modulePath);
      const missingExports: string[] = [];

      if (!packageMissing && !moduleMissing && entry.owner.export !== null) {
        try {
          const module = await import(pathToFileURL(modulePath).href);
          const exports = [entry.owner.export, ...(entry.roundTrip ?? [])];
          for (const name of new Set(exports)) {
            if (!(name in module)) missingExports.push(name);
          }
        } catch (error) {
          console.warn(`coverage warning: ${entry.id} module could not load`);
          console.warn(String(error instanceof Error ? error.message : error).split("\n")[0]);
          missingExports.push(entry.owner.export);
        }
      }

      return {
        entry,
        packageMissing,
        moduleMissing,
        missingExports,
        available: !packageMissing && !moduleMissing && missingExports.length === 0,
      };
    }),
  );
}

function registryCoverage(states: readonly RegistryState[], classification: string) {
  const selected = states.filter((state) => state.entry.classification === classification);
  return {
    selected,
    total: selected.length,
    available: selected.filter((state) => state.available).length,
  };
}

function prefixOwner(states: readonly RegistryState[]) {
  const priority = { canonical: 0, subdocument: 1, "generic-only": 2, unsupported: 3 } as const;
  const owners = new Map<string, RegistryState>();
  for (const state of [...states].sort(
    (left, right) =>
      priority[left.entry.classification] - priority[right.entry.classification] ||
      left.entry.id.localeCompare(right.entry.id),
  )) {
    for (const prefix of state.entry.schemaPrefixes) {
      if (!owners.has(prefix)) owners.set(prefix, state);
    }
  }
  return owners;
}

function printMissing(
  states: readonly RegistryState[],
  elements: ReadonlyMap<string, Set<string>>,
) {
  const actionable = states.filter(
    (state) => !state.available && state.entry.classification !== "unsupported",
  );
  if (actionable.length === 0) return;

  console.log("\nMissing codecs:");
  for (const state of actionable) {
    const reasons: string[] = [];
    if (state.packageMissing) reasons.push("package not integrated");
    if (state.moduleMissing) reasons.push(`module ${state.entry.owner.module} missing`);
    if (state.missingExports.length > 0)
      reasons.push(`missing exports ${state.missingExports.join(", ")}`);
    console.log(`  - ${state.entry.id}: ${reasons.join("; ")}`);
  }

  const owners = prefixOwner(states);
  const missingByPrefix = new Map<string, string[]>();
  for (const [prefix, scopedElements] of elements) {
    const owner = owners.get(prefix);
    if (!owner?.available || owner.entry.classification === "unsupported") continue;
    const missing = [...scopedElements].filter((element) => !owner.coveredElements?.has(element));
    if (missing.length > 0) missingByPrefix.set(prefix, missing.sort());
  }

  if (missingByPrefix.size > 0) {
    console.log("\nMissing schema elements:");
    for (const [prefix, missing] of missingByPrefix) {
      console.log(`  ${prefix}: ${missing.length}`);
      for (const element of missing) console.log(`    - ${element}`);
    }
  }
}

function percent(numerator: number, denominator: number): string {
  return denominator === 0 ? "n/a" : `${((numerator / denominator) * 100).toFixed(1)}%`;
}

async function main() {
  const { summary, format } = parseArguments(process.argv.slice(2));
  const elements = schemaElements();
  const allStates = await registryState();
  const states = format
    ? allStates.filter((state) => state.entry.id === format)
    : allStates.filter((state) => state.entry.id !== "generic-node");
  const genericState = allStates.find((state) => state.entry.id === "generic-node");

  if (elements.size === 0) {
    console.warn("ODF coverage: Relax NG schemas unavailable");
    process.exitCode = 1;
  }

  for (const state of states) {
    const scoped = new Set(
      state.entry.schemaPrefixes.flatMap((prefix) => [...(elements.get(prefix) ?? [])]),
    );
    state.coveredElements = scoped;
  }

  const canonical = registryCoverage(states, "canonical");
  const subdocument = registryCoverage(states, "subdocument");
  const semantic = states.filter((state) => state.entry.roundTrip && state.available).length;
  const semanticTotal = states.filter((state) => state.entry.roundTrip).length;
  const schemaElementsTotal = [...elements.values()].reduce(
    (total, scoped) => total + scoped.size,
    0,
  );
  const schemaCovered = states
    .filter((state) => state.available && state.entry.classification !== "unsupported")
    .flatMap((state) => state.entry.schemaPrefixes)
    .flatMap((prefix) => [...(elements.get(prefix) ?? [])]);
  const schemaCoveredCount = new Set(schemaCovered).size;

  console.log("======================================================================");
  console.log("ODF Codec Coverage");
  console.log("======================================================================");
  console.log(
    `Schema elements: ${schemaCoveredCount}/${schemaElementsTotal} (${percent(schemaCoveredCount, schemaElementsTotal)})`,
  );
  console.log(
    `Canonical codecs: ${canonical.available}/${canonical.total} (${percent(canonical.available, canonical.total)})`,
  );
  console.log(
    `Subdocuments: ${subdocument.available}/${subdocument.total} (${percent(subdocument.available, subdocument.total)})`,
  );
  console.log(
    `Semantic round-trip: ${semantic}/${semanticTotal} (${percent(semantic, semanticTotal)})`,
  );
  console.log(
    `Generic OdfXmlNode: ${genericState ? (genericState.available ? "available" : "missing") : "not registered"} (non-canonical)`,
  );

  if (!summary) {
    console.log("\nRegistry:");
    for (const state of states) {
      const stateLabel =
        state.entry.classification === "unsupported"
          ? "unsupported"
          : state.available
            ? "available"
            : state.packageMissing
              ? "missing package"
              : state.moduleMissing
                ? "missing module"
                : "missing exports";
      console.log(
        `  ${state.entry.id.padEnd(10)} ${stateLabel.padEnd(17)} ${state.entry.owner.package}/${state.entry.owner.module}`,
      );
    }
    printMissing(states, elements);
  }

  const required = [...canonical.selected, ...subdocument.selected];
  if (required.some((state) => !state.available)) process.exitCode = 1;
}

await main();
