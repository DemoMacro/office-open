import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import {
  runOdfCoverageFixture,
  runOdfNegativeFixture,
  type OdfCoverageFixtureResult,
} from "./odf-fixtures";
import { ODF_CODEC_REGISTRY } from "./odf-registry";
import { rngElementDescriptors } from "./odf-rng";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const SCHEMA_FILES = [
  "odf-schemas/OpenDocument-v1.3-schema.rng",
  "odf-schemas/OpenDocument-v1.3-manifest-schema.rng",
] as const;

interface RegistryState {
  entry: (typeof ODF_CODEC_REGISTRY)[number];
  available: boolean;
  missingExports: string[];
  capabilities: ReadonlyMap<string, ReadonlySet<string>>;
  verified: boolean;
}

function rngPrefixes(states: readonly RegistryState[]): string[] {
  return [
    ...new Set(
      states.flatMap((state) =>
        [...state.capabilities.keys()].map((name) => name.split(":")[0] ?? ""),
      ),
    ),
  ].filter(Boolean);
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
        verified: false,
      };
    }),
  );
}

function axis(numerator: number, denominator: number): string {
  return denominator === 0 ? "n/a" : `${((numerator / denominator) * 100).toFixed(1)}%`;
}

function rngAuditFailures(
  states: readonly RegistryState[],
  rng: ReadonlyMap<string, ReadonlySet<string>>,
): string[] {
  const failures: string[] = [];
  for (const state of states) {
    for (const [name, attributes] of state.capabilities) {
      const schemaAttributes = rng.get(name);
      if (!schemaAttributes) {
        failures.push(`${state.entry.id}: explicit element is absent from RNG: ${name}`);
        continue;
      }
      const missing = [...attributes].filter((attribute) => !schemaAttributes.has(attribute));
      if (missing.length > 0) {
        failures.push(
          `${state.entry.id}/${name}: explicit attributes outside RNG: ${missing.join(", ")}`,
        );
      }
    }
  }
  return failures;
}

async function main(): Promise<void> {
  const summary = process.argv.includes("--summary");
  const states = await registryStates();
  const failures: string[] = [];
  const fixtureResults = new Map<string, OdfCoverageFixtureResult>();

  for (const state of states) {
    const entry = state.entry;
    if (entry.classification === "unsupported" || entry.classification === "generic-only") {
      failures.push(`${entry.id}: ${entry.classification} ownership cannot pass`);
    }
    if (
      entry.classification === "canonical" &&
      entry.parts.join(",") !== "content.xml,styles.xml,meta.xml"
    ) {
      failures.push(`${entry.id}: canonical ownership must cover content, styles, and meta`);
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
      const result = runOdfCoverageFixture(
        entry.fixtureKey as Parameters<typeof runOdfCoverageFixture>[0],
      );
      if (!result.semanticRoundTrip) {
        failures.push(`${entry.id}/${entry.testId}: semantic round-trip was not verified`);
      } else {
        fixtureResults.set(entry.fixtureKey, result);
      }
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

  for (const state of states) state.verified = fixtureResults.has(state.entry.fixtureKey);

  const rngDescriptors = rngElementDescriptors(SCHEMA_FILES, rngPrefixes(states));
  const rng = new Map<string, Set<string>>();
  for (const descriptor of rngDescriptors) {
    const attributes = rng.get(descriptor.name) ?? new Set<string>();
    for (const attribute of descriptor.attributes) attributes.add(attribute);
    rng.set(descriptor.name, attributes);
  }
  if (rngDescriptors.length === 0)
    failures.push("RNG ownership audit: no schema descriptors loaded");
  failures.push(...rngAuditFailures(states, rng));

  const ownedCount = new Set(
    states
      .filter((state) => state.verified)
      .flatMap((state) =>
        [...state.capabilities.keys()].map((element) => `${state.entry.id}:${element}`),
      ),
  ).size;
  const totalCount = states.reduce((total, state) => total + state.capabilities.size, 0);
  const canonical = states.filter((state) => state.entry.classification === "canonical");
  const subdocuments = states.filter((state) => state.entry.classification === "subdocument");
  const negativeCount = states.reduce(
    (total, state) => total + (state.entry.negativeCapabilities?.length ?? 0),
    0,
  );
  const requiredRawFallbacks = [
    ...new Set(states.flatMap((state) => state.entry.rawFallbacks ?? [])),
  ];
  const verifiedRawFallbacks = new Set(
    [...fixtureResults.values()].flatMap((result) => result.rawFallbackElements ?? []),
  );
  for (const fallback of requiredRawFallbacks) {
    if (!verifiedRawFallbacks.has(fallback)) {
      failures.push(`raw fallback is unverified: ${fallback}`);
    }
  }

  console.log("======================================================================");
  console.log("ODF Codec Coverage");
  console.log("======================================================================");
  console.log(
    `ownershipCoverage: ${axis(ownedCount, totalCount)} (${ownedCount}/${totalCount} explicit element owners verified by execution)`,
  );
  console.log(
    `canonicalCoverage: ${axis(canonical.filter((state) => state.available && state.verified).length, canonical.length)}`,
  );
  console.log(
    `subdocCoverage: ${axis(subdocuments.filter((state) => state.available && state.verified).length, subdocuments.length)}`,
  );
  console.log(
    `semanticRoundTrip: ${fixtureResults.size}/${states.length} positive ownership entries`,
  );
  console.log(
    `executable fixture tests: ${new Set(ODF_CODEC_REGISTRY.map((entry) => entry.fixtureKey)).size}`,
  );
  console.log(`strict-throw negative tests: ${negativeCount}`);
  console.log(
    `rawFallback: ${verifiedRawFallbacks.size}/${requiredRawFallbacks.length} (${[...verifiedRawFallbacks].sort().join(", ") || "none"})`,
  );
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
