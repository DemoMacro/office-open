/**
 * Legacy and RTF capability report backed by an explicit registry.
 *
 * Capability names and owners are declared data; source strings are never
 * pattern-matched. The public export named by each capability owner must exist
 * at runtime, and an explicitly unsupported capability remains missing.
 *
 * Usage:
 *   pnpm tsx scripts/legacy-coverage.ts              # full report
 *   pnpm tsx scripts/legacy-coverage.ts doc|xls|ppt|rtf
 *   pnpm tsx scripts/legacy-coverage.ts --missing    # default
 *   pnpm tsx scripts/legacy-coverage.ts --summary    # summary table only
 */
import { existsSync } from "node:fs";
import * as path from "node:path";
import { pathToFileURL } from "node:url";
import { fileURLToPath } from "node:url";

import { LEGACY_CAPABILITY_REGISTRY } from "./lib/legacy-capability-registry";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

interface CapabilityResult {
  format: string;
  mode: string;
  capability: (typeof LEGACY_CAPABILITY_REGISTRY)[number]["capabilities"][number];
  packageMissing: boolean;
  moduleMissing: boolean;
  exportMissing: boolean;
  covered: boolean;
}

function filter(): string | undefined {
  const format = process.argv.slice(2).find((argument) => !argument.startsWith("--"));
  if (format === undefined) return undefined;
  if (!LEGACY_CAPABILITY_REGISTRY.some((entry) => entry.id === format)) {
    console.error(`unknown format: ${format}`);
    console.error(
      `valid formats: ${LEGACY_CAPABILITY_REGISTRY.map((entry) => entry.id).join(", ")}`,
    );
    process.exit(2);
  }
  return format;
}

async function results(): Promise<CapabilityResult[]> {
  const output: CapabilityResult[] = [];
  for (const format of LEGACY_CAPABILITY_REGISTRY) {
    const packageDirectory = path.join(ROOT, "packages", format.id);
    const modulePath = path.join(packageDirectory, format.capabilities[0]!.owner.module);
    const packageMissing = !existsSync(path.join(packageDirectory, "package.json"));
    const moduleMissing = !existsSync(modulePath);
    let module: Record<string, unknown> | undefined;
    if (!packageMissing && !moduleMissing) {
      try {
        module = await import(pathToFileURL(modulePath).href);
      } catch (error) {
        console.warn(`legacy coverage warning: ${format.id} module could not load`);
        console.warn(String(error instanceof Error ? error.message : error).split("\n")[0]);
      }
    }

    for (const capability of format.capabilities) {
      const exportMissing = module === undefined || !(capability.owner.export in module);
      output.push({
        format: format.id,
        mode: format.mode,
        capability,
        packageMissing,
        moduleMissing,
        exportMissing,
        covered: capability.supported && !packageMissing && !moduleMissing && !exportMissing,
      });
    }
  }
  return output;
}

async function main() {
  const format = filter();
  const selected = await results();
  const visible = format ? selected.filter((result) => result.format === format) : selected;
  const byFormat = new Map<string, CapabilityResult[]>();
  for (const result of visible) {
    const group = byFormat.get(result.format) ?? [];
    group.push(result);
    byFormat.set(result.format, group);
  }

  console.log("======================================================================");
  console.log("Legacy and RTF Capability Coverage");
  console.log("======================================================================");
  console.log("| Format | Mode       | Features     | Coverage |");
  console.log("|--------|------------|--------------|----------|");

  let totalCovered = 0;
  for (const [id, group] of byFormat) {
    const covered = group.filter((result) => result.covered).length;
    totalCovered += covered;
    console.log(
      `| ${id.padEnd(6)} | ${group[0]!.mode.padEnd(10)} | ${(covered + "/" + group.length).padEnd(12)} | ${((covered / group.length) * 100).toFixed(1).padStart(7)}% |`,
    );
  }
  console.log(
    `| **TOTAL** |            | **${totalCovered}/${visible.length}** | **${((totalCovered / visible.length) * 100).toFixed(1)}%** |`,
  );

  if (!process.argv.includes("--summary")) {
    const missing = visible.filter((result) => !result.covered);
    if (missing.length > 0) {
      console.log("\nMissing capabilities:");
      for (const result of missing) {
        const reason = result.packageMissing
          ? "package missing"
          : result.moduleMissing
            ? `module ${result.capability.owner.module} missing`
            : result.exportMissing
              ? `export ${result.capability.owner.export} missing`
              : "registry marks capability unsupported";
        console.log(`  - ${result.format}: ${result.capability.name} (${reason})`);
      }
    }
  }
}

await main();
