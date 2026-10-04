import { spawnSync } from "node:child_process";
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { LEGACY_CAPABILITY_REGISTRY } from "./lib/legacy-capability-registry";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

interface TestRun {
  success: boolean;
  testIds: ReadonlySet<string>;
}

interface CapabilityResult {
  format: string;
  capability: (typeof LEGACY_CAPABILITY_REGISTRY)[number]["capabilities"][number];
  packageMissing: boolean;
  moduleMissing: boolean;
  exportMissing: boolean;
  mapperMissing: boolean;
  testMissing: boolean;
  covered: boolean;
}

const testRuns = new Map<string, TestRun>();

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

async function moduleExports(
  packageId: string,
  modulePath: string,
): Promise<Record<string, unknown> | undefined> {
  const file = path.join(ROOT, "packages", packageId, modulePath);
  if (!fs.existsSync(file)) return undefined;
  try {
    return await import(pathToFileURL(file).href);
  } catch {
    return undefined;
  }
}

function runFormatTests(packageId: string, testModule: string): TestRun {
  const cacheKey = `${packageId}:${testModule}`;
  const cached = testRuns.get(cacheKey);
  if (cached) return cached;
  const vp = path.join(
    ROOT,
    "node_modules",
    ".bin",
    process.platform === "win32" ? "vp.ps1" : "vp",
  );
  const testArgs = ["test", "run", testModule, "--reporter=json"];
  const command = process.platform === "win32" ? "powershell.exe" : vp;
  const args =
    process.platform === "win32"
      ? ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", vp, ...testArgs]
      : testArgs;
  const result = spawnSync(command, args, {
    cwd: path.join(ROOT, "packages", packageId),
    encoding: "utf8",
  });
  const output = result.stdout ?? "";
  const jsonStart = output.indexOf("{");
  try {
    const report = JSON.parse(output.slice(jsonStart)) as {
      success?: boolean;
      testResults?: Array<{
        assertionResults?: Array<{ status?: string; fullName?: string; title?: string }>;
      }>;
    };
    const testIds = new Set<string>();
    for (const file of report.testResults ?? []) {
      for (const test of file.assertionResults ?? []) {
        if (test.status === "passed") {
          if (test.fullName) testIds.add(test.fullName);
          if (test.title) testIds.add(test.title);
        }
      }
    }
    const run = { success: result.status === 0 && report.success === true, testIds };
    testRuns.set(cacheKey, run);
    return run;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`legacy coverage: ${packageId} test manifest could not be read: ${message}`);
    console.error(
      `legacy coverage: runner status ${result.status}, error ${result.error?.message ?? "none"}`,
    );
    console.error((result.stderr ?? "").split(/\r?\n/).filter(Boolean).at(-1) ?? "no stderr");
    const run = { success: false, testIds: new Set<string>() };
    testRuns.set(cacheKey, run);
    return run;
  }
}

async function results(): Promise<CapabilityResult[]> {
  const output: CapabilityResult[] = [];
  for (const format of LEGACY_CAPABILITY_REGISTRY) {
    const packageDirectory = path.join(ROOT, "packages", format.id);
    const packageMissing = !fs.existsSync(path.join(packageDirectory, "package.json"));
    const publicModule = await moduleExports(format.id, format.capabilities[0]!.owner.module);
    const mapperModule = await moduleExports(format.id, format.capabilities[0]!.mapper.module);
    const tests = packageMissing ? undefined : runFormatTests(format.id, format.testModule);

    for (const capability of format.capabilities) {
      const formatModule =
        capability.owner.module === format.capabilities[0]!.owner.module
          ? publicModule
          : await moduleExports(format.id, capability.owner.module);
      const currentMapper =
        capability.mapper.module === format.capabilities[0]!.mapper.module
          ? mapperModule
          : await moduleExports(format.id, capability.mapper.module);
      const exportMissing = !formatModule || !(capability.owner.export in formatModule);
      const mapperMissing = !currentMapper || !(capability.mapper.export in currentMapper);
      const testMissing = !tests?.testIds.has(capability.testId);
      const covered =
        capability.supported === true &&
        !packageMissing &&
        !exportMissing &&
        !mapperMissing &&
        !testMissing &&
        tests?.success === true;
      output.push({
        format: format.id,
        capability,
        packageMissing,
        moduleMissing: !formatModule,
        exportMissing,
        mapperMissing,
        testMissing,
        covered,
      });
    }
  }
  return output;
}

async function main(): Promise<void> {
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
  console.log("Legacy and RTF Executable Capability Coverage");
  console.log("======================================================================");
  console.log("| Format | Features | Executable | Coverage |");
  console.log("|--------|----------|------------|----------|");
  let totalCovered = 0;
  for (const [id, group] of byFormat) {
    const covered = group.filter((result) => result.covered).length;
    totalCovered += covered;
    console.log(
      `| ${id.padEnd(6)} | ${String(group.length).padStart(8)} | ${String(covered).padStart(10)} | ${((covered / group.length) * 100).toFixed(1).padStart(7)}% |`,
    );
  }
  console.log(
    `| **TOTAL** | **${String(visible.length).padStart(8)}** | **${String(totalCovered).padStart(10)}** | **${((totalCovered / visible.length) * 100).toFixed(1)}%** |`,
  );

  const missing = visible.filter((result) => !result.covered);
  if (!process.argv.includes("--summary")) {
    for (const result of missing) {
      const reason = result.packageMissing
        ? "package missing"
        : result.moduleMissing
          ? `module ${result.capability.owner.module} missing`
          : result.exportMissing
            ? `export ${result.capability.owner.export} missing`
            : result.mapperMissing
              ? `low-level ${result.capability.mapper.layer} ${result.capability.mapper.module}:${result.capability.mapper.export} missing`
              : result.testMissing
                ? `test ${result.capability.testId} missing or failed`
                : "capability is not proven supported";
      console.log(`  FAIL ${result.format}: ${result.capability.name} (${reason})`);
    }
  }
  if (missing.length > 0) {
    console.error(`legacy coverage gate: FAILED (${missing.length} capabilities)`);
    process.exitCode = 1;
    return;
  }
  console.log("legacy coverage gate: OK");
}

await main();
