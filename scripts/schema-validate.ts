/**
 * Validate Options JSON against the generated JSON Schemas.
 *
 * Corpus:
 *   A. office-open/demo/*.json — hand-written options, validated directly.
 *   B. Round-trip: every package demo is executed (unique .temp outputs),
 *      parsed back into Options JSON by schema-roundtrip-worker.ts (running in
 *      the package dir so tsconfig aliases resolve to source), and validated.
 *
 * Failures are either schema gaps (fix in generate-schema.ts post-processing
 * or the source types) or genuine options/type drift — both must end at zero.
 *
 * Usage:
 *   npx tsx scripts/schema-validate.ts            # full corpus
 *   npx tsx scripts/schema-validate.ts docx       # one format only
 */
import { execFile } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import Ajv, { type DefinedError } from "ajv";
import addFormats from "ajv-formats";

const execFileAsync = promisify(execFile);

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "..");
const SCHEMA_DIR = path.join(ROOT_DIR, "packages/office-open/schemas");
const WORKER = path.join(__dirname, "schema-roundtrip-worker.ts");

interface FormatConfig {
  format: "docx" | "pptx" | "xlsx";
  dir: string;
  ext: string;
  /** Demo JSON in office-open/demo keyed by leading number. */
  demoJson?: string;
}

const FORMATS: FormatConfig[] = [
  { format: "docx", dir: "packages/docx", ext: ".docx", demoJson: "1-docx.json" },
  { format: "pptx", dir: "packages/pptx", ext: ".pptx", demoJson: "2-pptx.json" },
  { format: "xlsx", dir: "packages/xlsx", ext: ".xlsx", demoJson: "3-xlsx.json" },
];

/** Max error lines printed per failing file — enough to diagnose, not to drown. */
const MAX_ERRORS_PER_FILE = 8;

interface CorpusCase {
  label: string;
  options: unknown;
}

interface FileResult {
  file: string;
  ok: boolean;
  options?: unknown;
  error?: string;
}

async function runPool<T>(items: T[], limit: number, fn: (item: T) => Promise<void>) {
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const item = items[next++]!;
      await fn(item);
    }
  });
  await Promise.all(workers);
}

function listDemoArtifacts(config: FormatConfig): string[] {
  const tempDir = path.join(ROOT_DIR, config.dir, ".temp");
  if (!fs.existsSync(tempDir)) return [];
  return fs
    .readdirSync(tempDir)
    .filter((f) => f.endsWith(config.ext))
    .sort()
    .map((f) => path.join(tempDir, f));
}

/** Run all package demos (they write unique .temp/<stem><ext> outputs). */
async function runDemos(config: FormatConfig, failures: string[]): Promise<void> {
  const demoDir = path.join(ROOT_DIR, config.dir, "demo");
  const demos = fs
    .readdirSync(demoDir)
    .filter((f) => f.endsWith(".ts"))
    .sort();
  await runPool(demos, Math.min(8, os.cpus().length), async (demo) => {
    try {
      await execFileAsync("npx", ["tsx", `demo/${demo}`], {
        cwd: path.join(ROOT_DIR, config.dir),
        windowsHide: true,
        shell: true, // Windows: npx is a .cmd shim execFile cannot spawn directly
      });
    } catch {
      failures.push(`${config.format} demo RUN FAIL: ${demo}`);
    }
  });
}

/** Parse every artifact back into Options JSON (one worker process per package). */
async function parseArtifacts(config: FormatConfig, artifacts: string[]): Promise<FileResult[]> {
  const pkgDir = path.join(ROOT_DIR, config.dir);
  // Relative paths keep the command line under the Windows cmd limit; the
  // shell:true quoting (npx .cmd shim) is needed for "… (round-trip)" names.
  const quoted = artifacts.map((f) => `"${path.relative(pkgDir, f)}"`);
  const { stdout } = await execFileAsync("npx", ["tsx", WORKER, config.format, ...quoted], {
    cwd: pkgDir,
    windowsHide: true,
    shell: true,
    maxBuffer: 512 * 1024 * 1024,
  });
  return stdout
    .split("\n")
    .filter((line) => line.startsWith("{"))
    .map((line) => JSON.parse(line) as FileResult);
}

function formatAjvErrors(errors: DefinedError[], label: string): string[] {
  const lines = [`FAIL ${label}`];
  for (const err of errors) {
    lines.push(
      `  ${err.instancePath || "/"} ${err.message ?? ""}` +
        (err.params && Object.keys(err.params).length > 0 ? ` ${JSON.stringify(err.params)}` : ""),
    );
    if (lines.length - 1 >= MAX_ERRORS_PER_FILE) {
      lines.push(`  … ${errors.length - MAX_ERRORS_PER_FILE} more`);
      break;
    }
  }
  return lines;
}

async function main(): Promise<void> {
  const only = process.argv[2];
  // allowUnionTypes: tsj emits `type: ["string","number"]` for number|string
  // fields; formats: tsj carries JSDoc @format through (e.g. date-time).
  const ajv = new Ajv({ allErrors: true, verbose: false, allowUnionTypes: true });
  addFormats(ajv);
  let totalFail = 0;

  for (const config of FORMATS) {
    if (only && only !== config.format) continue;
    process.stderr.write(`\n--- ${config.format} ---\n`);

    const schema = JSON.parse(
      fs.readFileSync(path.join(SCHEMA_DIR, `${config.format}.schema.json`), "utf-8"),
    );
    const validateFn = ajv.compile(schema);

    // Corpus A: hand-written demo options.
    const cases: CorpusCase[] = [];
    if (config.demoJson) {
      const demoPath = path.join(ROOT_DIR, "packages/office-open/demo", config.demoJson);
      if (fs.existsSync(demoPath)) {
        cases.push({
          label: `office-open/demo/${config.demoJson}`,
          options: JSON.parse(fs.readFileSync(demoPath, "utf-8")),
        });
      }
    }

    // Corpus B: generate + parse back every package demo artifact.
    const runFailures: string[] = [];
    await runDemos(config, runFailures);
    for (const f of runFailures) console.error(`  ${f}`);
    totalFail += runFailures.length;

    const artifacts = listDemoArtifacts(config);
    const parsed = await parseArtifacts(config, artifacts);
    for (const result of parsed) {
      if (!result.ok) {
        console.error(`  PARSE FAIL ${path.basename(result.file)}: ${result.error}`);
        totalFail++;
        continue;
      }
      cases.push({
        label: `${config.dir}/.temp/${path.basename(result.file)}`,
        options: result.options,
      });
    }

    // Validate everything collected for this format.
    let pass = 0;
    let failed = 0;
    const lines: string[] = [];
    for (const item of cases) {
      if (validateFn(item.options)) {
        pass++;
      } else {
        failed++;
        // Snapshot immediately — the next validation overwrites ajv state.
        lines.push(...formatAjvErrors(validateFn.errors ?? [], item.label));
      }
    }
    totalFail += failed;
    for (const line of lines) console.log(line);
    process.stderr.write(`  ${pass} pass, ${failed} fail (${cases.length} cases)\n`);
  }

  if (totalFail > 0) {
    console.error(`\n${totalFail} failure(s)`);
    process.exit(1);
  }
  console.error("\nall cases valid");
}

await main();
