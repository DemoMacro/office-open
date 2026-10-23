import { execSync } from "node:child_process";
import * as fs from "node:fs";
import os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { Worker } from "node:worker_threads";

import type { CorpusFile, FileDiagnostic, Format, FormatCounts, LibraryRunResult } from "./library";
import { walkLibraryFiles } from "./library";
import corpusSources from "./sources.json";
import { runSyntheticCorpus } from "./synthetic";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "../..");
const BASELINE_PATH = path.join(__dirname, "baseline.json");
const DEFAULT_REPORT_PATH = path.resolve(__dirname, "../../.temp/corpus-report.json");
const STRICT_SEMANTIC_GATE = process.env.CORPUS_STRICT_SEMANTIC !== "0";

interface CorpusSource {
  id: string;
  repository: string;
  destination: string;
  commit: string;
}

interface Library {
  id: string;
  repo: string;
  dest: string;
  commit: string;
}

type Baseline = Record<string, Partial<Record<Format, FormatCounts>>>;

const LOCKED_SOURCES = corpusSources.sources as CorpusSource[];
const LIBRARIES = LOCKED_SOURCES.map((source) => ({
  id: source.id,
  repo: source.repository,
  dest: source.destination,
  commit: source.commit,
}));

function isCorpusCheckoutReady(destination: string): boolean {
  if (!fs.existsSync(path.join(destination, ".git"))) return false;
  try {
    execSync("git rev-parse --verify --quiet HEAD", {
      cwd: destination,
      encoding: "utf8",
      stdio: ["ignore", "pipe", "ignore"],
    });
    return true;
  } catch {
    return false;
  }
}

function setup(): void {
  for (const library of LIBRARIES) {
    const destination = path.resolve(ROOT_DIR, library.dest);
    if (fs.existsSync(path.join(destination, ".git"))) {
      let head = "";
      try {
        head = execSync("git rev-parse --verify --quiet HEAD", {
          cwd: destination,
          encoding: "utf8",
          stdio: ["ignore", "pipe", "ignore"],
        }).trim();
      } catch {
        head = "";
      }
      if (!head) {
        console.log(`[setup] ${library.id}: incomplete clone — fetching pinned ${library.commit}`);
        execSync(`git fetch --depth=1 origin "${library.commit}"`, {
          cwd: destination,
          stdio: "inherit",
        });
        execSync("git checkout --quiet FETCH_HEAD", { cwd: destination, stdio: "inherit" });
        continue;
      }
      const current = head;
      if (current === library.commit) {
        console.log(`[setup] ${library.id}: pinned at ${library.commit}`);
        continue;
      }
      console.log(`[setup] ${library.id}: checking out ${library.commit}`);
      execSync(`git fetch --depth=1 origin "${library.commit}"`, {
        cwd: destination,
        stdio: "inherit",
      });
      execSync(`git checkout --quiet "${library.commit}"`, { cwd: destination, stdio: "inherit" });
      continue;
    }
    console.log(`[setup] ${library.id}: fetching pinned ${library.commit}`);
    execSync(`git init --quiet "${destination}"`, { stdio: "inherit" });
    execSync(`git remote add origin "${library.repo}"`, { cwd: destination, stdio: "inherit" });
    execSync(`git fetch --depth=1 origin "${library.commit}"`, {
      cwd: destination,
      stdio: "inherit",
    });
    execSync("git checkout --quiet FETCH_HEAD", { cwd: destination, stdio: "inherit" });
  }
}

function loadBaseline(): Baseline {
  try {
    return JSON.parse(fs.readFileSync(BASELINE_PATH, "utf8")) as Baseline;
  } catch {
    return {};
  }
}

function runLibraryWorker(
  libraryId: string,
  destination: string,
  files: CorpusFile[],
): Promise<LibraryRunResult> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(path.join(__dirname, "library-worker.ts"), {
      workerData: {
        libraryId,
        input: { destination, root: ROOT_DIR, strictSemantic: STRICT_SEMANTIC_GATE },
        files,
      },
      resourceLimits: { maxOldGenerationSizeMb: 4096, maxYoungGenerationSizeMb: 256 },
    });
    worker.once("message", resolve);
    worker.once("error", reject);
    worker.once("exit", (code) => {
      if (code !== 0) reject(new Error(`corpus worker ${libraryId} exited with ${code}`));
    });
  });
}

function splitRawBlocker(
  key: string,
  count: number,
): { reason: string; part: string; count: number } {
  const separator = key.indexOf(":");
  return { reason: key.slice(0, separator), part: key.slice(separator + 1), count };
}

function mergeCountMaps(left: [string, number][], right: [string, number][]): [string, number][] {
  const counts = new Map(left);
  for (const [key, count] of right) counts.set(key, (counts.get(key) ?? 0) + count);
  return [...counts];
}

function mergeLibraryResults(results: LibraryRunResult[]): LibraryRunResult {
  const empty: FormatCounts = {
    total: 0,
    clean: 0,
    diff: 0,
    parseFail: 0,
    genFail: 0,
    malformed: 0,
  };
  const formats: Format[] = ["docx", "xlsx", "pptx"];
  return {
    counts: Object.fromEntries(
      formats.map((format) => {
        const merged = { ...empty };
        for (const result of results) {
          for (const [key, count] of Object.entries(result.counts[format]))
            merged[key as keyof FormatCounts] += count;
        }
        return [format, merged];
      }),
    ) as Record<Format, FormatCounts>,
    blockers: Object.fromEntries(
      formats.map((format) => [
        format,
        results.reduce<[string, number][]>(
          (merged, result) => mergeCountMaps(merged, result.blockers[format]),
          [],
        ),
      ]),
    ) as Record<Format, [string, number][]>,
    diagnostics: results.flatMap((result) => result.diagnostics),
    rawAudit: Object.fromEntries(
      formats.map((format) => [
        format,
        results.reduce(
          (merged, result) => ({
            files: merged.files + result.rawAudit[format].files,
            xmlParts: merged.xmlParts + result.rawAudit[format].xmlParts,
            binaryParts: merged.binaryParts + result.rawAudit[format].binaryParts,
          }),
          { files: 0, xmlParts: 0, binaryParts: 0 },
        ),
      ]),
    ) as Record<Format, { files: number; xmlParts: number; binaryParts: number }>,
    rawBlockers: Object.fromEntries(
      formats.map((format) => [
        format,
        results.reduce<[string, number][]>(
          (merged, result) => mergeCountMaps(merged, result.rawBlockers[format]),
          [],
        ),
      ]),
    ) as Record<Format, [string, number][]>,
  };
}

const args = process.argv.slice(2);
const reportFlagIndex = args.indexOf("--report-json");
const reportPath =
  reportFlagIndex >= 0
    ? path.resolve(ROOT_DIR, args[reportFlagIndex + 1] ?? DEFAULT_REPORT_PATH)
    : DEFAULT_REPORT_PATH;
if (args.includes("--setup")) {
  setup();
  if (args.length === 1) process.exit(0);
}

const only = args.includes("--only") ? args[args.indexOf("--only") + 1] : undefined;
if (only && !LIBRARIES.some((library) => library.id === only)) {
  console.error(
    `unknown library "${only}" — ids: ${LIBRARIES.map((library) => library.id).join(", ")}`,
  );
  process.exit(1);
}

const formatFilter = args.includes("--format") ? args[args.indexOf("--format") + 1] : undefined;
if (formatFilter && !["docx", "xlsx", "pptx"].includes(formatFilter)) {
  console.error(`unknown format "${formatFilter}" — ids: docx, xlsx, pptx`);
  process.exit(1);
}
const formats = (formatFilter ? [formatFilter] : ["docx", "xlsx", "pptx"]) as Format[];

const updateBaseline = args.includes("--update-baseline");
const baseline = loadBaseline();
const nextBaseline = structuredClone(baseline) as Baseline;
let failed = false;
const allDiagnostics: FileDiagnostic[] = [];
const rawBlockerReport: Record<
  string,
  Partial<Record<Format, { reason: string; part: string; count: number }[]>>
> = {};

const configuredWorkers = Number.parseInt(process.env.CORPUS_WORKERS ?? "", 10);
const maxWorkers = Number.isFinite(configuredWorkers)
  ? Math.max(1, Math.min(configuredWorkers, 32))
  : Math.max(1, Math.min(os.availableParallelism() - 1, 8));

console.log("\n[synthetic]");
const synthetic = await runSyntheticCorpus();
if (synthetic.diagnostics.length > 0) failed = true;

const missingLibraries = LIBRARIES.filter(
  (library) => !isCorpusCheckoutReady(path.resolve(ROOT_DIR, library.dest)),
);
if (missingLibraries.length > 0) {
  console.warn(
    `\ncorpus is incomplete (${missingLibraries.length}/${LIBRARIES.length} libraries missing) — synthetic gate only; run pnpm corpus:setup first`,
  );
}

const availableLibraries = LIBRARIES.filter(
  (library) => !missingLibraries.includes(library) && (!only || library.id === only),
);
const resultsByLibrary = new Map<string, LibraryRunResult[]>();
const tasks: { libraryId: string; destination: string; files: CorpusFile[] }[] = [];
let totalFiles = 0;
for (const library of availableLibraries) {
  const files = walkLibraryFiles(path.resolve(ROOT_DIR, library.dest));
  totalFiles += files.length;
  resultsByLibrary.set(library.id, []);
  const workerCount = Math.max(1, Math.min(maxWorkers, Math.ceil(files.length / 32)));
  const chunks = Array.from({ length: workerCount }, (): CorpusFile[] => []);
  files.forEach((file, index) => chunks[index % workerCount]!.push(file));
  for (const files of chunks) {
    if (files.length > 0) tasks.push({ libraryId: library.id, destination: library.dest, files });
  }
}

console.log(
  `\n[external] ${totalFiles} files across ${availableLibraries.length} repos using ${Math.min(maxWorkers, tasks.length)} workers`,
);
let completedTasks = 0;
const nextTask = { index: 0 };
await Promise.all(
  Array.from({ length: Math.min(maxWorkers, tasks.length) }, async () => {
    while (nextTask.index < tasks.length) {
      const task = tasks[nextTask.index++]!;
      try {
        const result = await runLibraryWorker(task.libraryId, task.destination, task.files);
        resultsByLibrary.get(task.libraryId)!.push(result);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        throw new Error(`corpus worker failed for ${task.libraryId}: ${message}`, { cause: error });
      }
      completedTasks += 1;
      if (completedTasks === tasks.length || completedTasks % 10 === 0) {
        console.log(`[external] completed ${completedTasks}/${tasks.length} batches`);
      }
    }
  }),
);

for (const library of availableLibraries) {
  console.log(`\n[${library.id}] ${library.dest}`);
  const result = mergeLibraryResults(resultsByLibrary.get(library.id) ?? []);
  allDiagnostics.push(...result.diagnostics);
  rawBlockerReport[library.id] = {
    ...Object.fromEntries(
      formats.map((format) => [
        format,
        result.rawBlockers[format].map(([key, count]) => splitRawBlocker(key, count)),
      ]),
    ),
  };
  nextBaseline[library.id] = Object.fromEntries(
    formats.map((format) => [format, result.counts[format]]),
  ) as Partial<Record<Format, FormatCounts>>;

  for (const format of formats) {
    const current = result.counts[format];
    console.log(
      `  ${format}  total ${current.total} | clean ${current.clean} | diff ${current.diff} | parseFail ${current.parseFail} | genFail ${current.genFail}`,
    );
    if (current.diff > 0 || current.genFail > 0 || current.parseFail > 0) {
      failed = true;
      console.log(
        `      FAIL valid corpus must round-trip cleanly (parseFail ${current.parseFail}, diff ${current.diff}, genFail ${current.genFail})`,
      );
    }
    const raw = result.rawAudit[format];
    if (raw.files > 0) {
      console.log(
        `      rawParts audit: ${raw.files} files | ${raw.xmlParts} XML parts (absorption gaps) | ${raw.binaryParts} opaque binaries`,
      );
    }
    const rawBlockers = result.rawBlockers[format]
      .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))
      .slice(0, 8);
    for (const [blocker, count] of rawBlockers) {
      console.log(`      rawParts blocker ${blocker}: ${count}`);
    }
    const top = result.blockers[format].sort((left, right) => right[1] - left[1]).slice(0, 8);
    for (const [blocker, count] of top) console.log(`      blocker ${blocker}: ${count}`);

    const expected = baseline[library.id]?.[format];
    if (!expected || expected.total === 0) {
      if (!updateBaseline)
        console.log("      (no baseline — run with --update-baseline to record)");
      continue;
    }
    if (
      current.clean < expected.clean ||
      current.parseFail > expected.parseFail ||
      current.genFail > expected.genFail ||
      current.malformed > (expected.malformed ?? Number.POSITIVE_INFINITY)
    ) {
      failed = true;
      console.log(
        `      FAIL regressed vs baseline (clean ${expected.clean}, parseFail ${expected.parseFail}, genFail ${expected.genFail})`,
      );
    } else if (current.clean > expected.clean) {
      console.log(`      improved: clean ${expected.clean} → ${current.clean} (refresh baseline)`);
    }
  }
}

if (updateBaseline) {
  fs.writeFileSync(BASELINE_PATH, JSON.stringify(nextBaseline, null, 2) + "\n");
  console.log(`\nbaseline written to ${path.relative(ROOT_DIR, BASELINE_PATH)}`);
}

if (reportPath) {
  fs.mkdirSync(path.dirname(reportPath), { recursive: true });
  fs.writeFileSync(
    reportPath,
    `${JSON.stringify(
      {
        diagnostics: allDiagnostics,
        synthetic: synthetic.diagnostics,
        rawBlockers: rawBlockerReport,
      },
      null,
      2,
    )}\n`,
  );
  console.log(`\ncorpus diagnostics: ${path.relative(ROOT_DIR, reportPath)}`);
}

if (failed) {
  console.error("\ncorpus gate: FAILED");
  process.exit(1);
}
if (missingLibraries.length > 0) {
  console.log("\ncorpus gate: SKIPPED (external corpus incomplete)");
  process.exit(0);
}
console.log("\ncorpus gate: OK");
