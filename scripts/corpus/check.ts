import { execSync } from "node:child_process";
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { Worker } from "node:worker_threads";

import type { FileDiagnostic, Format, FormatCounts, LibraryRunResult } from "./library";
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

function setup(): void {
  for (const library of LIBRARIES) {
    const destination = path.resolve(ROOT_DIR, library.dest);
    if (fs.existsSync(path.join(destination, ".git"))) {
      const current = execSync("git rev-parse HEAD", { cwd: destination, encoding: "utf8" }).trim();
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

function runLibraryWorker(libraryId: string, destination: string): Promise<LibraryRunResult> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(path.join(__dirname, "library-worker.ts"), {
      workerData: {
        libraryId,
        input: { destination, root: ROOT_DIR, strictSemantic: STRICT_SEMANTIC_GATE },
      },
      resourceLimits: { maxOldGenerationSizeMb: 8192, maxYoungGenerationSizeMb: 512 },
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

const updateBaseline = args.includes("--update-baseline");
const baseline = loadBaseline();
const nextBaseline: Baseline = {};
let failed = false;
const allDiagnostics: FileDiagnostic[] = [];
const rawBlockerReport: Record<
  string,
  Record<Format, { reason: string; part: string; count: number }[]>
> = {};

console.log("\n[synthetic]");
const synthetic = await runSyntheticCorpus();
if (synthetic.diagnostics.length > 0) failed = true;

const missingLibraries = LIBRARIES.filter(
  (library) => !fs.existsSync(path.resolve(ROOT_DIR, library.dest)),
);
if (missingLibraries.length > 0) {
  console.warn(
    `\nexternal corpus incomplete (${missingLibraries.length}/${LIBRARIES.length} libraries missing) — synthetic gate only; run pnpm corpus:setup first`,
  );
}

for (const library of LIBRARIES.filter((library) => !missingLibraries.includes(library))) {
  if (only && library.id !== only) continue;
  console.log(`\n[${library.id}] ${library.dest}`);
  const result = await runLibraryWorker(library.id, library.dest);
  allDiagnostics.push(...result.diagnostics);
  rawBlockerReport[library.id] = {
    docx: result.rawBlockers.docx.map(([key, count]) => splitRawBlocker(key, count)),
    xlsx: result.rawBlockers.xlsx.map(([key, count]) => splitRawBlocker(key, count)),
    pptx: result.rawBlockers.pptx.map(([key, count]) => splitRawBlocker(key, count)),
  };
  nextBaseline[library.id] = {
    docx: result.counts.docx,
    xlsx: result.counts.xlsx,
    pptx: result.counts.pptx,
  };

  for (const format of ["docx", "xlsx", "pptx"] as const) {
    const current = result.counts[format];
    console.log(
      `  ${format}  total ${current.total} | clean ${current.clean} | diff ${current.diff} | parseFail ${current.parseFail} | genFail ${current.genFail}`,
    );
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
      current.genFail > expected.genFail
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
console.log("\ncorpus gate: OK");
