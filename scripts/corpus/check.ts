/**
 * Third-party corpus round-trip gate.
 *
 * Real-world Office files are pinned by commit in corpus/sources.json, cloned
 * under gitignored .temp, round-tripped through parse → generate, and compared
 * with a namespace-aware canonical semantic digest. XML nodes, attributes,
 * text, relationships, content types, and binary payload identity are checked.
 *
 * Baseline gate (scripts/corpus/baseline.json): per library and format,
 * `clean` must not drop and `parseFail`/`genFail` must not rise — absolute
 * counts, so newly added upstream files never mask a regression in files
 * that were already clean. Improvements are reported with a hint to refresh
 * the baseline. A library missing from the baseline is informational only.
 *
 * Usage:
 *   npx tsx scripts/corpus/check.ts --setup        # clone missing libraries (shallow)
 *   npx tsx scripts/corpus/check.ts                # run all libraries
 *   npx tsx scripts/corpus/check.ts --only sdk     # one library
 *   npx tsx scripts/corpus/check.ts --update-baseline   # rewrite baseline from this run
 *
 * Requires a prior `pnpm build` — the runner imports package dist bundles.
 */
import { execSync } from "node:child_process";
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

// dist imports — deliberate, not a convenience: package sources use internal
// tsconfig aliases (@parts/*, @shared/*) that collide across packages (same
// alias, different roots), and tsx paths cannot route per importing package.
// Running from each package dir (the schema-validate worker pattern) would
// fix that, at the cost of three workers; importing the dist bundles instead
// also means the gate tests the exact artifacts consumers receive.
// Requires a prior `pnpm build`.
import { OOXML_PACKAGE_FORMATS } from "../../packages/core/dist/index.mjs";
import { parseDocument, generateDocument } from "../../packages/docx/dist/index.mjs";
import { parsePresentation, generatePresentation } from "../../packages/pptx/dist/index.mjs";
import { parseWorkbook, generateWorkbook } from "../../packages/xlsx/dist/index.mjs";
import {
  archiveSemanticDiffDetails,
  archiveTagDiffs,
  classifyPackageFailure,
  type CorpusFailureKind,
  type SemanticPartDiff,
} from "./semantics";
import corpusSources from "./sources.json";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "../..");
const BASELINE_PATH = path.join(__dirname, "baseline.json");
const DEFAULT_REPORT_PATH = path.resolve(__dirname, "../../.temp/corpus-report.json");
const STRICT_SEMANTIC_GATE = process.env.CORPUS_STRICT_SEMANTIC === "1";

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

const LOCKED_SOURCES = corpusSources.sources as CorpusSource[];

const LIBRARIES = LOCKED_SOURCES.map((source) => ({
  id: source.id,
  repo: source.repository,
  dest: source.destination,
  commit: source.commit,
}));

type Format = "docx" | "xlsx" | "pptx";
type PackageFormat = keyof typeof OOXML_PACKAGE_FORMATS;

const BY_EXT: Record<
  string,
  { format: Format; type: PackageFormat; parse: (b: Uint8Array) => Promise<unknown> }
> = {
  docx: { format: "docx", type: "docx", parse: parseDocument },
  docm: { format: "docx", type: "docm", parse: parseDocument },
  dotx: { format: "docx", type: "dotx", parse: parseDocument },
  dotm: { format: "docx", type: "dotm", parse: parseDocument },
  xlsx: { format: "xlsx", type: "xlsx", parse: parseWorkbook },
  xlsm: { format: "xlsx", type: "xlsm", parse: parseWorkbook },
  xltx: { format: "xlsx", type: "xltx", parse: parseWorkbook },
  xltm: { format: "xlsx", type: "xltm", parse: parseWorkbook },
  pptx: { format: "pptx", type: "pptx", parse: parsePresentation },
  pptm: { format: "pptx", type: "pptm", parse: parsePresentation },
  potx: { format: "pptx", type: "potx", parse: parsePresentation },
  potm: { format: "pptx", type: "potm", parse: parsePresentation },
};

interface FormatCounts {
  total: number;
  clean: number;
  diff: number;
  parseFail: number;
  genFail: number;
}

type Baseline = Record<string, Partial<Record<Format, FormatCounts>>>;

// ── corpus discovery ──

function walk(
  dir: string,
  out: { path: string; format: Format; type: PackageFormat }[] = [],
): typeof out {
  let entries: fs.Dirent[];
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return out;
  }
  for (const e of entries) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walk(p, out);
    else {
      const cfg = BY_EXT[e.name.split(".").pop()!.toLowerCase()];
      if (cfg) out.push({ path: p, format: cfg.format, type: cfg.type });
    }
  }
  return out;
}

// ── runner ──

export interface FileDiagnostic {
  file: string;
  format: Format;
  outcome: CorpusFailureKind;
  package?: string;
  diffCategories?: Record<string, number>;
  sampleParts?: SemanticPartDiff[];
  error?: string;
}

async function runLibrary(lib: Library): Promise<
  Record<Format, FormatCounts> & {
    blockers: Record<Format, Map<string, number>>;
    diagnostics: FileDiagnostic[];
  }
> {
  const counts = {
    docx: { total: 0, clean: 0, diff: 0, parseFail: 0, genFail: 0 },
    xlsx: { total: 0, clean: 0, diff: 0, parseFail: 0, genFail: 0 },
    pptx: { total: 0, clean: 0, diff: 0, parseFail: 0, genFail: 0 },
  } as Record<Format, FormatCounts>;
  const blockers: Record<Format, Map<string, number>> = {
    docx: new Map(),
    xlsx: new Map(),
    pptx: new Map(),
  };
  const diagnostics: FileDiagnostic[] = [];

  for (const { path: f, format, type } of walk(path.resolve(ROOT_DIR, lib.dest))) {
    const a = counts[format];
    a.total++;
    let opts: unknown;
    let out: Uint8Array;
    try {
      opts = await BY_EXT[path.extname(f).slice(1).toLowerCase()]!.parse(
        new Uint8Array(fs.readFileSync(f)),
      );
    } catch (e) {
      a.parseFail++;
      diagnostics.push({
        file: path.relative(ROOT_DIR, f),
        format,
        outcome: classifyPackageFailure(e),
        package: lib.id,
        error: String(e).slice(0, 240),
      });
      continue;
    }
    try {
      const packerOptions = { packageVariant: OOXML_PACKAGE_FORMATS[type].variant };
      out =
        format === "docx"
          ? await generateDocument(opts as Parameters<typeof generateDocument>[0], packerOptions)
          : format === "xlsx"
            ? await generateWorkbook(opts as Parameters<typeof generateWorkbook>[0], packerOptions)
            : await generatePresentation(
                opts as Parameters<typeof generatePresentation>[0],
                packerOptions,
              );
    } catch (e) {
      a.genFail++;
      diagnostics.push({
        file: path.relative(ROOT_DIR, f),
        format,
        outcome: "invalid-package",
        package: lib.id,
        error: String(e).slice(0, 240),
      });
      continue;
    }
    let parts: string[];
    let semanticDiffs: ReturnType<typeof archiveSemanticDiffDetails> = [];
    try {
      const source = new Uint8Array(fs.readFileSync(f));
      if (STRICT_SEMANTIC_GATE) {
        semanticDiffs = archiveSemanticDiffDetails(source, out);
        parts = semanticDiffs.map((diff) => diff.path);
      } else {
        parts = archiveTagDiffs(source, out);
      }
    } catch (e) {
      a.parseFail++;
      diagnostics.push({
        file: path.relative(ROOT_DIR, f),
        format,
        outcome: "invalid-zip",
        package: lib.id,
        error: String(e).slice(0, 240),
      });
      continue;
    }
    if (parts.length === 0) a.clean++;
    else {
      a.diff++;
      for (const part of parts) {
        const key = part.replace(/(?:word|xl|ppt|powerpoint)[\\/]/, "");
        blockers[format].set(key, (blockers[format].get(key) ?? 0) + 1);
      }
      diagnostics.push({
        file: path.relative(ROOT_DIR, f),
        format,
        outcome: "valid",
        package: lib.id,
        diffCategories: semanticDiffs.reduce<Record<string, number>>((counts, diff) => {
          const category = diff.category ?? "unknown";
          counts[category] = (counts[category] ?? 0) + 1;
          return counts;
        }, {}),
        sampleParts: semanticDiffs.slice(0, 20),
      });
    }
  }
  return { ...counts, blockers, diagnostics };
}

// ── setup & baseline ──

function setup(): void {
  for (const lib of LIBRARIES) {
    const dest = path.resolve(ROOT_DIR, lib.dest);
    if (fs.existsSync(path.join(dest, ".git"))) {
      const current = execSync("git rev-parse HEAD", { cwd: dest, encoding: "utf8" }).trim();
      if (current === lib.commit) {
        console.log(`[setup] ${lib.id}: pinned at ${lib.commit}`);
        continue;
      }
      console.log(`[setup] ${lib.id}: checking out ${lib.commit}`);
      execSync(`git fetch --depth=1 origin "${lib.commit}"`, { cwd: dest, stdio: "inherit" });
      execSync(`git checkout --quiet "${lib.commit}"`, { cwd: dest, stdio: "inherit" });
      continue;
    }
    console.log(`[setup] ${lib.id}: fetching pinned ${lib.commit}`);
    execSync(`git init --quiet "${dest}"`, { stdio: "inherit" });
    execSync(`git remote add origin "${lib.repo}"`, { cwd: dest, stdio: "inherit" });
    execSync(`git fetch --depth=1 origin "${lib.commit}"`, { cwd: dest, stdio: "inherit" });
    execSync("git checkout --quiet FETCH_HEAD", { cwd: dest, stdio: "inherit" });
  }
}

function loadBaseline(): Baseline {
  try {
    return JSON.parse(fs.readFileSync(BASELINE_PATH, "utf8")) as Baseline;
  } catch {
    return {};
  }
}

// ── main ──

const args = process.argv.slice(2);
const reportFlagIndex = args.indexOf("--report-json");
const reportPath =
  reportFlagIndex >= 0
    ? path.resolve(ROOT_DIR, args[reportFlagIndex + 1] ?? DEFAULT_REPORT_PATH)
    : undefined;
if (args.includes("--setup")) {
  setup();
  if (args.length === 1) process.exit(0);
}

const only = args.includes("--only") ? args[args.indexOf("--only") + 1] : undefined;
if (only && !LIBRARIES.some((l) => l.id === only)) {
  console.error(`unknown library "${only}" — ids: ${LIBRARIES.map((l) => l.id).join(", ")}`);
  process.exit(1);
}

const updateBaseline = args.includes("--update-baseline");
const baseline = loadBaseline();
const nextBaseline: Baseline = {};
let failed = false;
const allDiagnostics: FileDiagnostic[] = [];

if (LIBRARIES.some((lib) => !fs.existsSync(path.resolve(ROOT_DIR, lib.dest)))) {
  console.error("corpus gate: corpus is incomplete — run pnpm corpus:setup first");
  process.exit(1);
}

for (const lib of LIBRARIES) {
  if (only && lib.id !== only) continue;
  const dest = path.resolve(ROOT_DIR, lib.dest);
  console.log(`\n[${lib.id}] ${lib.dest}`);
  const result = await runLibrary(lib);
  allDiagnostics.push(...result.diagnostics);
  const libBaseline = baseline[lib.id];
  nextBaseline[lib.id] = { docx: result.docx, xlsx: result.xlsx, pptx: result.pptx };

  for (const format of ["docx", "xlsx", "pptx"] as const) {
    const a = result[format];
    console.log(
      `  ${format}  total ${a.total} | clean ${a.clean} | diff ${a.diff} | parseFail ${a.parseFail} | genFail ${a.genFail}`,
    );
    const top = [...result.blockers[format]].sort((x, y) => y[1] - x[1]).slice(0, 8);
    for (const [k, n] of top) console.log(`      blocker ${k}: ${n}`);

    const b = libBaseline?.[format];
    if (!b || b.total === 0) {
      if (!updateBaseline)
        console.log(`      (no baseline — run with --update-baseline to record)`);
      continue;
    }
    if (a.clean < b.clean || a.parseFail > b.parseFail || a.genFail > b.genFail) {
      failed = true;
      console.log(
        `      FAIL regressed vs baseline (clean ${b.clean}, parseFail ${b.parseFail}, genFail ${b.genFail})`,
      );
    } else if (a.clean > b.clean) {
      console.log(`      improved: clean ${b.clean} → ${a.clean} (refresh baseline)`);
    }
  }
}

if (updateBaseline) {
  fs.writeFileSync(BASELINE_PATH, JSON.stringify(nextBaseline, null, 2) + "\n");
  console.log(`\nbaseline written to ${path.relative(ROOT_DIR, BASELINE_PATH)}`);
}

if (reportPath) {
  fs.mkdirSync(path.dirname(reportPath), { recursive: true });
  fs.writeFileSync(reportPath, `${JSON.stringify({ diagnostics: allDiagnostics }, null, 2)}\n`);
  console.log(`\ncorpus diagnostics: ${path.relative(ROOT_DIR, reportPath)}`);
}

if (failed) {
  console.error("\ncorpus gate: FAILED");
  process.exit(1);
}
console.log("\ncorpus gate: OK");
