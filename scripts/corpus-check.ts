/**
 * Third-party corpus round-trip gate.
 *
 * Real-world Office files from open-source projects are cloned under .temp
 * (gitignored), round-tripped through parse → generate, and the output is
 * compared against the source archive per part. Comparison is prefix-blind
 * tag counting: producers bind arbitrary prefixes to the same namespaces
 * (ClosedXML writes x:workbook), so elements are counted by localname, and a
 * part passes when every localname count matches. theme parts are skipped
 * (tracked separately) and byte-equal parts short-circuit before counting.
 *
 * Baseline gate (scripts/corpus-baseline.json): per library and format,
 * `clean` must not drop and `parseFail`/`genFail` must not rise — absolute
 * counts, so newly added upstream files never mask a regression in files
 * that were already clean. Improvements are reported with a hint to refresh
 * the baseline. A library missing from the baseline is informational only.
 *
 * Usage:
 *   npx tsx scripts/corpus-check.ts --setup        # clone missing libraries (shallow)
 *   npx tsx scripts/corpus-check.ts                # run all libraries
 *   npx tsx scripts/corpus-check.ts --only sdk     # one library
 *   npx tsx scripts/corpus-check.ts --update-baseline   # rewrite baseline from this run
 *
 * Requires a prior `pnpm build` — the runner imports package dist bundles.
 */
import { execSync } from "node:child_process";
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { unzipSync } from "fflate";

// dist imports — deliberate, not a convenience: package sources use internal
// tsconfig aliases (@parts/*, @shared/*) that collide across packages (same
// alias, different roots), and tsx paths cannot route per importing package.
// Running from each package dir (the schema-validate worker pattern) would
// fix that, at the cost of three workers; importing the dist bundles instead
// also means the gate tests the exact artifacts consumers receive.
// Requires a prior `pnpm build`.
import { parseDocument, generateDocument } from "../packages/docx/dist/index.mjs";
import { parsePresentation, generatePresentation } from "../packages/pptx/dist/index.mjs";
import { parseWorkbook, generateWorkbook } from "../packages/xlsx/dist/index.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "..");
const BASELINE_PATH = path.join(__dirname, "corpus-baseline.json");

interface Library {
  /** Stable id used in --only and the baseline file. */
  id: string;
  repo: string;
  /** Clone destination relative to the repo root. */
  dest: string;
}

const LIBRARIES: Library[] = [
  // The primary gate corpus: Microsoft's own generated test assets, spanning
  // docx/pptx/xlsx far beyond what any single format project covers.
  { id: "sdk", repo: "https://github.com/dotnet/Open-XML-SDK", dest: ".temp/corpus/Open-XML-SDK" },
  { id: "calamine", repo: "https://github.com/tafia/calamine.git", dest: ".temp/corpus/calamine" },
  {
    id: "closedxml",
    repo: "https://github.com/ClosedXML/ClosedXML.git",
    dest: ".temp/corpus/closedxml",
  },
  {
    id: "oletools",
    repo: "https://github.com/decalage2/oletools.git",
    dest: ".temp/corpus/oletools",
  },
  { id: "pandoc", repo: "https://github.com/jgm/pandoc.git", dest: ".temp/corpus/pandoc" },
  {
    id: "python-pptx",
    repo: "https://github.com/scanny/python-pptx.git",
    dest: ".temp/corpus/python-pptx",
  },
  { id: "tika", repo: "https://github.com/apache/tika", dest: ".temp/corpus/tika" },
];

type Format = "docx" | "xlsx" | "pptx";

const BY_EXT: Record<string, { format: Format; parse: (b: Uint8Array) => Promise<unknown> }> = {
  docx: { format: "docx", parse: parseDocument },
  dotx: { format: "docx", parse: parseDocument },
  docm: { format: "docx", parse: parseDocument },
  xlsx: { format: "xlsx", parse: parseWorkbook },
  xlsm: { format: "xlsx", parse: parseWorkbook },
  pptx: { format: "pptx", parse: parsePresentation },
  pptm: { format: "pptx", parse: parsePresentation },
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

function walk(dir: string, out: { path: string; format: Format }[] = []): typeof out {
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
      if (cfg) out.push({ path: p, format: cfg.format });
    }
  }
  return out;
}

// ── comparison (same fidelity gate as the historical .temp corpus scripts) ──

const decoder = new TextDecoder();

/** Count elements by localname — prefix-insensitive by design (see header). */
function tagCounts(s: string): Map<string, number> {
  const m = new Map<string, number>();
  for (const match of s.matchAll(/<(?:[\w-]+:)?([\w-]+)[ >/]/g)) {
    m.set(match[1]!, (m.get(match[1]!) ?? 0) + 1);
  }
  return m;
}

function tagCountsDiffer(a: string, b: string): boolean {
  const ca = tagCounts(a);
  const cb = tagCounts(b);
  for (const [t, n] of ca) if ((cb.get(t) ?? 0) !== n) return true;
  for (const [t, n] of cb) if ((ca.get(t) ?? 0) !== n) return true;
  return false;
}

/** All parts whose localname tag counts differ between the two archives. */
function archiveDiffParts(src: Uint8Array, out: Uint8Array): string[] {
  const zs = unzipSync(src);
  const zo = unzipSync(out);
  const parts: string[] = [];
  for (const k of Object.keys(zs)) {
    if (!k.endsWith(".xml") && !k.endsWith(".rels")) continue;
    if (k.includes("theme")) continue;
    const x = decoder.decode(zs[k]!);
    const y = decoder.decode(zo[k] ?? new Uint8Array(0));
    if (x !== y && tagCountsDiffer(x, y)) parts.push(k);
  }
  return parts;
}

// ── runner ──

async function runLibrary(
  lib: Library,
): Promise<Record<Format, FormatCounts> & { blockers: Record<Format, Map<string, number>> }> {
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

  for (const { path: f, format } of walk(path.resolve(ROOT_DIR, lib.dest))) {
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
      console.error(`  parseFail ${path.relative(ROOT_DIR, f)}: ${String(e).slice(0, 120)}`);
      continue;
    }
    try {
      out =
        format === "docx"
          ? await generateDocument(opts as Parameters<typeof generateDocument>[0])
          : format === "xlsx"
            ? await generateWorkbook(opts as Parameters<typeof generateWorkbook>[0])
            : await generatePresentation(opts as Parameters<typeof generatePresentation>[0]);
    } catch (e) {
      a.genFail++;
      console.error(`  genFail ${path.relative(ROOT_DIR, f)}: ${String(e).slice(0, 120)}`);
      continue;
    }
    let parts: string[];
    try {
      parts = archiveDiffParts(new Uint8Array(fs.readFileSync(f)), out);
    } catch {
      a.parseFail++;
      continue;
    }
    if (parts.length === 0) a.clean++;
    else {
      a.diff++;
      for (const part of parts) {
        const key = part.replace(/(?:word|xl|ppt|powerpoint)[\\/]/, "");
        blockers[format].set(key, (blockers[format].get(key) ?? 0) + 1);
      }
    }
  }
  return { ...counts, blockers };
}

// ── setup & baseline ──

function setup(): void {
  for (const lib of LIBRARIES) {
    const dest = path.resolve(ROOT_DIR, lib.dest);
    if (fs.existsSync(path.join(dest, ".git"))) {
      console.log(`[setup] ${lib.id}: already cloned at ${lib.dest}`);
      continue;
    }
    console.log(`[setup] ${lib.id}: cloning ${lib.repo} (shallow) …`);
    execSync(`git clone --depth 1 --quiet "${lib.repo}" "${dest}"`, { stdio: "inherit" });
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

for (const lib of LIBRARIES) {
  if (only && lib.id !== only) continue;
  const dest = path.resolve(ROOT_DIR, lib.dest);
  if (!fs.existsSync(dest)) {
    console.log(`\n[${lib.id}] not cloned — run with --setup first (skipped)`);
    continue;
  }

  console.log(`\n[${lib.id}] ${lib.dest}`);
  const result = await runLibrary(lib);
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

if (failed) {
  console.error("\ncorpus gate: FAILED");
  process.exit(1);
}
console.log("\ncorpus gate: OK");
