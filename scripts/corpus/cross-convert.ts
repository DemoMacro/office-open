import { spawn } from "node:child_process";
import * as fs from "node:fs";
import os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

type Family = "document" | "workbook" | "presentation";
type SourceFormat =
  | "doc"
  | "docx"
  | "docm"
  | "dotx"
  | "dotm"
  | "rtf"
  | "odt"
  | "xls"
  | "xlsx"
  | "xlsm"
  | "xltx"
  | "xltm"
  | "ods"
  | "ppt"
  | "pptx"
  | "pptm"
  | "potx"
  | "potm"
  | "odp";

interface WorkItem {
  readonly id: string;
  readonly file: string;
  readonly source: SourceFormat;
}

interface FamilyResult {
  sources: number;
  parsePass: number;
  encrypted: number;
  invalid: number;
  unexpectedSources: number;
  conversions: number;
  generated: number;
  reparsed: number;
  unsupportedTarget: number;
  generateFailures: number;
  reparseFailures: number;
  writeFailures: number;
  reasons: Record<string, number>;
}

interface WorkerResult {
  readonly document: FamilyResult;
  readonly workbook: FamilyResult;
  readonly presentation: FamilyResult;
}

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "../..");
const CORPUS_DIR = path.resolve(ROOT_DIR, ".temp/corpus");
const OUTPUT_DIR = path.resolve(ROOT_DIR, ".temp/corpus-cross-convert");
const REPORT_PATH = path.resolve(ROOT_DIR, ".temp/corpus-cross-convert-report.json");
const EXTENSIONS: Record<SourceFormat, Family> = {
  doc: "document",
  docx: "document",
  docm: "document",
  dotx: "document",
  dotm: "document",
  rtf: "document",
  odt: "document",
  xls: "workbook",
  xlsx: "workbook",
  xlsm: "workbook",
  xltx: "workbook",
  xltm: "workbook",
  ods: "workbook",
  ppt: "presentation",
  pptx: "presentation",
  pptm: "presentation",
  potx: "presentation",
  potm: "presentation",
  odp: "presentation",
};
const FAMILY_NAMES: readonly Family[] = ["document", "workbook", "presentation"];

function walk(root: string, files: WorkItem[] = []): WorkItem[] {
  let entries: fs.Dirent[];
  try {
    entries = fs.readdirSync(root, { withFileTypes: true });
  } catch {
    return files;
  }
  for (const entry of entries) {
    const item = path.join(root, entry.name);
    if (entry.isDirectory()) {
      if (entry.name !== ".git" && entry.name !== "node_modules") walk(item, files);
      continue;
    }
    const extension = path.extname(entry.name).slice(1).toLowerCase() as SourceFormat;
    const family = EXTENSIONS[extension];
    if (!family) continue;
    files.push({
      id: path.relative(ROOT_DIR, item).replaceAll("\\", "/"),
      file: item,
      source: extension,
    });
  }
  return files;
}

function emptyFamily(): FamilyResult {
  return {
    sources: 0,
    parsePass: 0,
    encrypted: 0,
    invalid: 0,
    unexpectedSources: 0,
    conversions: 0,
    generated: 0,
    reparsed: 0,
    unsupportedTarget: 0,
    generateFailures: 0,
    reparseFailures: 0,
    writeFailures: 0,
    reasons: {},
  };
}

function emptyResult(): WorkerResult {
  return { document: emptyFamily(), workbook: emptyFamily(), presentation: emptyFamily() };
}

function mergeCounts(target: Record<string, number>, source: Record<string, number>): void {
  for (const [reason, count] of Object.entries(source))
    target[reason] = (target[reason] ?? 0) + count;
}

function merge(target: WorkerResult, source: WorkerResult): void {
  for (const family of FAMILY_NAMES) {
    const left = target[family];
    const right = source[family];
    const keys: readonly (keyof Omit<FamilyResult, "reasons">)[] = [
      "sources",
      "parsePass",
      "encrypted",
      "invalid",
      "unexpectedSources",
      "conversions",
      "generated",
      "reparsed",
      "unsupportedTarget",
      "generateFailures",
      "reparseFailures",
      "writeFailures",
    ];
    for (const key of keys) {
      left[key] += right[key];
    }
    mergeCounts(left.reasons, right.reasons);
  }
}

function runWorker(items: readonly WorkItem[], write: boolean): Promise<WorkerResult> {
  if (items.length === 0) return Promise.resolve(emptyResult());
  return new Promise((resolve, reject) => {
    const child = spawn(
      process.execPath,
      [
        "--import",
        "tsx",
        "--max-old-space-size=4096",
        path.join(__dirname, "cross-convert-worker.ts"),
      ],
      { stdio: ["pipe", "pipe", "pipe"] },
    );
    const output: Buffer[] = [];
    child.stdout.on("data", (chunk: Buffer) => output.push(chunk));
    child.stderr.on("data", (chunk: Buffer) => process.stderr.write(chunk));
    child.once("error", reject);
    child.once("close", (code) => {
      if (code !== 0) {
        reject(new Error(`cross-conversion worker exited with ${code}`));
        return;
      }
      resolve(JSON.parse(Buffer.concat(output).toString()) as WorkerResult);
    });
    child.stdin.end(JSON.stringify({ outputRoot: write ? OUTPUT_DIR : undefined, items }));
  });
}

async function runWithFallback(items: readonly WorkItem[], write: boolean): Promise<WorkerResult> {
  try {
    return await runWorker(items, write);
  } catch (error) {
    process.stderr.write(`retrying batch of ${items.length}: ${String(error)}\n`);
    if (items.length === 1) {
      const result = emptyResult();
      const family = EXTENSIONS[items[0]!.source];
      result[family].sources = 1;
      result[family].invalid = 1;
      result[family].reasons["invalid: worker crashed on source"] = 1;
      return result;
    }
    const middle = Math.floor(items.length / 2);
    const merged = emptyResult();
    merge(merged, await runWithFallback(items.slice(0, middle), write));
    merge(merged, await runWithFallback(items.slice(middle), write));
    return merged;
  }
}

if (!fs.existsSync(CORPUS_DIR)) {
  console.error("corpus cross-conversion gate: corpus is incomplete — run pnpm corpus:setup first");
  process.exit(1);
}

const files = walk(CORPUS_DIR).sort((left, right) => left.id.localeCompare(right.id));
const familyArgument = process.argv.indexOf("--family");
const requestedFamily = familyArgument === -1 ? undefined : process.argv[familyArgument + 1];
if (requestedFamily && !FAMILY_NAMES.includes(requestedFamily as Family))
  throw new Error(`Unknown family: ${requestedFamily}`);
const sourceArgument = process.argv.indexOf("--source");
const requestedSource = sourceArgument === -1 ? undefined : process.argv[sourceArgument + 1];
if (requestedSource && !(requestedSource in EXTENSIONS))
  throw new Error(`Unknown source format: ${requestedSource}`);
const selected = files.filter(
  (file) =>
    (!requestedFamily || EXTENSIONS[file.source] === requestedFamily) &&
    (!requestedSource || file.source === requestedSource),
);
const write = process.argv.includes("--write");
if (write) {
  fs.rmSync(OUTPUT_DIR, { recursive: true, force: true });
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

const batchArgument = process.argv.indexOf("--batch-size");
const requestedBatchSize = batchArgument === -1 ? 8 : Number(process.argv[batchArgument + 1]);
const batchSize =
  Number.isFinite(requestedBatchSize) && requestedBatchSize > 0
    ? Math.floor(requestedBatchSize)
    : 8;
const batches: WorkItem[][] = [];
for (let index = 0; index < selected.length; index += batchSize)
  batches.push(selected.slice(index, index + batchSize));

const result = emptyResult();
let completed = 0;
let nextBatch = 0;
const concurrency = Math.max(1, Math.min(4, os.availableParallelism?.() ?? os.cpus().length));

await Promise.all(
  Array.from({ length: concurrency }, async () => {
    while (nextBatch < batches.length) {
      const batch = batches[nextBatch++]!;
      merge(result, await runWithFallback(batch, write));
      completed += batch.length;
      if (completed % 160 === 0 || completed === selected.length)
        process.stderr.write(`cross-conversion progress ${completed}/${selected.length}\n`);
    }
  }),
);

const summary = {
  totals: Object.fromEntries(FAMILY_NAMES.map((family) => [family, result[family]])),
  failures: Object.fromEntries(
    FAMILY_NAMES.map((family) => [
      family,
      [...Object.entries(result[family].reasons)]
        .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))
        .slice(0, 60),
    ]),
  ),
};
fs.writeFileSync(REPORT_PATH, `${JSON.stringify(summary, null, 2)}\n`);

console.log(
  "family sources parse encrypted invalid unexpected conversions generated reparsed unsupported failed",
);
let failed = false;
for (const family of FAMILY_NAMES) {
  const item = result[family];
  const unexpected =
    item.unexpectedSources + item.generateFailures + item.reparseFailures + item.writeFailures;
  console.log(
    `${family.padEnd(12)} ${item.sources} ${item.parsePass} ${item.encrypted} ${item.invalid} ${unexpected} ${item.conversions} ${item.generated} ${item.reparsed} ${item.unsupportedTarget} ${unexpected}`,
  );
  if (
    item.writeFailures > 0 ||
    item.reparseFailures > 0 ||
    item.unexpectedSources > 0 ||
    item.generated !== item.conversions - item.unsupportedTarget - item.generateFailures
  )
    failed = true;
}
console.log(`report: ${path.relative(ROOT_DIR, REPORT_PATH)}`);
if (failed) {
  console.error("corpus cross-conversion gate: FAILED");
  process.exit(1);
}
console.log("corpus cross-conversion gate: OK");
