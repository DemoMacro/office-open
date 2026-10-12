import { spawn } from "node:child_process";
/**
 * Regression gate for the non-OOXML format corpus.
 *
 * Legacy binary formats are parse-only, RTF is projected to DocumentOptions and
 * verified through a DOCX round trip, while ODF formats round-trip through their
 * own generators. Output is intentionally aggregate-only: third-party corpus
 * paths and document content must not enter CI logs.
 *
 * Requires a prior `pnpm build` because this gate exercises package dist files.
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import type { LegacyFormat } from "./legacy-verify";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "../..");
const CORPUS_DIRS = [
  "Open-XML-SDK",
  "calamine",
  "closedxml",
  "oletools",
  "pandoc",
  "python-pptx",
].map((name) => path.resolve(ROOT_DIR, ".temp/corpus", name));

type Outcome = "pass" | "encrypted" | "invalid" | "unexpected";

interface FormatCounts {
  total: number;
  pass: number;
  encrypted: number;
  invalid: number;
  unexpected: number;
}

const emptyCounts = (): Record<Outcome, number> => ({
  pass: 0,
  encrypted: 0,
  invalid: 0,
  unexpected: 0,
});

const counts: Record<LegacyFormat, FormatCounts> = {
  doc: { total: 0, ...emptyCounts() },
  xls: { total: 0, ...emptyCounts() },
  ppt: { total: 0, ...emptyCounts() },
  rtf: { total: 0, ...emptyCounts() },
  odt: { total: 0, ...emptyCounts() },
  ods: { total: 0, ...emptyCounts() },
  odp: { total: 0, ...emptyCounts() },
};
const failureReasons: Record<LegacyFormat, Map<string, number>> = {
  doc: new Map(),
  xls: new Map(),
  ppt: new Map(),
  rtf: new Map(),
  odt: new Map(),
  ods: new Map(),
  odp: new Map(),
};

function walk(dir: string, out: string[] = []): string[] {
  let entries: fs.Dirent[];
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return out;
  }
  for (const entry of entries) {
    const file = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(file, out);
    else out.push(file);
  }
  return out;
}

interface WorkerResult {
  readonly [format in LegacyFormat]: {
    total: number;
    pass: number;
    encrypted: number;
    invalid: number;
    unexpected: number;
    reasons: Record<string, number>;
  };
}

const emptyFormatResult = () => ({
  total: 0,
  pass: 0,
  encrypted: 0,
  invalid: 0,
  unexpected: 0,
  reasons: {},
});

const emptyWorkerResult = (): WorkerResult => ({
  doc: emptyFormatResult(),
  xls: emptyFormatResult(),
  ppt: emptyFormatResult(),
  rtf: emptyFormatResult(),
  odt: emptyFormatResult(),
  ods: emptyFormatResult(),
  odp: emptyFormatResult(),
});

function runWorker(items: { file: string; format: LegacyFormat }[]): Promise<WorkerResult> {
  return new Promise((resolve, reject) => {
    const child = spawn(
      process.execPath,
      ["--import", "tsx", "--max-old-space-size=4096", path.join(__dirname, "legacy-worker.ts")],
      { stdio: ["pipe", "pipe", "pipe"] },
    );
    const output: Buffer[] = [];
    child.stdout.on("data", (chunk: Buffer) => output.push(chunk));
    child.stderr.on("data", (chunk: Buffer) => process.stderr.write(chunk));
    child.once("error", reject);
    child.once("close", (code) => {
      if (code !== 0) {
        reject(new Error(`legacy corpus verification process exited with ${code}`));
        return;
      }
      resolve(JSON.parse(Buffer.concat(output).toString()) as WorkerResult);
    });
    child.stdin.end(items.map((item) => JSON.stringify(item)).join("\n"));
  });
}

function mergeWorkerResult(workerResult: WorkerResult): void {
  for (const [format, result] of Object.entries(workerResult) as [
    LegacyFormat,
    WorkerResult[LegacyFormat],
  ][]) {
    const formatCounts = counts[format];
    formatCounts.total += result.total;
    formatCounts.pass += result.pass;
    formatCounts.encrypted += result.encrypted;
    formatCounts.invalid += result.invalid;
    formatCounts.unexpected += result.unexpected;
    for (const [reason, count] of Object.entries(result.reasons)) {
      failureReasons[format].set(reason, (failureReasons[format].get(reason) ?? 0) + count);
    }
  }
}

function combineWorkerResults(results: readonly WorkerResult[]): WorkerResult {
  return results.reduce<WorkerResult>((left, right) => {
    for (const [format, result] of Object.entries(right) as [
      LegacyFormat,
      WorkerResult[LegacyFormat],
    ][]) {
      const target = left[format];
      target.total += result.total;
      target.pass += result.pass;
      target.encrypted += result.encrypted;
      target.invalid += result.invalid;
      target.unexpected += result.unexpected;
      for (const [reason, count] of Object.entries(result.reasons)) {
        target.reasons[reason] = (target.reasons[reason] ?? 0) + count;
      }
    }
    return left;
  }, emptyWorkerResult());
}

async function runChunkWithFallback(
  items: {
    file: string;
    format: LegacyFormat;
  }[],
): Promise<WorkerResult> {
  if (items.length === 0) return emptyWorkerResult();
  try {
    return await runWorker(items);
  } catch {
    if (items.length === 1) {
      const result = emptyWorkerResult();
      const formatResult = result[items[0]!.format];
      formatResult.total = 1;
      formatResult.unexpected = 1;
      formatResult.reasons["unexpected: verification process crashed"] = 1;
      return result;
    }
    if (items.length === 0) return emptyWorkerResult();
    const middle = Math.floor(items.length / 2);
    return combineWorkerResults([
      await runChunkWithFallback(items.slice(0, middle)),
      await runChunkWithFallback(items.slice(middle)),
    ]);
  }
}

const files = CORPUS_DIRS.flatMap((dir) => walk(dir)).sort();
async function verifyAllFiles(): Promise<void> {
  const supported = files
    .map((file) => ({
      file,
      format: path.extname(file).slice(1).toLowerCase() as LegacyFormat,
    }))
    .filter((item) => item.format in counts);
  const chunks: { file: string; format: LegacyFormat }[][] = [];
  for (let index = 0; index < supported.length; index += 5) {
    chunks.push(supported.slice(index, index + 5));
  }
  let nextChunk = 0;
  await Promise.all(
    Array.from({ length: Math.min(2, chunks.length) }, async () => {
      while (nextChunk < chunks.length) {
        const chunk = chunks[nextChunk++];
        const result = await runChunkWithFallback(chunk);
        mergeWorkerResult(result);
      }
    }),
  );
}

await verifyAllFiles();

if (!CORPUS_DIRS.every((dir) => fs.existsSync(dir))) {
  console.error("legacy corpus gate: corpus is incomplete — run pnpm corpus:setup first");
  process.exit(1);
}

console.log("format total pass encrypted invalid unexpected");
let failed = false;
for (const [format, result] of Object.entries(counts) as [LegacyFormat, FormatCounts][]) {
  console.log(
    `${format.padEnd(6)} ${String(result.total).padStart(5)} ${String(result.pass).padStart(4)} ${String(result.encrypted).padStart(9)} ${String(result.invalid).padStart(7)} ${String(result.unexpected).padStart(10)}`,
  );
  if (result.unexpected > 0) failed = true;
  for (const [reason, count] of [...failureReasons[format]]
    .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))
    .slice(0, 8)) {
    console.log(`  ${format} ${outcomeForReason(reason)} x${count}: ${reason}`);
  }
  if (result.total === 0) {
    failed = true;
    console.log(`${format.padEnd(6)} empty corpus`);
  }
}

if (failed) {
  console.error("legacy corpus gate: FAILED");
  process.exit(1);
}
console.log("legacy corpus gate: OK");

function outcomeForReason(reason: string): string {
  const [prefix] = reason.split(": ", 1);
  if (prefix === "encrypted" || prefix === "invalid" || prefix === "unexpected") return prefix;
  return "unexpected";
}
