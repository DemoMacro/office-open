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

import { parseDocument as parseLegacyDocument } from "../../packages/doc/dist/index.mjs";
import { generateDocument, parseDocument } from "../../packages/docx/dist/index.mjs";
import {
  generatePresentation as generateOdpPresentation,
  parsePresentation as parseOdpPresentation,
} from "../../packages/odp/dist/index.mjs";
import {
  generateWorkbook as generateOdsWorkbook,
  parseWorkbook as parseOdsWorkbook,
} from "../../packages/ods/dist/index.mjs";
import {
  generateDocument as generateOdtDocument,
  parseDocument as parseOdtDocument,
} from "../../packages/odt/dist/index.mjs";
import { parsePresentation as parseLegacyPresentation } from "../../packages/ppt/dist/index.mjs";
import {
  generatePresentation as generatePptxPresentation,
  parsePresentation,
} from "../../packages/pptx/dist/index.mjs";
import { parseDocument as parseRtfDocument } from "../../packages/rtf/dist/index.mjs";
import { parseWorkbook as parseLegacyWorkbook } from "../../packages/xls/dist/index.mjs";
import { generateWorkbook as generateXlsxWorkbook } from "../../packages/xlsx/dist/index.mjs";
import { parseWorkbook as parseXlsxDocument } from "../../packages/xlsx/dist/index.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "../..");
const CORPUS_DIRS = [
  "Open-XML-SDK",
  "calamine",
  "closedxml",
  "oletools",
  "pandoc",
  "python-pptx",
  "tika",
].map((name) => path.resolve(ROOT_DIR, ".temp/corpus", name));

type LegacyFormat = "doc" | "xls" | "ppt" | "rtf" | "odt" | "ods" | "odp";
type Outcome = "pass" | "encrypted" | "invalid" | "unexpected";

const CORPUS_PASSWORDS = ["tika", "VelvetSweatshop", "Password1234_", "password"];

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

function failureCategory(error: unknown): string {
  return String((error as Error)?.message ?? error)
    .replace(/\s+/g, " ")
    .replace(/(["']).*?\1/g, "$1…$1")
    .replace(/\b[0-9a-f]{8,}\b/gi, "…")
    .replace(/\b\d+(?:\.\d+)?\b/g, "N")
    .trim()
    .slice(0, 160);
}
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

function classify(error: unknown): Exclude<Outcome, "pass"> {
  const message = String((error as Error)?.message ?? error);
  if (/encrypted/i.test(message)) return "encrypted";
  if (
    /(?:bad signature|truncated|must begin with|invalid CFB|invalid ODF|invalid zip|invalid URL|not a supported Compound|signature is not|unsupported RTF destination|objdata requires binary|Manifest does not declare|Manifest declares missing package path|embedded object subdocument is missing)/i.test(
      message,
    )
  ) {
    return "invalid";
  }
  return "unexpected";
}

async function verify(format: LegacyFormat, data: Uint8Array): Promise<void> {
  if (format === "doc") {
    for (const password of [undefined, ...CORPUS_PASSWORDS]) {
      try {
        const options = parseLegacyDocument(data, password ? { password } : undefined);
        parseDocument(await generateDocument(options));
        return;
      } catch (error) {
        if (
          password === undefined &&
          !/encrypted|password/i.test(String((error as Error)?.message))
        ) {
          throw error;
        }
      }
    }
    throw new Error("Encrypted legacy DOC input could not be decrypted");
  }
  if (format === "xls") {
    for (const password of [undefined, ...CORPUS_PASSWORDS]) {
      try {
        parseXlsxDocument(
          await generateXlsxWorkbook(
            parseLegacyWorkbook(data, password ? { password } : undefined),
          ),
        );
        return;
      } catch (error) {
        if (
          password === undefined &&
          !/encrypted|password/i.test(String((error as Error)?.message))
        ) {
          throw error;
        }
      }
    }
    throw new Error("Encrypted legacy XLS input could not be decrypted");
  }
  if (format === "ppt") {
    for (const password of [undefined, ...CORPUS_PASSWORDS]) {
      try {
        const options = parseLegacyPresentation(data, password ? { password } : undefined);
        parsePresentation(await generatePptxPresentation(options));
        return;
      } catch (error) {
        if (
          password === undefined &&
          !/encrypted|password/i.test(String((error as Error)?.message))
        ) {
          throw error;
        }
      }
    }
    throw new Error("Encrypted legacy PPT input could not be decrypted");
  }
  if (format === "rtf") {
    const options = parseRtfDocument(Buffer.from(data).toString("latin1"));
    parseDocument(await generateDocument(options));
    return;
  }
  if (format === "odt") {
    parseOdtDocument(generateOdtDocument(parseOdtDocument(data)));
    return;
  }
  if (format === "ods") {
    parseOdsWorkbook(generateOdsWorkbook(parseOdsWorkbook(data)));
    return;
  }
  parseOdpPresentation(generateOdpPresentation(parseOdpPresentation(data)));
}

const files = CORPUS_DIRS.flatMap((dir) => walk(dir)).sort();
async function verifyAllFiles(): Promise<void> {
  const originalConsole = {
    log: console.log,
    info: console.info,
    warn: console.warn,
    error: console.error,
  };
  console.log = () => {};
  console.info = () => {};
  console.warn = () => {};
  console.error = () => {};
  try {
    for (const file of files) {
      const format = path.extname(file).slice(1).toLowerCase();
      if (!(format in counts)) continue;
      const result = counts[format as LegacyFormat];
      result.total++;
      try {
        await verify(format as LegacyFormat, new Uint8Array(fs.readFileSync(file)));
        result.pass++;
      } catch (error) {
        const outcome = classify(error);
        result[outcome]++;
        const reason = failureCategory(error);
        failureReasons[format as LegacyFormat].set(
          reason,
          (failureReasons[format as LegacyFormat].get(reason) ?? 0) + 1,
        );
      }
    }
  } finally {
    Object.assign(console, originalConsole);
  }
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
  if (/encrypted|password/i.test(reason)) return "encrypted";
  if (
    /(?:bad signature|truncated|must begin with|invalid CFB|invalid ODF|invalid zip|invalid URL|not a supported Compound|signature is not|unsupported RTF destination|objdata requires binary|Manifest does not declare|Manifest declares missing package path|embedded object subdocument is missing)/i.test(
      reason,
    )
  )
    return "invalid";
  return "unexpected";
}
