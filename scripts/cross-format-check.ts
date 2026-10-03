/**
 * Aggregate audit gate for OOXML files converted to legacy binary formats.
 *
 * Requires a prior `pnpm build` because package dist files are exercised.
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { parseDocument as parseLegacyDocument } from "../packages/doc/dist/index.mjs";
import { generateDocument, parseDocument as parseDocx } from "../packages/docx/dist/index.mjs";
import { parsePresentation as parseLegacyPresentation } from "../packages/ppt/dist/index.mjs";
import {
  generatePresentation,
  parsePresentation as parsePptx,
} from "../packages/pptx/dist/index.mjs";
import { parseWorkbook as parseLegacyWorkbook } from "../packages/xls/dist/index.mjs";
import { generateWorkbook, parseWorkbook as parseXlsx } from "../packages/xlsx/dist/index.mjs";

type JsonRecord = Record<string, unknown>;
type Projection = string[];
type Parse = (data: Uint8Array) => Promise<JsonRecord>;
type Generate = (options: JsonRecord, output?: { type: "uint8array" }) => Promise<Uint8Array>;

interface FormatConfig {
  name: string;
  sourceDirectory: string;
  sourceExtension: string;
  convertedDirectory: string;
  convertedExtension: string;
  parseSource: Parse;
  parseLegacy: Parse;
  generate: Generate;
  parseGenerated: Parse;
  projection: (options: JsonRecord, ignoredCellReferences?: Set<string>) => Projection;
}

type Outcome =
  | "total"
  | "converted-missing"
  | "source-parse-failed"
  | "legacy-parse-failed"
  | "legacy-encrypted"
  | "office-conversion-loss"
  | "projection-order-difference"
  | "generate-failed"
  | "generated-parse-failed"
  | "generated-projection-mismatch"
  | "pass";

const FORMATS: FormatConfig[] = [
  {
    name: "docx",
    sourceDirectory: "packages/docx/.temp",
    sourceExtension: ".docx",
    convertedDirectory: ".temp/cross-format-check/word",
    convertedExtension: ".doc",
    parseSource: parseDocx,
    parseLegacy: parseLegacyDocument,
    generate: generateDocument,
    parseGenerated: parseDocx,
    projection: documentProjection,
  },
  {
    name: "pptx",
    sourceDirectory: "packages/pptx/.temp",
    sourceExtension: ".pptx",
    convertedDirectory: ".temp/cross-format-check/powerpoint",
    convertedExtension: ".ppt",
    parseSource: parsePptx,
    parseLegacy: parseLegacyPresentation,
    generate: generatePresentation,
    parseGenerated: parsePptx,
    projection: presentationProjection,
  },
  {
    name: "xlsx",
    sourceDirectory: "packages/xlsx/.temp",
    sourceExtension: ".xlsx",
    convertedDirectory: ".temp/cross-format-check/excel",
    convertedExtension: ".xls",
    parseSource: parseXlsx,
    parseLegacy: parseLegacyWorkbook,
    generate: generateWorkbook,
    parseGenerated: parseXlsx,
    projection: workbookProjection,
  },
];

const OUTCOMES: Outcome[] = [
  "total",
  "converted-missing",
  "source-parse-failed",
  "legacy-parse-failed",
  "legacy-encrypted",
  "office-conversion-loss",
  "projection-order-difference",
  "generate-failed",
  "generated-parse-failed",
  "generated-projection-mismatch",
  "pass",
];

const SCRIPT_DIRECTORY = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIRECTORY = path.resolve(SCRIPT_DIRECTORY, "..");
const counts = Object.fromEntries(
  FORMATS.map((format) => [
    format.name,
    Object.fromEntries(OUTCOMES.map((outcome) => [outcome, 0])),
  ]),
) as Record<string, Record<Outcome, number>>;

function isRecord(value: unknown): value is JsonRecord {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function runText(value: unknown, output: string[] = []): string[] {
  if (typeof value === "string") output.push(value);
  else if (Array.isArray(value)) {
    for (const item of value) runText(item, output);
  } else if (isRecord(value)) {
    if (typeof value.text === "string") output.push(value.text);
    if (value.tab === true) output.push(" ");
    if (Array.isArray(value.children)) runText(value.children, output);
    if (Array.isArray(value.runs)) runText(value.runs, output);
  }
  return output;
}

function normalizeText(value: string): string {
  return value.replace(/\s+/g, "");
}

function collectText(value: unknown, output: string[] = []): string[] {
  if (Array.isArray(value)) {
    for (const item of value) collectText(item, output);
  } else if (isRecord(value)) {
    for (const [key, child] of Object.entries(value)) {
      if (key === "paragraph") output.push(runText(child).join(""));
      else if (key === "paragraphs" && Array.isArray(child)) {
        for (const paragraph of child) output.push(runText(paragraph).join(""));
      } else if (key === "text" && typeof child === "string") output.push(child);
      else collectText(child, output);
    }
  }
  return output;
}

function documentProjection(options: JsonRecord): Projection {
  const sections = Array.isArray(options.sections) ? options.sections : [];
  return sections
    .filter(isRecord)
    .flatMap((section) => collectText(section.children))
    .map(normalizeText)
    .filter(Boolean);
}

function presentationProjection(options: JsonRecord): Projection {
  return collectText(options.slides).map(normalizeText).filter(Boolean);
}

function cellValue(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value instanceof Date) return value.toISOString();
  if (isRecord(value)) {
    if (typeof value.text === "string") return value.text;
    if (Array.isArray(value.runs)) return runText(value.runs).join("");
  }
  return JSON.stringify(value);
}

function formulaCellReferences(options: JsonRecord): Set<string> {
  const references = new Set<string>();
  const worksheets = Array.isArray(options.worksheets) ? options.worksheets : [];
  for (const worksheetValue of worksheets) {
    if (!isRecord(worksheetValue) || !Array.isArray(worksheetValue.rows)) continue;
    for (const rowValue of worksheetValue.rows) {
      if (!isRecord(rowValue) || !Array.isArray(rowValue.cells)) continue;
      for (const cell of rowValue.cells) {
        if (isRecord(cell) && cell.formula !== undefined && typeof cell.reference === "string") {
          references.add(cell.reference);
        }
      }
    }
  }
  return references;
}

function workbookProjection(
  options: JsonRecord,
  ignoredCellReferences = new Set<string>(),
): Projection {
  const worksheets = Array.isArray(options.worksheets) ? options.worksheets : [];
  return worksheets.flatMap((worksheetValue) => {
    if (!isRecord(worksheetValue)) return [];
    const rows = Array.isArray(worksheetValue.rows) ? worksheetValue.rows : [];
    return rows.flatMap((rowValue) => {
      if (!isRecord(rowValue) || !Array.isArray(rowValue.cells)) return [];
      return rowValue.cells
        .filter((candidate) => {
          if (!isRecord(candidate)) return false;
          const reference = candidate.reference;
          return (
            typeof reference === "string" &&
            !ignoredCellReferences.has(reference) &&
            cellValue(candidate.value) !== ""
          );
        })
        .map((candidate) => {
          if (!isRecord(candidate) || typeof candidate.reference !== "string") return "";
          return `${candidate.reference}:${normalizeText(cellValue(candidate.value))}`;
        });
    });
  });
}

function countValues(values: readonly string[]): Map<string, number> {
  const result = new Map<string, number>();
  for (const value of values) result.set(value, (result.get(value) ?? 0) + 1);
  return result;
}

function missingValues(source: readonly string[], legacy: readonly string[]): string[] {
  const legacyCounts = countValues(legacy);
  const result: string[] = [];
  for (const value of source) {
    const count = legacyCounts.get(value) ?? 0;
    if (count === 0) result.push(value);
    else legacyCounts.set(value, count - 1);
  }
  return result;
}

function classifyProjectionMismatch(
  source: Projection,
  legacy: Projection,
): Exclude<Outcome, "total" | "pass"> {
  const missing = missingValues(source, legacy);
  if (missing.length === 0) return "projection-order-difference";
  const legacySequence = legacy.join("");
  return missing.every((value) => legacySequence.includes(value))
    ? "projection-order-difference"
    : "office-conversion-loss";
}

function isEncryptedError(error: unknown): boolean {
  return /encrypted/i.test(String((error as Error)?.message ?? error));
}

function record(format: FormatConfig, outcome: Exclude<Outcome, "total">): void {
  counts[format.name]![outcome]! += 1;
  counts[format.name]!.total += 1;
}

function equalProjection(left: Projection, right: Projection): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function readBytes(file: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(file));
}

function sourceFiles(format: FormatConfig): string[] {
  const directory = path.resolve(ROOT_DIRECTORY, format.sourceDirectory);
  return fs
    .readdirSync(directory)
    .filter((name) => name.endsWith(format.sourceExtension) && !name.startsWith("~$"))
    .sort()
    .map((name) => path.join(directory, name));
}

function isUnexpectedOutcome(outcome: Exclude<Outcome, "total" | "pass">): boolean {
  return !(
    outcome === "converted-missing" ||
    outcome === "office-conversion-loss" ||
    outcome === "projection-order-difference" ||
    outcome === "legacy-encrypted"
  );
}

async function verifyFile(format: FormatConfig, sourcePath: string): Promise<void> {
  const convertedPath = path.join(
    path.resolve(ROOT_DIRECTORY, format.convertedDirectory),
    path.basename(sourcePath, format.sourceExtension) + format.convertedExtension,
  );
  if (!fs.existsSync(convertedPath)) {
    record(format, "converted-missing");
    return;
  }

  let sourceOptions: JsonRecord;
  try {
    sourceOptions = await format.parseSource(readBytes(sourcePath));
  } catch {
    record(format, "source-parse-failed");
    return;
  }

  let legacyOptions: JsonRecord;
  const legacyBytes = readBytes(convertedPath);
  try {
    legacyOptions = await format.parseLegacy(legacyBytes);
  } catch (error) {
    record(format, isEncryptedError(error) ? "legacy-encrypted" : "legacy-parse-failed");
    return;
  }

  const ignoredCellReferences =
    format.name === "xlsx" ? formulaCellReferences(sourceOptions) : undefined;
  const sourceProjection = format.projection(sourceOptions, ignoredCellReferences);
  const legacyProjection = format.projection(legacyOptions, ignoredCellReferences);
  if (!equalProjection(sourceProjection, legacyProjection)) {
    record(format, classifyProjectionMismatch(sourceProjection, legacyProjection));
    return;
  }

  let generated: Uint8Array;
  try {
    generated = await format.generate(legacyOptions, { type: "uint8array" });
  } catch {
    record(format, "generate-failed");
    return;
  }

  let generatedOptions: JsonRecord;
  try {
    generatedOptions = await format.parseGenerated(generated);
  } catch {
    record(format, "generated-parse-failed");
    return;
  }
  if (
    !equalProjection(legacyProjection, format.projection(generatedOptions, ignoredCellReferences))
  ) {
    record(format, "generated-projection-mismatch");
    return;
  }
  record(format, "pass");
}

for (const format of FORMATS) {
  if (!fs.existsSync(path.resolve(ROOT_DIRECTORY, format.convertedDirectory))) {
    console.error(`cross-format gate: ${format.name} converted corpus is missing`);
    process.exit(1);
  }
  for (const source of sourceFiles(format)) await verifyFile(format, source);
}

const columns = ["format", ...OUTCOMES];
console.log(columns.map((column) => column.padEnd(27)).join(""));
let failed = false;
for (const format of FORMATS) {
  const values = OUTCOMES.map((outcome) => String(counts[format.name]![outcome]!));
  console.log(
    [format.name, ...values].map((value, index) => value.padEnd(columns[index]!.length)).join(""),
  );
  for (const outcome of OUTCOMES) {
    if (
      outcome !== "total" &&
      outcome !== "pass" &&
      isUnexpectedOutcome(outcome) &&
      counts[format.name]![outcome]! > 0
    ) {
      failed = true;
    }
  }
}
if (failed) {
  console.error("cross-format gate: FAILED");
  process.exit(1);
}
