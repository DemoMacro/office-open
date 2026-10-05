import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { inflateRawSync } from "node:zlib";

import { parseDocument as parseLegacyDocument } from "../../packages/doc/dist/index.mjs";
import { generateDocument, parseDocument as parseDocx } from "../../packages/docx/dist/index.mjs";
import { parsePresentation as parseLegacyPresentation } from "../../packages/ppt/dist/index.mjs";
import {
  generatePresentation,
  parsePresentation as parsePptx,
} from "../../packages/pptx/dist/index.mjs";
import { parseWorkbook as parseLegacyWorkbook } from "../../packages/xls/dist/index.mjs";
import { generateWorkbook, parseWorkbook as parseXlsx } from "../../packages/xlsx/dist/index.mjs";

type JsonRecord = Record<string, unknown>;
type Projection = string[];

interface MatrixEntry {
  name: string;
  convertedFixture: string;
  parseConverted: (data: Uint8Array) => Promise<JsonRecord>;
  generateSource: (options: JsonRecord) => Promise<Uint8Array>;
  parseSource: (data: Uint8Array) => Promise<JsonRecord>;
  projection: (options: JsonRecord) => Projection;
  expectedProjection: Projection;
  expectedLosses: readonly string[];
}

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

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

function normalizeText(value: string): string {
  return value.replace(/\s+/g, "");
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
  if (isRecord(value) && typeof value.text === "string") return value.text;
  return "";
}

function workbookProjection(options: JsonRecord): Projection {
  const worksheets = Array.isArray(options.worksheets) ? options.worksheets : [];
  return worksheets.flatMap((worksheetValue) => {
    if (!isRecord(worksheetValue) || !Array.isArray(worksheetValue.rows)) return [];
    return worksheetValue.rows.flatMap((rowValue) => {
      if (!isRecord(rowValue) || !Array.isArray(rowValue.cells)) return [];
      return rowValue.cells
        .filter((cell) => isRecord(cell) && typeof cell.reference === "string")
        .map((cell) => {
          if (!isRecord(cell) || typeof cell.reference !== "string") return "";
          return `${cell.reference}:${normalizeText(cellValue(cell.value))}`;
        })
        .filter(Boolean);
    });
  });
}

function fixture(fixtureName: string): Uint8Array {
  const compressed = fs.readFileSync(path.join(ROOT, "scripts/fixtures", fixtureName));
  return new Uint8Array(inflateRawSync(Buffer.from(compressed.toString("utf8").trim(), "base64")));
}

const MATRIX: readonly MatrixEntry[] = [
  {
    name: "docx↔doc",
    convertedFixture: "doc.b64z",
    parseConverted: parseLegacyDocument,
    generateSource: generateDocument,
    parseSource: parseDocx,
    projection: documentProjection,
    expectedProjection: ["Thequickbrownfoxjumpsoverthelazydog"],
    expectedLosses: [],
  },
  {
    name: "pptx↔ppt",
    convertedFixture: "ppt.b64z",
    parseConverted: parseLegacyPresentation,
    generateSource: generatePresentation,
    parseSource: parsePptx,
    projection: presentationProjection,
    expectedProjection: ["First", "Second"],
    expectedLosses: [],
  },
  {
    name: "xlsx↔xls",
    convertedFixture: "xls.b64z",
    parseConverted: parseLegacyWorkbook,
    generateSource: generateWorkbook,
    parseSource: parseXlsx,
    projection: workbookProjection,
    expectedProjection: ["A1:1.5", "B2:Legacylabel"],
    expectedLosses: [],
  },
];

function equalProjection(left: Projection, right: Projection): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

async function verify(entry: MatrixEntry): Promise<string[]> {
  const failures: string[] = [];
  const converted = fixture(entry.convertedFixture);
  if (converted.byteLength === 0) failures.push("converted sample is empty");

  let convertedOptions: JsonRecord;
  try {
    convertedOptions = await entry.parseConverted(converted);
  } catch (error) {
    return [
      `converted sample parse failed: ${error instanceof Error ? error.message : String(error)}`,
    ];
  }

  const convertedProjection = entry.projection(convertedOptions);
  if (convertedProjection.length === 0) failures.push("converted semantic projection is empty");
  if (!equalProjection(convertedProjection, entry.expectedProjection)) {
    failures.push(
      `converted projection mismatch: expected ${entry.expectedProjection.length} values, received ${convertedProjection.length}`,
    );
  }

  let source: Uint8Array;
  try {
    source = await entry.generateSource(convertedOptions);
  } catch (error) {
    return [
      ...failures,
      `source generation failed: ${error instanceof Error ? error.message : String(error)}`,
    ];
  }
  if (source.byteLength === 0) failures.push("generated source sample is empty");

  try {
    const sourceOptions = await entry.parseSource(source);
    const sourceProjection = entry.projection(sourceOptions);
    if (sourceProjection.length === 0) failures.push("source semantic projection is empty");
    if (!equalProjection(convertedProjection, sourceProjection)) {
      failures.push(
        `unexplained semantic loss: converted ${convertedProjection.length} values, source ${sourceProjection.length}`,
      );
    }
  } catch (error) {
    failures.push(
      `generated source parse failed: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
  return failures;
}

const rows: Array<{
  name: string;
  samples: number;
  values: number;
  losses: number;
  pass: boolean;
}> = [];
for (const entry of MATRIX) {
  const failures = await verify(entry);
  const values = entry.expectedProjection.length;
  rows.push({
    name: entry.name,
    samples: 1,
    values,
    losses: entry.expectedLosses.length,
    pass: failures.length === 0,
  });
  console.log(
    `${entry.name}: ${failures.length === 0 ? "PASS" : "FAIL"} sample=1 projectionValues=${values} expectedLoss=${entry.expectedLosses.length}`,
  );
  for (const failure of failures) console.log(`  FAIL ${failure}`);
}

const samples = rows.reduce((total, row) => total + row.samples, 0);
const values = rows.reduce((total, row) => total + row.values, 0);
const losses = rows.reduce((total, row) => total + row.losses, 0);
const passed = rows.filter((row) => row.pass).length;
console.log(
  `matrix: ${passed}/${MATRIX.length} pairs; samples=${samples}; projectionValues=${values}; expectedLosses=${losses}`,
);
if (passed !== MATRIX.length || samples !== MATRIX.length) {
  console.error("cross-format gate: FAILED");
  process.exitCode = 1;
} else {
  console.log("cross-format gate: OK");
}
