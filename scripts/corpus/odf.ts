/**
 * Read-only ODF corpus smoke harness.
 *
 * Recursively scans a caller-supplied directory and parses .odt, .ods, and
 * .odp files through the ODF package source. Results are aggregated only, so
 * no document names or paths are emitted. Encrypted inputs are identified
 * before parsing; package errors are separated from parser errors.
 */
import { readdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { unzipSync } from "fflate";

import { parsePresentation } from "../../packages/odp/src";
import { OdpParseError } from "../../packages/odp/src/error";
import { parseWorkbook } from "../../packages/ods/src";
import { OdsParseError } from "../../packages/ods/src/error";
import { parseDocument } from "../../packages/odt/src";
import { OdtParseError } from "../../packages/odt/src/error";

const decoder = new TextDecoder();

const FORMATS = {
  odp: { parse: parsePresentation },
  ods: { parse: parseWorkbook },
  odt: { parse: parseDocument },
} as const;

type Format = keyof typeof FORMATS;

interface Counts {
  total: number;
  success: number;
  encrypted: number;
  malformed: number;
  packageFailure: number;
  parseFailure: number;
}

interface ParseDiagnostic {
  file: string;
  format: Format;
  category: string;
  error: string;
  part?: string;
  name?: string;
  reason?: string;
}

function emptyCounts(): Counts {
  return {
    total: 0,
    success: 0,
    encrypted: 0,
    malformed: 0,
    packageFailure: 0,
    parseFailure: 0,
  };
}

const MALFORMED_PARSE_REASONS = new Set([
  "invalid-element-placement",
  "invalid-length",
  "invalid-repeat-count",
  "invalid-variable-value-type",
]);
const PACKAGE_PARSE_REASONS = new Set(["missing-package-part"]);

async function collectFiles(directory: string, files: string[] = []): Promise<string[]> {
  const entries = await readdir(directory, { withFileTypes: true });

  for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name))) {
    const filePath = path.join(directory, entry.name);
    if (entry.isDirectory()) await collectFiles(filePath, files);

    const extension = path.extname(entry.name).slice(1).toLowerCase() as Format;
    if (entry.isFile() && extension in FORMATS) files.push(filePath);
  }

  return files;
}

function isEncrypted(data: Uint8Array): boolean {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let eocd = -1;

  for (let offset = data.byteLength - 22; offset >= 0 && offset <= data.byteLength - 4; offset--) {
    if (view.getUint32(offset, true) === 0x0605_4b50) {
      eocd = offset;
      break;
    }
  }
  if (eocd < 0) return false;

  let offset = view.getUint32(eocd + 16, true);
  const count = view.getUint16(eocd + 10, true);
  for (let index = 0; index < count; index++) {
    if (offset + 46 > view.byteLength || view.getUint32(offset, true) !== 0x0201_4b50) {
      return false;
    }
    if ((view.getUint16(offset + 8, true) & 0x0001) !== 0) return true;
    offset +=
      46 +
      view.getUint16(offset + 28, true) +
      view.getUint16(offset + 30, true) +
      view.getUint16(offset + 32, true);
  }

  return false;
}

function hasEncryptedManifest(data: Uint8Array): boolean {
  try {
    const manifest = unzipSync(data)["META-INF/manifest.xml"];
    return manifest !== undefined && decoder.decode(manifest).includes("manifest:encryption-data");
  } catch {
    return false;
  }
}

function packageCategory(message: string): string {
  const normalized = message.toLowerCase();
  if (normalized.includes("zip") || normalized.includes("compression"))
    return "package:invalid-zip";
  if (normalized.includes("unexpected odf mime type")) return "package:unsupported-mime";
  if (normalized.includes("manifest")) return "package:invalid-manifest";
  if (normalized.includes("cannot read content.xml")) return "package:missing-content";
  if (normalized.includes("cannot read styles.xml")) return "package:missing-styles";
  if (normalized.includes("embedded object subdocument is missing")) {
    return "package:missing-embedded-object";
  }
  return "package:other";
}

function parseCategory(error: unknown): string {
  if (error instanceof OdtParseError) return `parse:${error.elementName}:${error.reason}`;
  if (error instanceof OdsParseError) return `parse:${error.elementName}:${error.reason}`;
  if (error instanceof OdpParseError) return `parse:${error.elementName}:${error.reason}`;
  return "parse:unknown";
}

function printLine(label: string, value: number | string): void {
  console.log(`${label}: ${value}`);
}

function printCounts(label: string, counts: Counts): void {
  printLine(label, counts.total);
  printLine(`${label}.success`, counts.success);
  printLine(`${label}.encrypted`, counts.encrypted);
  printLine(`${label}.malformed`, counts.malformed);
  printLine(`${label}.packageFailure`, counts.packageFailure);
  printLine(`${label}.parseFailure`, counts.parseFailure);
}

async function run(directory: string): Promise<void> {
  const files = await collectFiles(directory);
  const diagnosticPath = process.argv[4];
  const writeDiagnostics = process.argv[3] === "--diagnostics" && diagnosticPath !== undefined;
  const diagnostics: ParseDiagnostic[] = [];
  const formats = {
    odp: emptyCounts(),
    ods: emptyCounts(),
    odt: emptyCounts(),
  } as Record<Format, Counts>;
  const totals = emptyCounts();
  const errorCategories = new Map<string, number>();

  for (const file of files) {
    const extension = path.extname(file).slice(1).toLowerCase() as Format;
    const counts = formats[extension]!;
    counts.total++;
    totals.total++;

    let data: Uint8Array;
    try {
      data = new Uint8Array(await readFile(file));
    } catch {
      counts.packageFailure++;
      totals.packageFailure++;
      errorCategories.set("package:read", (errorCategories.get("package:read") ?? 0) + 1);
      continue;
    }

    if (isEncrypted(data) || hasEncryptedManifest(data)) {
      counts.encrypted++;
      totals.encrypted++;
      continue;
    }

    try {
      await FORMATS[extension].parse(data);
      counts.success++;
      totals.success++;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const structuredReason =
        error instanceof OdtParseError ||
        error instanceof OdsParseError ||
        error instanceof OdpParseError
          ? error.reason
          : undefined;
      const isMalformed =
        structuredReason !== undefined && MALFORMED_PARSE_REASONS.has(structuredReason);
      const isPackageReason =
        structuredReason !== undefined && PACKAGE_PARSE_REASONS.has(structuredReason);
      const isPackageFailure =
        isPackageReason ||
        /(?:unexpected odf mime type|odf package|invalid odf manifest|manifest does not declare|invalid zip|unknown compression|embedded object subdocument is missing)/i.test(
          message,
        ) ||
        ((error instanceof OdtParseError ||
          error instanceof OdsParseError ||
          error instanceof OdpParseError) &&
          error.part === "mimetype");
      const category = isPackageFailure ? packageCategory(message) : parseCategory(error);

      if (isMalformed) {
        counts.malformed++;
        totals.malformed++;
      } else if (isPackageFailure) {
        counts.packageFailure++;
        totals.packageFailure++;
      } else {
        counts.parseFailure++;
        totals.parseFailure++;
      }
      errorCategories.set(category, (errorCategories.get(category) ?? 0) + 1);
      if (writeDiagnostics) {
        diagnostics.push({
          file,
          format: extension,
          category,
          error: message,
          ...(error instanceof OdtParseError ||
          error instanceof OdsParseError ||
          error instanceof OdpParseError
            ? { part: error.part, name: error.elementName, reason: error.reason }
            : {}),
        });
      }
    }
  }

  printLine("files", totals.total);
  printCounts("total", totals);
  for (const format of Object.keys(formats) as Format[]) printCounts(format, formats[format]!);

  if (errorCategories.size > 0) {
    console.log("errorCategories:");
    for (const [category, count] of [...errorCategories].sort(([left], [right]) =>
      left.localeCompare(right),
    )) {
      printLine(`  ${category}`, count);
    }
  }

  if (writeDiagnostics) {
    await writeFile(diagnosticPath!, `${JSON.stringify(diagnostics, null, 2)}\n`);
  }

  if (totals.parseFailure > 0) process.exitCode = 1;
}

const directory = process.argv[2];
if (!directory || (process.argv.length !== 3 && process.argv.length !== 5)) {
  console.error("usage: tsx scripts/corpus/odf.ts <directory>");
  process.exitCode = 1;
} else {
  await run(path.resolve(directory)).catch(() => {
    console.error("error: directory is not readable");
    process.exitCode = 1;
  });
}
