/**
 * Read-only ODF corpus smoke harness.
 *
 * Recursively scans a caller-supplied directory and parses .odt, .ods, and
 * .odp files through the ODF package source. Results are aggregated only, so
 * no document names or paths are emitted. Encrypted inputs are identified
 * before parsing; package errors are separated from parser errors.
 */
import { readdir, readFile } from "node:fs/promises";
import path from "node:path";

import { unzipSync } from "fflate";

import { parseOdp, parseOds, parseOdt } from "../packages/odf/src";

const decoder = new TextDecoder();

const FORMATS = {
  odp: { parse: parseOdp },
  ods: { parse: parseOds },
  odt: { parse: parseOdt },
} as const;

type Format = keyof typeof FORMATS;

interface Counts {
  total: number;
  success: number;
  encrypted: number;
  packageFailure: number;
  parseFailure: number;
}

function emptyCounts(): Counts {
  return {
    total: 0,
    success: 0,
    encrypted: 0,
    packageFailure: 0,
    parseFailure: 0,
  };
}

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
  if (normalized.includes("zip") || normalized.includes("compression")) return "package:zip";
  if (normalized.includes("mime")) return "package:mime";
  if (normalized.includes("manifest")) return "package:manifest";
  if (normalized.includes("missing content.xml")) return "package:missing-content";
  return "package:other";
}

function printLine(label: string, value: number | string): void {
  console.log(`${label}: ${value}`);
}

function printCounts(label: string, counts: Counts): void {
  printLine(label, counts.total);
  printLine(`${label}.success`, counts.success);
  printLine(`${label}.encrypted`, counts.encrypted);
  printLine(`${label}.packageFailure`, counts.packageFailure);
  printLine(`${label}.parseFailure`, counts.parseFailure);
}

async function run(directory: string): Promise<void> {
  const files = await collectFiles(directory);
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
      const isPackageFailure =
        /^(unexpected odf mime type|odf package|invalid odf manifest|manifest does not declare|invalid zip|unknown compression)/i.test(
          message,
        );
      const category = isPackageFailure ? packageCategory(message) : "parse";

      if (isPackageFailure) {
        counts.packageFailure++;
        totals.packageFailure++;
      } else {
        counts.parseFailure++;
        totals.parseFailure++;
      }
      errorCategories.set(category, (errorCategories.get(category) ?? 0) + 1);
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

  if (totals.packageFailure > 0 || totals.parseFailure > 0) process.exitCode = 1;
}

const directory = process.argv[2];
if (!directory || process.argv.length !== 3) {
  console.error("usage: tsx scripts/odf-corpus.ts <directory>");
  process.exitCode = 1;
} else {
  await run(path.resolve(directory)).catch(() => {
    console.error("error: directory is not readable");
    process.exitCode = 1;
  });
}
