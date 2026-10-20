import { createHash } from "node:crypto";
import * as fs from "node:fs";
import * as path from "node:path";

import { CfbError } from "../../packages/core/dist/cfb/index.mjs";
import { OOXML_PACKAGE_FORMATS, OpcPackageError } from "../../packages/core/dist/opc/index.mjs";
import { parseDocument as parseDocDocument } from "../../packages/doc/dist/index.mjs";
import { DocParseError } from "../../packages/doc/dist/index.mjs";
import {
  generateDocument as generateDocxDocument,
  parseDocument as parseDocxDocument,
} from "../../packages/docx/dist/index.mjs";
import { DocxParseError } from "../../packages/docx/dist/index.mjs";
import {
  OcfError,
  OcfManifestError,
  OcfMimeTypeError,
  OdfSchemaError,
  OdfXmlError,
} from "../../packages/odf/dist/index.mjs";
import {
  generatePresentation as generateOdpPresentation,
  parsePresentation as parseOdpPresentation,
} from "../../packages/odp/dist/index.mjs";
import { OdpGenerateError, OdpParseError } from "../../packages/odp/dist/index.mjs";
import {
  generateWorkbook as generateOdsWorkbook,
  parseWorkbook as parseOdsWorkbook,
} from "../../packages/ods/dist/index.mjs";
import { OdsParseError } from "../../packages/ods/dist/index.mjs";
import {
  generateDocument as generateOdtDocument,
  parseDocument as parseOdtDocument,
} from "../../packages/odt/dist/index.mjs";
import { OdtGenerateError, OdtParseError } from "../../packages/odt/dist/index.mjs";
import { parsePresentation as parsePptDocument } from "../../packages/ppt/dist/index.mjs";
import { PptParseError } from "../../packages/ppt/dist/index.mjs";
import {
  generatePresentation as generatePptxPresentation,
  parsePresentation as parsePptxPresentation,
} from "../../packages/pptx/dist/index.mjs";
import { PptxParseError } from "../../packages/pptx/dist/index.mjs";
import {
  generateDocument as generateRtfDocument,
  parseDocument as parseRtfDocument,
} from "../../packages/rtf/dist/index.mjs";
import { RtfGenerateError, RtfParseError } from "../../packages/rtf/dist/index.mjs";
import { parseWorkbook as parseXlsDocument } from "../../packages/xls/dist/index.mjs";
import { XlsParseError } from "../../packages/xls/dist/index.mjs";
import {
  generateWorkbook as generateXlsxWorkbook,
  parseWorkbook as parseXlsxDocument,
} from "../../packages/xlsx/dist/index.mjs";
import { XlsxParseError } from "../../packages/xlsx/dist/index.mjs";

type DocumentSource = "doc" | "docx" | "docm" | "dotx" | "dotm" | "rtf" | "odt";
type WorkbookSource = "xls" | "xlsx" | "xlsm" | "xltx" | "xltm" | "ods";
type PresentationSource = "ppt" | "pptx" | "pptm" | "potx" | "potm" | "odp";
type SourceFormat = DocumentSource | WorkbookSource | PresentationSource;
type Family = "document" | "workbook" | "presentation";
type Outcome = "generated" | "encrypted" | "invalid" | "unsupported-target" | "failed";

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
  issues: CrossConvertIssue[];
}

interface CrossConvertIssue {
  readonly source: string;
  readonly target?: string;
  readonly phase: "parse" | "generate" | "write" | "reparse" | "worker";
  readonly reason: string;
}

interface WorkerResult {
  readonly document: FamilyResult;
  readonly workbook: FamilyResult;
  readonly presentation: FamilyResult;
}

interface WorkerInput {
  readonly items: WorkItem[];
  readonly outputRoot?: string;
}

const FAMILY_TARGETS: Record<Family, readonly string[]> = {
  document: ["docx", "docm", "dotx", "dotm", "rtf", "odt"],
  workbook: ["xlsx", "xlsm", "xltx", "xltm", "ods"],
  presentation: ["pptx", "pptm", "potx", "potm", "odp"],
} as const satisfies Record<Family, readonly string[]>;

const LEGACY_PASSWORDS = ["VelvetSweatshop", "Password1234_", "password"] as const;
const MAX_STRING_BYTES = 0x1fffffe8;

function decodeRtf(data: Uint8Array): string {
  if (data[0] === 0xff && data[1] === 0xfe) return new TextDecoder("utf-16le").decode(data);
  if (data[0] === 0xfe && data[1] === 0xff) return new TextDecoder("utf-16be").decode(data);
  if (data[0] === 0xef && data[1] === 0xbb && data[2] === 0xbf)
    return new TextDecoder("utf-8").decode(data);
  return Buffer.from(data).toString("latin1");
}

const emptyFamily = (): FamilyResult => ({
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
  issues: [],
});

const emptyResult = (): WorkerResult => ({
  document: emptyFamily(),
  workbook: emptyFamily(),
  presentation: emptyFamily(),
});

function familyOf(source: SourceFormat): Family {
  if (
    source === "xls" ||
    source === "xlsx" ||
    source === "xlsm" ||
    source === "xltx" ||
    source === "xltm" ||
    source === "ods"
  )
    return "workbook";
  if (
    source === "ppt" ||
    source === "pptx" ||
    source === "pptm" ||
    source === "potx" ||
    source === "potm" ||
    source === "odp"
  )
    return "presentation";
  return "document";
}

function outcomeOf(error: unknown): Exclude<Outcome, "generated"> {
  const name = (error as { name?: string })?.name;
  const reason =
    (error as { context?: { reason?: string }; reason?: string })?.context?.reason ??
    (error as { reason?: string })?.reason;
  if (
    name === RtfGenerateError.name ||
    name === OdpGenerateError.name ||
    name === OdtGenerateError.name
  ) {
    return "unsupported-target";
  }
  if (
    name === DocParseError.name ||
    name === XlsParseError.name ||
    name === PptParseError.name ||
    name === DocxParseError.name ||
    name === PptxParseError.name ||
    name === XlsxParseError.name ||
    name === RtfParseError.name ||
    name === OcfError.name ||
    name === OcfMimeTypeError.name ||
    name === OcfManifestError.name ||
    name === OdfXmlError.name ||
    name === OdfSchemaError.name ||
    name === OdtParseError.name ||
    name === OdsParseError.name ||
    name === OdpParseError.name ||
    name === CfbError.name ||
    name === OpcPackageError.name
  ) {
    return "invalid";
  }
  return "failed";
}

function structuredReason(error: unknown): string | undefined {
  return (
    (error as { context?: { reason?: string }; reason?: string })?.context?.reason ??
    (error as { reason?: string })?.reason
  );
}

function isEncrypted(error: unknown): boolean {
  return structuredReason(error) === "encrypted-unsupported";
}

function reasonOf(error: unknown): string {
  const message = String((error as Error)?.message ?? error);
  return message.replace(/"(.{32})[^"]*"/g, '"$1…"').slice(0, 220);
}

function issueSource(id: string): string {
  return createHash("sha256").update(id).digest("hex").slice(0, 12);
}

function count(family: FamilyResult, outcome: Outcome, error?: unknown): void {
  if (outcome === "generated") {
    family.generated++;
    return;
  }
  if (outcome === "encrypted") family.encrypted++;
  else if (outcome === "invalid") family.invalid++;
  else if (outcome === "unsupported-target") family.unsupportedTarget++;
  else family.unexpectedSources++;
  if (error) {
    const reason = `${outcome}: ${reasonOf(error)}`;
    family.reasons[reason] = (family.reasons[reason] ?? 0) + 1;
  }
}

function addIssue(
  family: FamilyResult,
  issue: Omit<CrossConvertIssue, "source"> & { sourceId: string },
): void {
  family.issues.push({ ...issue, source: issueSource(issue.sourceId) });
}

async function withLegacyPasswords<T>(parse: (password?: string) => Promise<T> | T): Promise<T> {
  try {
    return await parse();
  } catch (error) {
    if (!isEncrypted(error)) throw error;
  }
  let lastError: unknown;
  for (const password of LEGACY_PASSWORDS) {
    try {
      return await parse(password);
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError ?? new Error("Encrypted legacy document could not be decrypted");
}

async function parseSource(source: SourceFormat, data: Uint8Array): Promise<unknown> {
  if (source === "doc")
    return withLegacyPasswords((password) =>
      parseDocDocument(data, password ? { password } : undefined),
    );
  if (["docx", "docm", "dotx", "dotm"].includes(source)) return parseDocxDocument(data);
  if (source === "rtf") return parseRtfDocument(decodeRtf(data));
  if (source === "odt") return parseOdtDocument(data);
  if (source === "xls")
    return withLegacyPasswords((password) =>
      parseXlsDocument(data, password ? { password } : undefined),
    );
  if (["xlsx", "xlsm", "xltx", "xltm"].includes(source)) return parseXlsxDocument(data);
  if (source === "ods") return parseOdsWorkbook(data);
  if (source === "ppt")
    return withLegacyPasswords((password) =>
      parsePptDocument(data, password ? { password } : undefined),
    );
  if (["pptx", "pptm", "potx", "potm"].includes(source)) return parsePptxPresentation(data);
  return parseOdpPresentation(data);
}

function generateTarget(family: Family, target: string, options: unknown): Promise<Uint8Array> {
  return Promise.resolve(generateBytes(family, target, options));
}

function generateBytes(
  family: Family,
  target: string,
  options: unknown,
): Uint8Array | Promise<Uint8Array> {
  const ooxmlTarget = target as keyof typeof OOXML_PACKAGE_FORMATS;
  if (["docx", "docm", "dotx", "dotm"].includes(target))
    return generateDocxDocument(requestedOptions(ooxmlTarget, options) as never, {
      packageVariant: OOXML_PACKAGE_FORMATS[ooxmlTarget].variant,
    });
  if (target === "rtf") return generateRtfDocument(options as never);
  if (target === "odt") return generateOdtDocument(options as never);
  if (["xlsx", "xlsm", "xltx", "xltm"].includes(target))
    return generateXlsxWorkbook(requestedOptions(ooxmlTarget, options) as never, {
      packageVariant: OOXML_PACKAGE_FORMATS[ooxmlTarget].variant,
    });
  if (target === "ods") return generateOdsWorkbook(options as never);
  if (["pptx", "pptm", "potx", "potm"].includes(target))
    return generatePptxPresentation(requestedOptions(ooxmlTarget, options) as never, {
      packageVariant: OOXML_PACKAGE_FORMATS[ooxmlTarget].variant,
    });
  return generateOdpPresentation(options as never);
}

async function reparse(family: Family, target: string, data: Uint8Array): Promise<void> {
  if (family === "workbook") {
    if (target === "ods") await parseOdsWorkbook(data);
    else await parseXlsxDocument(data);
  } else if (family === "presentation") {
    if (target === "odp") await parseOdpPresentation(data);
    else await parsePptxPresentation(data);
  } else if (target === "odt") await parseOdtDocument(data);
  else if (target === "rtf") await parseRtfDocument(decodeRtf(data));
  else await parseDocxDocument(data);
  if (!FAMILY_TARGETS[family].includes(target)) throw new Error(`Unexpected target ${target}`);
}

function requestedOptions(target: keyof typeof OOXML_PACKAGE_FORMATS, options: unknown): unknown {
  const info = OOXML_PACKAGE_FORMATS[target];
  const container = options as {
    contentTypes?: { overrides?: { partName?: string; contentType?: string }[] };
  };
  const overrides = container.contentTypes?.overrides;
  const partName = `/${info.mainPartPath}`.toLowerCase();
  const declared = overrides?.find(
    (override) => override.partName?.toLowerCase() === partName,
  )?.contentType;
  if (declared === undefined || declared === info.mainContentType) return options;
  const isFamilyMain = Object.values(OOXML_PACKAGE_FORMATS).some(
    (candidate) => candidate.family === info.family && candidate.mainContentType === declared,
  );
  if (!isFamilyMain) return options;
  const source = options as Record<string, unknown>;
  return {
    ...source,
    contentTypes: {
      ...container.contentTypes,
      overrides: overrides?.map((override) =>
        override.partName?.toLowerCase() === partName
          ? { ...override, contentType: info.mainContentType }
          : override,
      ),
    },
  };
}

function outputPath(
  outputRoot: string,
  family: Family,
  source: SourceFormat,
  target: string,
  id: string,
): string {
  const safeName = id.replaceAll(/[\\/:*?"<>|]/g, "_");
  return path.join(outputRoot, family, `${source}-to-${target}`, `${safeName}.${target}`);
}

const originalConsole = {
  log: console.log,
  info: console.info,
  warn: console.warn,
  error: console.error,
};

const silence = (): void => {
  console.log = () => {};
  console.info = () => {};
  console.warn = () => {};
  console.error = () => {};
};

const restore = (): void => {
  console.log = originalConsole.log;
  console.info = originalConsole.info;
  console.warn = originalConsole.warn;
  console.error = originalConsole.error;
};

let input = "";
for await (const chunk of process.stdin) input += chunk;
const { items, outputRoot } = JSON.parse(input) as WorkerInput;
const result = emptyResult();

for (const item of items) {
  const familyName = familyOf(item.source);
  const family = result[familyName];
  family.sources++;
  let options: unknown;
  try {
    silence();
    options = await parseSource(item.source, new Uint8Array(fs.readFileSync(item.file)));
  } catch (error) {
    restore();
    if (process.env.CROSS_CONVERT_DEBUG === "1")
      process.stderr.write(`parse failed: ${item.id}: ${String(error)}\n`);
    count(family, outcomeOf(error), error);
    addIssue(family, { sourceId: item.id, phase: "parse", reason: reasonOf(error) });
    continue;
  } finally {
    restore();
  }

  family.parsePass++;
  for (const target of FAMILY_TARGETS[familyName]) {
    if (target === item.source) continue;
    family.conversions++;
    let bytes: Uint8Array;
    try {
      silence();
      bytes = new Uint8Array(await generateTarget(familyName, target, options));
    } catch (error) {
      restore();
      if (process.env.CROSS_CONVERT_DEBUG === "1")
        process.stderr.write(`generate failed: ${item.id} -> ${target}: ${String(error)}\n`);
      const outcome = outcomeOf(error);
      count(family, outcome, error);
      addIssue(family, {
        sourceId: item.id,
        target,
        phase: "generate",
        reason: reasonOf(error),
      });
      if (outcome === "failed") family.generateFailures++;
      continue;
    } finally {
      restore();
    }

    if (outputRoot !== undefined) {
      try {
        const output = outputPath(outputRoot, familyName, item.source, target, item.id);
        fs.mkdirSync(path.dirname(output), { recursive: true });
        fs.writeFileSync(output, bytes);
      } catch {
        family.writeFailures++;
        addIssue(family, {
          sourceId: item.id,
          target,
          phase: "write",
          reason: "failed to write generated output",
        });
        continue;
      }
    }

    if (bytes.byteLength > MAX_STRING_BYTES) {
      const error = new Error(`Generated output exceeds V8 string limit: ${bytes.byteLength}`);
      count(family, "unsupported-target", error);
      addIssue(family, {
        sourceId: item.id,
        target,
        phase: "generate",
        reason: `generated output exceeds V8 string limit: ${bytes.byteLength} bytes`,
      });
      continue;
    }

    try {
      silence();
      await reparse(familyName, target, bytes);
    } catch (error) {
      restore();
      if (process.env.CROSS_CONVERT_DEBUG === "1")
        process.stderr.write(`reparse failed: ${item.id} -> ${target}: ${String(error)}\n`);
      family.reparseFailures++;
      count(family, outcomeOf(error), error);
      addIssue(family, {
        sourceId: item.id,
        target,
        phase: "reparse",
        reason: reasonOf(error),
      });
      continue;
    } finally {
      restore();
    }

    family.reparsed++;
    count(family, "generated");
  }
  options = undefined;
}

process.stdout.write(JSON.stringify(result));
