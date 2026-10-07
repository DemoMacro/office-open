import * as fs from "node:fs";
import * as path from "node:path";

import { OOXML_PACKAGE_FORMATS, opaquePassthroughPolicy } from "../../packages/core/dist/index.mjs";
import { parseDocument } from "../../packages/docx/dist/index.mjs";
import { parsePresentation } from "../../packages/pptx/dist/index.mjs";
import { parseWorkbook } from "../../packages/xlsx/dist/index.mjs";
import {
  assertEncryptedContainerRoundTrip,
  archiveSemanticDiffDetails,
  archiveTagDiffs,
  classifyPackageFailure,
  type CorpusFailureKind,
  orphanedPackageMembers,
  type SemanticPartDiff,
} from "./semantics";

const OS_JUNK_PART_PATTERNS: readonly RegExp[] = [
  /^(?:.*\/)?\.DS_Store$/i,
  /^(?:.*\/)?Thumbs\.db$/i,
  /^(?:.*\/)?desktop\.ini$/i,
];

export const EXTERNAL_OPAQUE_PARTS: Record<Format, readonly RegExp[]> = {
  docx: [...opaquePassthroughPolicy("docx").opaquePatterns, ...OS_JUNK_PART_PATTERNS],
  xlsx: [...opaquePassthroughPolicy("xlsx").opaquePatterns, ...OS_JUNK_PART_PATTERNS],
  pptx: [...opaquePassthroughPolicy("pptx").opaquePatterns, ...OS_JUNK_PART_PATTERNS],
};

export type Format = "docx" | "xlsx" | "pptx";
export type PackageFormat = keyof typeof OOXML_PACKAGE_FORMATS;

const BY_EXT: Record<
  string,
  { format: Format; type: PackageFormat; parse: (bytes: Uint8Array) => Promise<unknown> }
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

export interface FormatCounts {
  total: number;
  clean: number;
  diff: number;
  parseFail: number;
  genFail: number;
}

export interface LibraryRunInput {
  destination: string;
  root: string;
  strictSemantic: boolean;
}

export interface FileDiagnostic {
  file: string;
  format: Format;
  outcome: CorpusFailureKind;
  package?: string;
  diffCategories?: Record<string, number>;
  sampleParts?: SemanticPartDiff[];
  error?: string;
}

export interface LibraryRunResult {
  counts: Record<Format, FormatCounts>;
  blockers: Record<Format, [string, number][]>;
  diagnostics: FileDiagnostic[];
  rawAudit: Record<Format, { files: number; xmlParts: number; binaryParts: number }>;
  rawBlockers: Record<Format, [string, number][]>;
}

/**
 * Encrypted OOXML containers are opaque CFB payloads. Compare them bytewise;
 * unwrapping one as a ZIP is a comparator bug, not a parse failure.
 */
export function isEncryptedPassthrough(options: unknown): boolean {
  return (
    typeof options === "object" &&
    options !== null &&
    "encrypted" in options &&
    (options as { encrypted?: unknown }).encrypted !== undefined
  );
}

export function encryptedPassthroughMatches(source: Uint8Array, output: Uint8Array): boolean {
  return source.length === output.length && source.every((byte, index) => byte === output[index]);
}

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
  for (const entry of entries) {
    const item = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(item, out);
    else {
      const config = BY_EXT[entry.name.split(".").pop()!.toLowerCase()];
      if (config) out.push({ path: item, format: config.format, type: config.type });
    }
  }
  return out;
}

export async function runLibrary(
  libraryId: string,
  { destination, root, strictSemantic }: LibraryRunInput,
): Promise<LibraryRunResult> {
  const counts = {
    docx: { total: 0, clean: 0, diff: 0, parseFail: 0, genFail: 0 },
    xlsx: { total: 0, clean: 0, diff: 0, parseFail: 0, genFail: 0 },
    pptx: { total: 0, clean: 0, diff: 0, parseFail: 0, genFail: 0 },
  } as Record<Format, FormatCounts>;
  const blockerMaps: Record<Format, Map<string, number>> = {
    docx: new Map(),
    xlsx: new Map(),
    pptx: new Map(),
  };
  const diagnostics: FileDiagnostic[] = [];
  const rawAudit: Record<Format, { files: number; xmlParts: number; binaryParts: number }> = {
    docx: { files: 0, xmlParts: 0, binaryParts: 0 },
    xlsx: { files: 0, xmlParts: 0, binaryParts: 0 },
    pptx: { files: 0, xmlParts: 0, binaryParts: 0 },
  };
  const rawBlockerMaps: Record<Format, Map<string, number>> = {
    docx: new Map(),
    xlsx: new Map(),
    pptx: new Map(),
  };

  for (const { path: file, format, type } of walk(path.resolve(root, destination))) {
    const current = counts[format];
    current.total++;
    let options: unknown;
    let output: Uint8Array;
    const source = new Uint8Array(fs.readFileSync(file));
    try {
      options = await BY_EXT[path.extname(file).slice(1).toLowerCase()]!.parse(source);
    } catch (error) {
      current.parseFail++;
      diagnostics.push({
        file: path.relative(root, file),
        format,
        outcome: classifyPackageFailure(error),
        package: libraryId,
        error: String(error).slice(0, 240),
      });
      continue;
    }
    try {
      const packerOptions = { packageVariant: OOXML_PACKAGE_FORMATS[type].variant };
      output =
        format === "docx"
          ? await import("../../packages/docx/dist/index.mjs").then(({ generateDocument }) =>
              generateDocument(options as Parameters<typeof generateDocument>[0], packerOptions),
            )
          : format === "xlsx"
            ? await import("../../packages/xlsx/dist/index.mjs").then(({ generateWorkbook }) =>
                generateWorkbook(options as Parameters<typeof generateWorkbook>[0], packerOptions),
              )
            : await import("../../packages/pptx/dist/index.mjs").then(({ generatePresentation }) =>
                generatePresentation(
                  options as Parameters<typeof generatePresentation>[0],
                  packerOptions,
                ),
              );
    } catch (error) {
      current.genFail++;
      diagnostics.push({
        file: path.relative(root, file),
        format,
        outcome: "invalid-package",
        package: libraryId,
        error: String(error).slice(0, 240),
      });
      continue;
    }

    if (isEncryptedPassthrough(options)) {
      if (encryptedPassthroughMatches(source, output)) {
        current.clean++;
      } else {
        current.diff++;
        diagnostics.push({
          file: path.relative(root, file),
          format,
          outcome: "valid",
          package: libraryId,
          diffCategories: { binary: 1 },
          sampleParts: [
            {
              path: path.relative(root, file),
              kind: "binary",
              category: "binary",
              detail: "encrypted container bytes differ",
            },
          ],
        });
      }
      continue;
    }

    const rawParts = (options as { rawParts?: { path: string }[] }).rawParts;
    if (rawParts?.length) {
      rawAudit[format].files++;
      for (const part of rawParts) {
        if (/\.xml$|\.rels$/i.test(part.path)) rawAudit[format].xmlParts++;
        else rawAudit[format].binaryParts++;
      }
    }
    const { auditCanonicalOptions } = await import("./synthetic/raw-audit");
    const rawPartsForAudit = (options as { rawParts?: { path: string }[] }).rawParts;
    const orphanedXmlParts =
      (rawPartsForAudit?.some((part) => /\.xml$|\.rels$|\.vml$/i.test(part.path)) ?? false)
        ? orphanedPackageMembers(source)
        : undefined;
    for (const blocker of auditCanonicalOptions(
      options,
      EXTERNAL_OPAQUE_PARTS[format],
      orphanedXmlParts,
    )) {
      if (blocker.reason === "orphaned-independent-part") continue;
      const key = `${blocker.reason}:${blocker.part.replace(/(?:word|xl|ppt|powerpoint)[\\/]/, "")}`;
      rawBlockerMaps[format].set(key, (rawBlockerMaps[format].get(key) ?? 0) + 1);
    }

    let parts: string[];
    let semanticDiffs: SemanticPartDiff[] = [];
    try {
      assertEncryptedContainerRoundTrip(source, output);
      if (strictSemantic) {
        semanticDiffs = archiveSemanticDiffDetails(source, output);
        parts = semanticDiffs.map((diff) => diff.path);
      } else {
        parts = archiveTagDiffs(source, output);
      }
    } catch (error) {
      current.parseFail++;
      diagnostics.push({
        file: path.relative(root, file),
        format,
        outcome: "invalid-zip",
        package: libraryId,
        error: String(error).slice(0, 240),
      });
      continue;
    }

    if (parts.length === 0) current.clean++;
    else {
      current.diff++;
      for (const part of parts) {
        const key = part.replace(/(?:word|xl|ppt|powerpoint)[\\/]/, "");
        blockerMaps[format].set(key, (blockerMaps[format].get(key) ?? 0) + 1);
      }
      diagnostics.push({
        file: path.relative(root, file),
        format,
        outcome: "valid",
        package: libraryId,
        diffCategories: semanticDiffs.reduce<Record<string, number>>((countsByCategory, diff) => {
          const category = diff.category ?? "unknown";
          countsByCategory[category] = (countsByCategory[category] ?? 0) + 1;
          return countsByCategory;
        }, {}),
        sampleParts: semanticDiffs.slice(0, 20),
      });
    }
  }

  return {
    counts,
    blockers: {
      docx: [...blockerMaps.docx],
      xlsx: [...blockerMaps.xlsx],
      pptx: [...blockerMaps.pptx],
    },
    diagnostics,
    rawAudit,
    rawBlockers: {
      docx: [...rawBlockerMaps.docx],
      xlsx: [...rawBlockerMaps.xlsx],
      pptx: [...rawBlockerMaps.pptx],
    },
  };
}
