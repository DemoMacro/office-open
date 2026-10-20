import * as fs from "node:fs";
import * as path from "node:path";

type Format = "docx" | "xlsx" | "pptx";

interface SemanticDiff {
  path?: string;
  category?: string;
  xpath?: string;
  detail?: string;
}

interface ReportDiagnostic {
  file: string;
  format: Format;
  outcome: string;
  stage?: string;
  malformed?: boolean;
  diffs?: SemanticDiff[];
}

interface CorpusReport {
  diagnostics?: ReportDiagnostic[];
}

interface TriageCluster {
  format: Format;
  owner: string;
  part: string;
  category: string;
  signature: string;
  files: Set<string>;
  occurrences: number;
  examples: string[];
}

const OWNER_RULES: readonly { test: RegExp; owner: string }[] = [
  { test: /(?:^|\/)(?:wp|a|pic|wpg|wps|v):/, owner: "drawing" },
  { test: /(?:^|\/)c:/, owner: "chart" },
  { test: /(?:^|\/)dgm:/, owner: "smartart" },
  { test: /(?:^|\/)xdr:/, owner: "spreadsheet-drawing" },
  { test: /\[Content_Types\]\.xml$/i, owner: "opc-content-types" },
  { test: /\.rels$/i, owner: "opc-relationships" },
  { test: /^docProps\//i, owner: "package-properties" },
  {
    test: /^(?:styles|settings|numbering|fontTable|webSettings|customizations)/i,
    owner: "document-package",
  },
  { test: /^(?:footnotes|endnotes|headerN|footerN)/i, owner: "document-package" },
  {
    test: /^(?:worksheets|workbook|sharedStrings|pivotCache|pivotTable|tableN)/i,
    owner: "spreadsheet-package",
  },
  {
    test: /^(?:slides|slideLayouts|slideMasters|notesSlides|notesMasters|presProps|viewProps|tableStyles)/i,
    owner: "presentation-package",
  },
  { test: /^(?:diagrams|charts|embeddings|media|theme)/i, owner: "package-member" },
];

function posix(value: string): string {
  return value.replace(/\\/g, "/");
}

function normalizePart(value = ""): string {
  const normalized = posix(value)
    .replace(/^(?:word|xl|ppt)\//i, "")
    .replace(/\d+/g, "N");
  return normalized.split("/").slice(-2).join("/");
}

function normalizeSignature(xpath = ""): string {
  const tokens = xpath.split("/").filter(Boolean);
  const attribute = tokens.findLast((token) => token.startsWith("@"));
  const elements = tokens
    .filter((token) => !token.startsWith("@"))
    .slice(-4)
    .map((token) => token.replace(/\[\d+\]/g, "[*]"));
  const signature = elements.join("/");
  return attribute ? `${signature}/${attribute}` : signature || "/";
}

function normalizeDetail(value = ""): string {
  const sourceOrOutput = /^(source-only|output-only)\s+([^\s{]+)/.exec(value);
  if (sourceOrOutput) return `${sourceOrOutput[1]} ${sourceOrOutput[2]}`;
  if (value.startsWith("value mismatch")) return "value mismatch";
  if (value.startsWith("children reordered")) return "children reordered";
  if (value.startsWith("text mismatch")) return "text mismatch";
  if (value.startsWith("output missing")) return "output missing";
  if (value.startsWith("bytes differ")) return "bytes differ";
  return value.split(/\s+/).slice(0, 3).join(" ");
}

function ownerFor(diff: SemanticDiff): string {
  const haystack = `${posix(diff.path ?? "")} ${diff.xpath ?? ""}`;
  for (const rule of OWNER_RULES) {
    if (rule.test.test(haystack)) return rule.owner;
  }
  return "unclassified";
}

function clusterKey(format: Format, diff: SemanticDiff): string {
  return [
    format,
    ownerFor(diff),
    normalizePart(diff.path),
    diff.category ?? "unknown",
    normalizeDetail(diff.detail),
    normalizeSignature(diff.xpath),
  ].join("|");
}

function triage(report: CorpusReport): TriageCluster[] {
  const clusters = new Map<string, TriageCluster>();
  for (const diagnostic of report.diagnostics ?? []) {
    if (diagnostic.outcome !== "valid") {
      if (diagnostic.stage === "parse" && diagnostic.malformed === false) {
        const key = clusterKey(diagnostic.format, {
          path: "parse",
          category: "error",
          detail: diagnostic.error,
        });
        const cluster = clusters.get(key) ?? {
          format: diagnostic.format,
          owner: "parse",
          part: "parse",
          category: "error",
          signature: normalizeDetail(diagnostic.error),
          files: new Set<string>(),
          occurrences: 0,
          examples: [],
        };
        cluster.files.add(diagnostic.file);
        cluster.occurrences += 1;
        if (cluster.examples.length < 5)
          cluster.examples.push(path.posix.basename(posix(diagnostic.file)));
        clusters.set(key, cluster);
      }
      continue;
    }
    for (const diff of diagnostic.diffs ?? diagnostic.sampleParts ?? []) {
      const key = clusterKey(diagnostic.format, diff);
      const cluster = clusters.get(key) ?? {
        format: diagnostic.format,
        owner: ownerFor(diff),
        part: normalizePart(diff.path),
        category: diff.category ?? "unknown",
        signature: `${normalizeDetail(diff.detail)} @ ${normalizeSignature(diff.xpath)}`,
        files: new Set<string>(),
        occurrences: 0,
        examples: [],
      };
      cluster.files.add(diagnostic.file);
      cluster.occurrences += 1;
      if (cluster.examples.length < 5)
        cluster.examples.push(path.posix.basename(posix(diagnostic.file)));
      clusters.set(key, cluster);
    }
  }
  return [...clusters.values()].sort(
    (left, right) =>
      right.files.size - left.files.size ||
      right.occurrences - left.occurrences ||
      left.signature.localeCompare(right.signature),
  );
}

function loadReport(reportPath: string): CorpusReport {
  return JSON.parse(fs.readFileSync(path.resolve(reportPath), "utf8")) as CorpusReport;
}

function printSummary(report: CorpusReport, clusters: TriageCluster[], limit: number): void {
  const failures = new Map<Format, { diff: number; parse: number; generate: number }>([
    ["docx", { diff: 0, parse: 0, generate: 0 }],
    ["xlsx", { diff: 0, parse: 0, generate: 0 }],
    ["pptx", { diff: 0, parse: 0, generate: 0 }],
  ]);
  for (const diagnostic of report.diagnostics ?? []) {
    const counts = failures.get(diagnostic.format);
    if (!counts) continue;
    if (diagnostic.outcome === "valid") counts.diff++;
    else if (diagnostic.stage === "parse" && diagnostic.malformed === false) counts.parse++;
    else if (diagnostic.stage === "generate") counts.generate++;
  }
  console.log("actionable failures by format:");
  for (const format of ["docx", "xlsx", "pptx"] as const) {
    const item = failures.get(format);
    if (item)
      console.log(`  ${format}: diff=${item.diff} parse=${item.parse} generate=${item.generate}`);
  }
  console.log(`\ntop ${Math.min(limit, clusters.length)} root-cause clusters:`);
  for (const cluster of clusters.slice(0, limit)) {
    console.log(
      `  [${cluster.format}] ${cluster.owner} · ${cluster.part} · ${cluster.category} · ` +
        `files=${cluster.files.size} occurrences=${cluster.occurrences}\n    ${cluster.signature}`,
    );
  }
}

const args = process.argv.slice(2);
const reportArgIndex = args.indexOf("--report");
const reportPath =
  reportArgIndex >= 0
    ? (args[reportArgIndex + 1] ?? ".temp/corpus-report.json")
    : ".temp/corpus-report.json";
const jsonArgIndex = args.indexOf("--json");
const jsonPath = jsonArgIndex >= 0 ? args[jsonArgIndex + 1] : undefined;
const limitArgIndex = args.indexOf("--limit");
const limit = Number(limitArgIndex >= 0 ? args[limitArgIndex + 1] : 30) || 30;

const report = loadReport(reportPath);
const clusters = triage(report);
printSummary(report, clusters, limit);

if (jsonPath) {
  const serializable = clusters.map(({ files, examples, ...cluster }) => ({
    ...cluster,
    fileCount: files.size,
    exampleFiles: examples,
  }));
  fs.mkdirSync(path.dirname(path.resolve(jsonPath)), { recursive: true });
  fs.writeFileSync(jsonPath, `${JSON.stringify(serializable, null, 2)}\n`);
}
