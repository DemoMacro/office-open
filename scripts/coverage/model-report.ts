import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { parse } from "../../packages/xml/dist/index.mjs";
import type { OoxmlExecutableEvidenceManifest } from "./ooxml-executable";
import { MODEL_OWNERSHIP } from "./ownership";
import type {
  ModelCoverageReport,
  ModelCoverageViolation,
  ModelOwnershipEntry,
} from "./ownership/types";

const ROOT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const REPORT_PATH = path.join(ROOT_DIR, ".temp", "model-coverage.json");
const EVIDENCE_PATH = path.join(ROOT_DIR, ".temp", "ooxml-executable-evidence.json");
const DOMAIN_SCHEMAS = [
  { domain: "dml-chart", schema: "ooxml-schemas/transitional/dml-chart.xsd", prefix: "c" },
  { domain: "dml-diagram", schema: "ooxml-schemas/transitional/dml-diagram.xsd", prefix: "dgm" },
  { domain: "dml-main", schema: "ooxml-schemas/transitional/dml-main.xsd", prefix: "a" },
  {
    domain: "dml-wp",
    schema: "ooxml-schemas/transitional/dml-wordprocessingDrawing.xsd",
    prefix: "wp",
  },
  { domain: "wml", schema: "ooxml-schemas/transitional/wml.xsd", prefix: "w" },
  { domain: "math", schema: "ooxml-schemas/transitional/shared-math.xsd", prefix: "m" },
  { domain: "sml", schema: "ooxml-schemas/transitional/sml.xsd", prefix: "sml" },
  { domain: "pml", schema: "ooxml-schemas/transitional/pml.xsd", prefix: "p" },
  { domain: "vml-main", schema: "ooxml-schemas/transitional/vml-main.xsd", prefix: "v" },
] as const;

const DOMAINS = DOMAIN_SCHEMAS.map((entry) => entry.domain);

function collectSchemaQNames(schemaPath: string, prefix: string): Set<string> {
  const content = fs.readFileSync(path.join(ROOT_DIR, schemaPath), "utf8");
  const schema = parse(content).elements?.[0];
  if (!schema) throw new Error(`${schemaPath}: no root element`);
  const qnames = new Set<string>();
  const walk = (element: typeof schema): void => {
    if (
      (element.name === "xsd:element" || element.name === "xs:element") &&
      element.attributes?.name
    ) {
      qnames.add(`${prefix}:${element.attributes.name}`);
    }
    for (const child of element.elements ?? []) walk(child);
  };
  walk(schema);
  return qnames;
}

function violation(
  code: ModelCoverageViolation["code"],
  domain: string,
  qname: string,
  reason: string,
): ModelCoverageViolation {
  return { severity: "error", code, domain, qname, reason };
}

function ownerArtifactExists(reference: string): boolean {
  const [file, symbol] = reference.split(":");
  if (!file) return false;
  const filePath = path.join(ROOT_DIR, file);
  if (!fs.existsSync(filePath)) return false;
  return !symbol || fs.readFileSync(filePath, "utf8").includes(symbol);
}

function validateEvidence(entry: ModelOwnershipEntry): boolean {
  const owner = [entry.owner.stringify, entry.owner.parse].filter(
    (reference): reference is string => Boolean(reference),
  );
  return owner.every(ownerArtifactExists);
}

function executableEvidenceStatus(
  entry: ModelOwnershipEntry,
  evidence: OoxmlExecutableEvidenceManifest,
): { ok: boolean; reason: string } {
  if (entry.evidence.executable.length === 0) {
    return { ok: false, reason: "No executable case is present." };
  }
  const missing = entry.evidence.executable.filter((caseId) => {
    const record = evidence.records.find((candidate) => candidate.caseId === caseId);
    return record?.status !== "passed" || !record.coveredElements.includes(entry.qname);
  });
  if (missing.length > 0) {
    return {
      ok: false,
      reason: `Executable cases do not pass with ${entry.qname}: ${missing.join(", ")}.`,
    };
  }
  return { ok: true, reason: "" };
}

function loadEvidenceManifest(): OoxmlExecutableEvidenceManifest {
  if (!fs.existsSync(EVIDENCE_PATH)) {
    throw new Error("Missing executable evidence; run pnpm coverage:ooxml-executable first.");
  }
  const manifest = JSON.parse(fs.readFileSync(EVIDENCE_PATH, "utf8"));
  if (manifest.schemaVersion !== 1 || !Array.isArray(manifest.records)) {
    throw new Error("Executable evidence manifest is invalid.");
  }
  return manifest;
}

export function buildModelCoverageReport(
  evidenceManifest: OoxmlExecutableEvidenceManifest = loadEvidenceManifest(),
): ModelCoverageReport {
  const executableEvidenceByQName = new Map<string, string[]>();
  for (const record of evidenceManifest.records) {
    if (record.status !== "passed") continue;
    for (const qname of record.coveredElements) {
      executableEvidenceByQName.set(qname, [
        ...(executableEvidenceByQName.get(qname) ?? []),
        record.caseId,
      ]);
    }
  }
  const entries = [...MODEL_OWNERSHIP]
    .map((entry) => ({
      ...entry,
      evidence: {
        ...entry.evidence,
        executable: executableEvidenceByQName.get(entry.qname) ?? [],
      },
    }))
    .sort(
      (left, right) =>
        left.domain.localeCompare(right.domain) || left.qname.localeCompare(right.qname),
    );
  const violations: ModelCoverageViolation[] = [];
  const unowned: { domain: string; qname: string }[] = [];
  const owned = new Set(entries.map((entry) => `${entry.domain}:${entry.qname}`));

  const schemaQNames = new Set<string>();
  for (const schema of DOMAIN_SCHEMAS) {
    for (const qname of collectSchemaQNames(schema.schema, schema.prefix)) {
      schemaQNames.add(qname);
      if (!owned.has(`${schema.domain}:${qname}`)) {
        unowned.push({ domain: schema.domain, qname });
        violations.push(
          violation(
            "unowned-qname",
            schema.domain,
            qname,
            "XSD global element is missing from explicit model ownership.",
          ),
        );
      }
    }
  }

  for (const entry of entries) {
    if (!schemaQNames.has(entry.qname)) {
      violations.push(
        violation(
          "unknown-ownership-qname",
          entry.domain,
          entry.qname,
          "Ownership entry is not declared by the source schema.",
        ),
      );
    }
    if (!DOMAINS.includes(entry.domain as (typeof DOMAINS)[number])) {
      violations.push(
        violation(
          "unknown-ownership-qname",
          entry.domain,
          entry.qname,
          `Ownership domain is not enabled; expected one of ${DOMAINS.join(", ")}.`,
        ),
      );
      continue;
    }
    if (entry.classification !== "modeled") continue;
    if (!entry.canonicalOptionPath) {
      violations.push(
        violation(
          "modeled-without-option-path",
          entry.domain,
          entry.qname,
          "A modeled entry must name its canonical option path.",
        ),
      );
    }
    if (!entry.owner.stringify || !entry.owner.parse) {
      violations.push(
        violation(
          "modeled-without-both-directions",
          entry.domain,
          entry.qname,
          "A modeled entry must name both stringify and parse owners.",
        ),
      );
    }
    if (entry.evidence.executable.length === 0 || entry.evidence.rawOnly) {
      violations.push(
        violation(
          entry.evidence.rawOnly
            ? "raw-only-modeled-evidence"
            : "modeled-without-executable-evidence",
          entry.domain,
          entry.qname,
          "A modeled entry must be exercised by a real parse/generate test.",
        ),
      );
    }
    const evidenceStatus = executableEvidenceStatus(entry, evidenceManifest);
    if (!evidenceStatus.ok) {
      violations.push(
        violation(
          "modeled-without-executable-evidence",
          entry.domain,
          entry.qname,
          evidenceStatus.reason,
        ),
      );
    } else if (!validateEvidence(entry)) {
      violations.push(
        violation(
          "modeled-without-executable-evidence",
          entry.domain,
          entry.qname,
          "An owner source file or symbol could not be found.",
        ),
      );
    }
  }

  const summary: Record<string, number> = {
    total: entries.length,
    unowned: unowned.length,
  };
  for (const entry of entries) {
    summary[entry.classification] = (summary[entry.classification] ?? 0) + 1;
  }

  return {
    schemaVersion: 1,
    domains: DOMAINS,
    entries,
    unowned,
    summary,
    violations,
  };
}

function main(): void {
  const report = buildModelCoverageReport();
  fs.mkdirSync(path.dirname(REPORT_PATH), { recursive: true });
  fs.writeFileSync(REPORT_PATH, `${JSON.stringify(report, null, 2)}\n`);

  console.log(`model coverage: ${JSON.stringify(report.summary)}`);
  for (const item of report.violations) {
    console.error(`${item.domain}:${item.qname}: ${item.code} — ${item.reason}`);
  }
  if (report.violations.length > 0) process.exitCode = 1;
}

main();
