export type ModelCoverageClassification =
  | "modeled"
  | "opaque-passthrough"
  | "not-modeled-by-design"
  | "planned";

export interface ModelOwner {
  stringify: string | null;
  parse: string | null;
  compiler: string | null;
  publicApi: string | null;
}

export interface ModelEvidence {
  static: readonly string[];
  executable: readonly string[];
  rawOnly: boolean;
}

export interface ModelOwnershipEntry {
  domain: string;
  qname: string;
  canonicalOptionPath: string | null;
  classification: ModelCoverageClassification;
  classificationReason: string;
  owner: ModelOwner;
  evidence: ModelEvidence;
}

export interface ModelCoverageViolation {
  severity: "error";
  code:
    | "unowned-qname"
    | "modeled-without-option-path"
    | "modeled-without-both-directions"
    | "modeled-without-executable-evidence"
    | "raw-only-modeled-evidence"
    | "unknown-ownership-qname";
  domain: string;
  qname: string;
  reason: string;
}

export interface ModelCoverageReport {
  schemaVersion: 1;
  domains: readonly string[];
  entries: readonly ModelOwnershipEntry[];
  unowned: readonly { domain: string; qname: string }[];
  summary: Readonly<Record<string, number>>;
  violations: readonly ModelCoverageViolation[];
}
