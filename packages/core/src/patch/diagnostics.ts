/** Part path relative to the package root. */
export type PatchPart = string;

/** Stable dot-separated selector for the operation field or model node. */
export type PatchPath = string;

/** Structured reason emitted by a patch compiler or context. */
export type PatchReason =
  | "invalid-operation"
  | "invalid-target"
  | "invalid-part"
  | "missing-part"
  | "missing-target"
  | "unsupported-content";

/** Structured diagnostic for invalid or ambiguous patch input. */
export interface PatchDiagnostic {
  /** Package part being patched, or `"[package]"` for OPC-level edits. */
  part: PatchPart;
  /** Canonical path to the target operation or model node. */
  path: PatchPath;
  /** Operation type when the diagnostic belongs to one operation. */
  operation?: string;
  /** Stable classification used by callers and tests. */
  reason: PatchReason;
  /** Human-readable explanation without embedding source document content. */
  message: string;
}

/** Error raised for invalid or impossible source-preserving patch edits. */
export class PatchError extends Error {
  public readonly diagnostic: PatchDiagnostic;

  constructor(diagnostic: PatchDiagnostic) {
    super(diagnostic.message);
    this.name = "PatchError";
    this.diagnostic = diagnostic;
  }
}
