import { opaquePassthroughPolicy, unzipSync, zipSync } from "../../../packages/core/dist/index.mjs";
import { auditCanonicalOptions } from "./raw-audit";

export const ENCODER = new TextEncoder();
export const DECODER = new TextDecoder();
const ZIP_MTIME = new Date("1980-01-01T00:00:00Z").getTime();

export class ProjectionError extends Error {
  constructor(
    readonly part: string,
    readonly reason: string,
  ) {
    super(reason);
  }
}

export function assert(condition: unknown, part: string, reason: string): asserts condition {
  if (!condition) throw new ProjectionError(part, reason);
}

export function assertEqual(
  actual: unknown,
  expected: unknown,
  part: string,
  reason: string,
): void {
  if (actual !== expected) {
    const text = (value: unknown) =>
      JSON.stringify(value, (_, child) => {
        if (child instanceof Uint8Array) return Array.from(child);
        return child;
      });
    throw new ProjectionError(part, `${reason}: expected ${text(expected)}, got ${text(actual)}`);
  }
}

export function syntheticZip(files: Record<string, Uint8Array>): Uint8Array {
  const reproducible = Object.fromEntries(
    Object.entries(files).map(([path, data]) => [path, [data, { mtime: ZIP_MTIME }]]),
  );
  return zipSync(reproducible);
}

export function partText(packageBytes: Uint8Array, path: string): string {
  const value = unzipSync(packageBytes)[path];
  assert(value, path, "generated part is missing");
  return DECODER.decode(value);
}

export function auditSyntheticOptions(options: unknown, part: string): void {
  const blockers = auditCanonicalOptions(options, opaquePassthroughPolicy("docx").opaquePatterns);
  const blocker = blockers[0];
  if (blocker) throw new ProjectionError(part, `${blocker.reason}:${blocker.part}`);
}
