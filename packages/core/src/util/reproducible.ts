/**
 * Opt-in reproducible generation: the same input produces the same bytes.
 *
 * The scope is a plain value object — create it once per generation and thread
 * it explicitly through the packer options ({@link
 * PackerOptions.reproducible}); it never installs global state, so concurrent
 * generations are isolated by construction.
 *
 * @module
 */

/**
 * Options for creating a reproducible scope ({@link createReproducibleScope}).
 */
export interface ReproducibleGenerationOptions {
  /**
   * ISO-8601 timestamp substituted for every wall-clock default
   * (dcterms:created/modified, comment w:date). Defaults to the Unix epoch,
   * "1970-01-01T00:00:00.000Z".
   */
  date?: string;
}

/**
 * Per-generation deterministic state: a fixed ISO-8601 date plus
 * counter-backed replacements for the random id generators (alphanumeric ids,
 * UUIDs, VML shape ids, drawing ids). Scopes are independent — counters start
 * at the same values each time.
 */
export interface ReproducibleScope {
  /** ISO-8601 timestamp used in place of the wall clock. */
  readonly date: string;
  /** Deterministic replacement for uniqueId: a 21-char lowercase alphanumeric id. */
  nextId(): string;
  /** Deterministic UUID replacement for uniqueUuid/crypto.randomUUID. */
  nextUuid(): string;
  /** Next VML shape id (the numeric part of `_x0000_sN`). */
  nextVmlShapeId(): number;
  /** Next wp:docPr id. */
  nextDrawingId(): number;
}

/** Unix epoch — the conventional fixed timestamp for reproducible builds. */
const EPOCH = "1970-01-01T00:00:00.000Z";

/**
 * Create a fresh deterministic scope. Counters start where the un-scoped
 * readers would: VML ids at 1024 (first id 1025), the rest at 0.
 */
export function createReproducibleScope(
  options?: ReproducibleGenerationOptions,
): ReproducibleScope {
  let ids = 0;
  let uuids = 0;
  let vmlShapeIds = 1024;
  let drawingIds = 0;
  return {
    date: options?.date ?? EPOCH,
    nextId: () => (++ids).toString(36).padStart(21, "0"),
    nextUuid: () =>
      `00000000-0000-4000-8000-${(++uuids).toString(16).padStart(12, "0").slice(-12)}`,
    nextVmlShapeId: () => ++vmlShapeIds,
    nextDrawingId: () => ++drawingIds,
  };
}
