/**
 * Opt-in reproducible generation: the same input produces the same bytes.
 *
 * @module
 */

/**
 * Options for {@link withReproducibleGeneration}.
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
 * Per-generation deterministic state. Read it through
 * {@link activeReproducibleScope}: the scope exists only between
 * {@link withReproducibleGeneration} and its return.
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
 * `globalThis` slot holding the scope stack. `Symbol.for` (not a module-level
 * binding) so every copy of core in one process shares the same scopes.
 */
const REPRODUCIBLE_SCOPE = Symbol.for("office-open.reproducible");

/** The process-wide scope stack, created on first use. */
function scopeStack(): ReproducibleScope[] {
  const globals = globalThis as Record<symbol, unknown>;
  const existing = globals[REPRODUCIBLE_SCOPE];
  if (Array.isArray(existing)) return existing as ReproducibleScope[];
  const stack: ReproducibleScope[] = [];
  globals[REPRODUCIBLE_SCOPE] = stack;
  return stack;
}

/** Counters start where the un-scoped readers would: VML ids at 1024 (first
 *  id 1025), the rest at 0. */
function createScope(date: string): ReproducibleScope {
  let ids = 0;
  let uuids = 0;
  let vmlShapeIds = 1024;
  let drawingIds = 0;
  return {
    date,
    nextId: () => (++ids).toString(36).padStart(21, "0"),
    nextUuid: () =>
      `00000000-0000-4000-8000-${(++uuids).toString(16).padStart(12, "0").slice(-12)}`,
    nextVmlShapeId: () => ++vmlShapeIds,
    nextDrawingId: () => ++drawingIds,
  };
}

/**
 * The scope that id and date reads resolve against, or undefined outside
 * {@link withReproducibleGeneration}. With nested scopes the innermost wins.
 */
export function activeReproducibleScope(): ReproducibleScope | undefined {
  const stack = scopeStack();
  return stack[stack.length - 1];
}

/**
 * Runs `fn` with deterministic auto-ids, date defaults and ZIP timestamps, so
 * the same input yields byte-identical output. Scopes nest (innermost wins).
 * An async `fn` keeps its scope installed until the returned promise settles;
 * concurrent generations each get their own scope because id reads happen
 * synchronously during compile, before any packer await.
 */
export function withReproducibleGeneration<T>(
  options: ReproducibleGenerationOptions,
  fn: () => T,
): T {
  const stack = scopeStack();
  const scope = createScope(options.date ?? EPOCH);
  stack.push(scope);
  const restore = (): void => {
    const index = stack.lastIndexOf(scope);
    if (index >= 0) stack.splice(index, 1);
  };
  try {
    const result = fn();
    if (result != null && typeof (result as { then?: unknown }).then === "function") {
      return Promise.resolve(result).finally(restore) as T;
    }
    restore();
    return result;
  } catch (error) {
    restore();
    throw error;
  }
}
