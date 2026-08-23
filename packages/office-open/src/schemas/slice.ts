/**
 * Schema slicing — extract an on-demand sub-schema by definition name.
 *
 * A full format schema (docx: 442 definitions, ~675 KB) is far too large to
 * hand to an LLM in one piece. `sliceSchema` walks the `$ref` dependency
 * closure of the requested definitions and returns only that subgraph.
 * Cataloged entry names (see entries.ts) act as stub boundaries: any entry
 * that is requested expands fully; any entry merely encountered collapses to
 * an empty-schema stub that names the follow-up lookup — recursive
 * disclosure, the same pattern Agent Skills use for reference files.
 *
 * The slice is a pure subgraph extraction: field names, descriptions, and
 * enums are kept verbatim; the source schema is never mutated.
 *
 * @module
 */

import { entryNames } from "./entries";
import { SCHEMAS, type DocumentType, type JsonSchema } from "./schemas";

/** Error thrown for unknown definition names; carries did-you-mean suggestions. */
export class UnknownDefinitionError extends Error {
  readonly suggestions: readonly string[];
  constructor(names: readonly string[], suggestions: readonly string[]) {
    super(`Unknown definition(s): ${names.join(", ")}`);
    this.name = "UnknownDefinitionError";
    this.suggestions = suggestions;
  }
}

/** Definition names referenced (directly or transitively) by one definition. */
function refsOf(def: unknown): string[] {
  const refs = JSON.stringify(def)?.match(/#\/definitions\/([A-Za-z0-9_%.$-]+)/g) ?? [];
  return [...new Set(refs.map((r) => decodeURIComponent(r.slice(14))))];
}

/**
 * Build the empty-schema stub that replaces a non-requested entry.
 * No `type` keyword: boundary definitions may legitimately be scalars
 * (UniversalMeasure strings, enums), and `type: "object"` would wrongly
 * reject those values. Kept minimal — a slice can carry hundreds of stubs,
 * so every byte here is multiplied.
 */
function buildStub(
  name: string,
  original: Record<string, unknown>,
  format: DocumentType | undefined,
): Record<string, unknown> {
  const description = typeof original.description === "string" ? `${original.description}\n` : "";
  return {
    title: `${name} (stub)`,
    description:
      `${description}` +
      `Accepts any value. Expand: office-open-schema-lookup ` +
      (format
        ? `{ type: "${format}", definitions: ["${name}"] }`
        : `with definitions: ["${name}"]`),
  };
}

/**
 * Hard cap on the serialized slice handed back to the model. Tool responses
 * must stay well inside a model's comfortable context budget (Anthropic's
 * tool guidance caps responses at 25k tokens); the entry catalog is the
 * primary size control and this cap is the backstop — a slice that still
 * exceeds it demotes its largest non-requested definitions to stubs until
 * it fits.
 */
const MAX_SLICE_BYTES = 64 * 1024;

/**
 * Extract the `$ref` closure of `definitions` from `schema`.
 *
 * Cataloged entries that were not requested become stubs, which keeps the
 * slice proportional to the request instead of the whole schema. The result
 * is a standalone draft-07 schema: when a single definition is requested it
 * is also wired as the root `$ref` so `ajv.compile(slice)` validates that
 * fragment directly.
 */
export function sliceSchema(
  schema: JsonSchema,
  definitions: readonly string[],
  format?: DocumentType,
): JsonSchema {
  const all = schema.definitions as Record<string, Record<string, unknown>> | undefined;
  if (!all) throw new Error("schema has no definitions");

  const unknown = definitions.filter((name) => !(name in all));
  if (unknown.length > 0)
    throw new UnknownDefinitionError(unknown, suggestNames(Object.keys(all), unknown[0]));

  const stubs = format ? entryNames(format) : new Set<string>();
  const requested = new Set(definitions);

  const expanded = new Set<string>();
  const stubbed = new Set<string>();
  const queue = [...definitions];
  while (queue.length > 0) {
    const name = queue.pop()!;
    if (expanded.has(name) || stubbed.has(name)) continue;
    if (stubs.has(name) && !requested.has(name)) {
      stubbed.add(name);
      continue;
    }
    expanded.add(name);
    for (const ref of refsOf(all[name])) {
      if (ref in all) queue.push(ref);
    }
  }

  const out: Record<string, unknown> = {};
  for (const name of expanded) out[name] = structuredClone(all[name]);
  for (const name of stubbed) out[name] = buildStub(name, all[name], format);

  // Size backstop: demote the largest non-requested definitions to stubs
  // until the slice fits. Requested definitions are never demoted.
  while (JSON.stringify(out).length > MAX_SLICE_BYTES) {
    const candidates = [...expanded]
      .filter((name) => !requested.has(name))
      .map((name) => [name, JSON.stringify(out[name]).length] as const)
      .sort((a, b) => b[1] - a[1]);
    if (candidates.length === 0) break;
    const [name] = candidates[0]!;
    expanded.delete(name);
    stubbed.add(name);
    out[name] = buildStub(name, all[name], format);
  }

  const result: Record<string, unknown> = {
    $schema: "http://json-schema.org/draft-07/schema#",
    title: `office-open ${format ?? ""} schema slice: ${definitions.join(", ")}`.trim(),
    description:
      "Sub-schema extracted on demand. Definitions marked as stubs accept any value — " +
      "look one up by name to expand it. The full schema ships with the package.",
    definitions: out,
  };
  if (definitions.length === 1) result.$ref = `#/definitions/${definitions[0]}`;
  return result;
}

/** Slice a format's schema by definition name (stub guidance names the format). */
export function sliceDocumentSchema(
  type: DocumentType,
  definitions: readonly string[],
): JsonSchema {
  return sliceSchema(SCHEMAS[type], definitions, type);
}

/** Throw {@link UnknownDefinitionError} for names missing from the format's schema. */
export function assertKnownDefinitions(type: DocumentType, definitions: readonly string[]): void {
  const all = SCHEMAS[type].definitions as Record<string, unknown>;
  const unknown = definitions.filter((name) => !(name in all));
  if (unknown.length > 0) {
    throw new UnknownDefinitionError(unknown, suggestNames(Object.keys(all), unknown[0]));
  }
}

/** Up to three close candidates for an unknown name (substring match, then edit distance). */
function suggestNames(candidates: readonly string[], query: string): string[] {
  const lower = query.toLowerCase();
  const contains = candidates.filter(
    (c) => c.toLowerCase().includes(lower) || lower.includes(c.toLowerCase()),
  );
  if (contains.length > 0) return contains.slice(0, 3);
  return candidates.filter((c) => levenshtein(c.toLowerCase(), lower) <= 3).slice(0, 3);
}

/** Classic Levenshtein distance, bounded to short strings. */
function levenshtein(a: string, b: string): number {
  const m = a.length;
  const n = b.length;
  if (Math.abs(m - n) > 3) return 4;
  let prev = Array.from({ length: n + 1 }, (_, i) => i);
  for (let i = 1; i <= m; i++) {
    const curr = [i];
    for (let j = 1; j <= n; j++) {
      curr[j] = Math.min(
        prev[j]! + 1,
        curr[j - 1]! + 1,
        prev[j - 1]! + (a[i - 1] === b[j - 1] ? 0 : 1),
      );
    }
    prev = curr;
  }
  return prev[n]!;
}
