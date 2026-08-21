/**
 * Extension-uri matching for a:ext/p:ext style extension lists.
 *
 * @module
 */

/**
 * Match an `@uri` attribute against a known extension uri. Producers write
 * GUID-style uris with or without the enclosing braces and in mixed case —
 * both spellings denote the same extension, so compare brace-stripped and
 * case-insensitively.
 */
export function extUriMatches(actual: string | undefined, expected: string): boolean {
  if (actual === undefined) return false;
  const norm = (uri: string) => uri.replace(/^[{]+|[}]+$/g, "").toLowerCase();
  return norm(actual) === norm(expected);
}
