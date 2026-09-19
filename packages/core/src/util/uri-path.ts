/**
 * Percent-encoding for OPC part names and relationship Targets.
 *
 * @module
 */

/**
 * Percent-encode each path segment of a URI reference, keeping the `/`
 * separators: a Target is a URI, so spaces, `#`, `?` and non-ASCII in a part
 * name must be escaped. Not idempotent — an already-encoded segment gets its
 * `%` re-escaped — so callers pass the raw path exactly once (parse stores the
 * decoded form precisely so re-compilation escapes once).
 */
export function encodeUriPath(path: string): string {
  return path.split("/").map(encodeURIComponent).join("/");
}

/**
 * Decode each path segment percent-encoded by {@link encodeUriPath} — the
 * lookup form for a ZIP entry stored under the raw part name. Returns the
 * input unchanged when a segment is not valid percent-encoding.
 */
export function decodeUriPath(path: string): string {
  try {
    return path.split("/").map(decodeURIComponent).join("/");
  } catch {
    return path;
  }
}
