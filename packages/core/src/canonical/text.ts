import { sanitizeXmlText } from "@office-open/xml";

/**
 * Normalize every string in a canonical options tree for XML serialization.
 *
 * Binary payloads and dates are preserved unchanged; canonical options are
 * otherwise expected to be plain serializable data.
 */
export function sanitizeOptionsText<T>(value: T): T {
  if (typeof value === "string") return sanitizeXmlText(value) as T;
  if (value instanceof Uint8Array || value instanceof Date) return value;
  if (Array.isArray(value)) return value.map(sanitizeOptionsText) as T;
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, child]) => [key, sanitizeOptionsText(child)]),
    ) as T;
  }
  return value;
}
