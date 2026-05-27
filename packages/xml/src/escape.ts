const XML_CHAR_MAP: Record<string, string> = {
  "&": "&amp;",
  '"': "&quot;",
  "'": "&apos;",
  "<": "&lt;",
  ">": "&gt;",
};

const XML_CHAR_PATTERN = /([&"<>'])/g;

/** Escape text content for XML. Fast path returns original string when no special chars. */
export function escapeXml(str: string): string {
  // Fast path: most text content doesn't contain XML-special characters.
  // Manual scan avoids regex overhead; returning the original string reference
  // means zero allocation for the common case.
  for (let i = 0; i < str.length; i++) {
    const c = str.charCodeAt(i);
    if (c === 38 || c === 34 || c === 39 || c === 60 || c === 62) {
      // & " ' < >
      return str.replace(XML_CHAR_PATTERN, (ch) => XML_CHAR_MAP[ch]);
    }
  }
  return str;
}

/**
 * Escape attribute value matching xml-js's js2xml behavior.
 * Handles already-escaped entities to prevent double-escaping.
 */
export function escapeAttributeValue(str: string): string {
  return String(str)
    .replace(/&(?!amp;|lt;|gt;|quot;|apos;)/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Build an XML attribute string fragment from a record.
 * `undefined` values are automatically skipped.
 * String values are escaped via `escapeXml`.
 *
 * @example
 * attrs({ id: 1, name: "foo", hidden: undefined })
 * // => ' id="1" name="foo"'
 */
export function attrs(record: Record<string, string | number | boolean | undefined>): string {
  let s = "";
  for (const [k, v] of Object.entries(record)) {
    if (v !== undefined) {
      s += ` ${k}="${typeof v === "string" ? escapeXml(v) : v}"`;
    }
  }
  return s;
}
