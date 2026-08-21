// Non-global on purpose: a /g regex would carry stateful lastIndex across calls.
const XML_SPECIALS = /[&"'<>]/;

/** Escape text content for XML. Fast path returns original string when no special chars. */
export function escapeXml(str: string): string {
  // Fast path: most text content doesn't contain XML-special characters.
  // A character-class regex test beats a manual charCodeAt loop by ~10× (V8
  // compiles it to a native SIMD scan); returning the original string
  // reference means zero allocation for the common case.
  if (!XML_SPECIALS.test(str)) return str;

  // Chained native replaces: specials are sparse, so each pass after the
  // first usually finds nothing and exits quickly. Raw /g literals carry no
  // lastIndex state.
  return str
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * Build an XML attribute string fragment from a record.
 * `undefined` values are automatically skipped.
 * String values are escaped via `escapeXml`. Booleans serialize as 0/1 —
 * the spelling Office itself writes for xsd:boolean attributes and
 * ST_OnOff unions alike.
 *
 * @example
 * attrs({ id: 1, name: "foo", hidden: undefined })
 * // => ' id="1" name="foo"'
 */
export function attrs(record: Record<string, string | number | boolean | undefined>): string {
  const parts: string[] = [];
  for (const [key, v] of Object.entries(record)) {
    if (v !== undefined) {
      const value = typeof v === "string" ? escapeXml(v) : typeof v === "boolean" ? (v ? 1 : 0) : v;
      parts.push(` ${key}="${value}"`);
    }
  }
  return parts.join("");
}

/**
 * Build an XML attribute string without escaping.
 *
 * Same as `attrs()` but skips `typeof` checks and `escapeXml` — use only when
 * all values are known-safe (numbers, booleans, or strings free of `& " ' < >`).
 * Avoids per-call array and `Object.keys()` allocation in hot loops.
 *
 * @example
 * attrsRaw({ r: "A1", s: 5 })
 * // => ' r="A1" s="5"'
 */
export function attrsRaw(record: Record<string, string | number | boolean | undefined>): string {
  let s = "";
  for (const key in record) {
    const v = record[key];
    if (v !== undefined) {
      s += ` ${key}="${v}"`;
    }
  }
  return s;
}

/**
 * Build a self-closing XML element: `<tag attrStr/>`.
 * `attrStr` is a pre-serialized attribute string (from `attrs()`) or undefined.
 */
export function selfCloseElement(tag: string, attrStr?: string): string {
  return attrStr ? `<${tag}${attrStr}/>` : `<${tag}/>`;
}

/**
 * Build a complete XML element string from name, optional attributes, and string children.
 *
 * Replaces `new BuilderElement({...})` + `.toXml()` / `.serialize()` with a
 * single function call returning a string — zero object allocation.
 *
 * @param name  Element tag name (e.g. `"a:srgbClr"`)
 * @param attrRecord  Optional flat attribute map; `undefined` values are skipped
 * @param children  Optional pre-serialized child XML strings
 *
 * @example
 * ```ts
 * element("a:solidFill", undefined, [element("a:srgbClr", { val: "FF0000" })])
 * // => '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'
 * ```
 */
export function element(
  name: string,
  attrRecord?: Readonly<Record<string, string | number | boolean | undefined>>,
  children?: readonly string[],
): string {
  const attrStr = attrRecord ? attrs(attrRecord) : undefined;
  if (!children || children.length === 0) return selfCloseElement(name, attrStr);
  const body = children.join("");
  return body.length === 0
    ? selfCloseElement(name, attrStr)
    : `<${name}${attrStr ?? ""}>${body}</${name}>`;
}
