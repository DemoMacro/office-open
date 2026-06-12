/** Escape text content for XML. Fast path returns original string when no special chars. */
export function escapeXml(str: string): string {
  // Fast path: most text content doesn't contain XML-special characters.
  // Manual scan avoids regex overhead; returning the original string reference
  // means zero allocation for the common case.
  for (let i = 0; i < str.length; i++) {
    const c = str.charCodeAt(i);
    if (c === 38 || c === 34 || c === 39 || c === 60 || c === 62) {
      // & " ' < >
      // Slow path: slice-and-append avoids regex + temporary match objects.
      let s = "";
      let last = 0;
      for (let j = i; j < str.length; j++) {
        const cj = str.charCodeAt(j);
        if (cj === 38) {
          s += str.slice(last, j) + "&amp;";
          last = j + 1;
        } else if (cj === 34) {
          s += str.slice(last, j) + "&quot;";
          last = j + 1;
        } else if (cj === 39) {
          s += str.slice(last, j) + "&apos;";
          last = j + 1;
        } else if (cj === 60) {
          s += str.slice(last, j) + "&lt;";
          last = j + 1;
        } else if (cj === 62) {
          s += str.slice(last, j) + "&gt;";
          last = j + 1;
        }
      }
      return s + str.slice(last);
    }
  }
  return str;
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
  const parts: string[] = [];
  const keys = Object.keys(record);
  for (let i = 0; i < keys.length; i++) {
    const v = record[keys[i]];
    if (v !== undefined) {
      parts.push(` ${keys[i]}="${typeof v === "string" ? escapeXml(v) : v}"`);
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
