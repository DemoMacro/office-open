import type { Element, ParseOptions } from "./types";

const ENTITY_MAP: Record<string, string> = {
  "&amp;": "&",
  "&lt;": "<",
  "&gt;": ">",
  "&quot;": '"',
  "&apos;": "'",
};
// Matches the five named entities plus numeric character references
// (&#65; decimal, &#x42; hex).
const ENTITY_PATTERN = /&(?:amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#[0-9]+);/g;

export function unescapeXml(str: string): string {
  // Fast path: entities all start with '&', and OOXML parts overwhelmingly
  // contain none (a 63 MB worksheet measured zero occurrences). The regex
  // scan + replace setup per call showed up to ~12% of large-file parse
  // profiles, so gate it on the sentinel byte.
  if (str.indexOf("&") === -1) return str;
  return str.replace(ENTITY_PATTERN, (match) => {
    if (ENTITY_MAP[match] !== undefined) return ENTITY_MAP[match];
    // Numeric character reference: strip "&#" prefix and ";" suffix.
    const body = match.slice(2, -1);
    const code =
      body[0] === "x" || body[0] === "X" ? parseInt(body.slice(1), 16) : parseInt(body, 10);
    return Number.isFinite(code) && code >= 0 ? String.fromCodePoint(code) : match;
  });
}

export function nativeTypeValue(value: string): string | number | boolean {
  if (value === "") return value;
  // Digit-only fast path — plain integers are the most common numeric shape
  // in OOXML (row/column indexes, sizes, ids). At most 15 digits is always
  // exact in float64, so the scan replaces the Number() + String(n)
  // round-trip (the String(n) side allocates) without re-checking losslessness.
  const neg = value.charCodeAt(0) === 0x2d /* - */;
  const start = neg ? 1 : 0;
  const digits = value.length - start;
  if (digits > 0 && digits <= 15) {
    // Leading zeros ("00992297") must stay strings; only a lone "0" passes.
    // "-0" falls through too — Number coerces it to -0 whose String() is "0",
    // so the slow path must keep it a string.
    const head = value.charCodeAt(start);
    if (head !== 0x30 || (!neg && digits === 1)) {
      let n = 0;
      let allDigits = true;
      for (let i = start; i < value.length; i++) {
        const c = value.charCodeAt(i);
        if (c < 0x30 || c > 0x39) {
          allDigits = false;
          break;
        }
        n = n * 10 + (c - 0x30);
      }
      if (allDigits) return neg ? -n : n;
    }
  }
  const n = Number(value);
  // Only coerce when lossless: leading zeros ("00992297"), exponential
  // notation ("1e5"), and a leading sign ("+5") must stay strings so hex-like
  // values (rsid, color) survive parse → stringify round-trips byte-exact.
  if (!isNaN(n) && String(n) === value) return n;
  // Length gate before toLowerCase: every non-numeric attribute value paid
  // two throwaway lowercase strings (cell refs, format names, ids …), which
  // parse profiles attributed to GC pressure.
  if (value.length === 4 || value.length === 5) {
    const lower = value.toLowerCase();
    if (lower === "true") return true;
    if (lower === "false") return false;
  }
  return value;
}

export function parse(xmlString: string, options?: ParseOptions): Element {
  const captureSpaces = options?.captureSpacesBetweenElements ?? false;
  const trim = options?.trim ?? false;
  const ignoreDeclaration = options?.ignoreDeclaration ?? false;
  const ignoreText = options?.ignoreText ?? false;
  const ignoreComment = options?.ignoreComment ?? false;
  const ignoreCdata = options?.ignoreCdata ?? false;
  const ignoreDoctype = options?.ignoreDoctype ?? false;
  const nativeTypeAttributes = options?.nativeTypeAttributes ?? false;
  // Lookup set for deferred elements (raw inner-XML capture). Undefined when
  // the option is absent so the common path pays one truthiness check.
  const deferSet =
    options?.deferElements !== undefined && options.deferElements.length > 0
      ? new Set(options.deferElements)
      : undefined;

  const result: Element = {};
  const stack: Element[] = [result];

  let i = 0;
  const len = xmlString.length;

  while (i < len) {
    // Text node: read up to the next '<'. Pure-whitespace nodes (indentation)
    // are dropped below unless captureSpaces is on, but leading/trailing
    // spaces of nodes that have content are preserved.
    if (xmlString.charCodeAt(i) !== 0x3c /* < */) {
      const start = i;
      while (i < len && xmlString.charCodeAt(i) !== 0x3c) i++;
      let text = unescapeXml(xmlString.slice(start, i));
      if (trim) text = text.trim();
      if (ignoreText) continue;
      if (text.length > 0) {
        if (captureSpaces || text.trim().length > 0 || isPreserveContext(stack)) {
          // Text-node hot path, inlined from addField("text"): one lookup of
          // the last child covers both the adjacent-merge case (same shape as
          // addField — a split CDATA/text run must reassemble) and the fresh
          // push. addField stays for the cold node types.
          const parent = stack[stack.length - 1]!;
          const elements = parent.elements;
          const last = elements === undefined ? undefined : elements[elements.length - 1];
          if (last !== undefined && last.type === "text") {
            last.text = (last.text as string) + text;
          } else {
            const node: Element = { type: "text", text };
            if (elements === undefined) {
              parent.elements = [node];
            } else {
              elements.push(node);
            }
          }
        }
      }
      continue;
    }

    i++;

    // <? processing instruction / declaration
    if (xmlString.charCodeAt(i) === 0x3f /* ? */) {
      const end = xmlString.indexOf("?>", i + 1);
      if (end === -1) break;
      const body = xmlString.slice(i + 1, end);
      i = end + 2;

      const xmlMatch = body.match(/^xml\s+(.*)$/s);
      if (xmlMatch) {
        if (!ignoreDeclaration) {
          if (!result.declaration) {
            result.declaration = {};
          }
          const attrs = parseAttributes(xmlMatch[1] ?? "");
          if (nativeTypeAttributes) {
            for (const key in attrs) {
              attrs[key] = nativeTypeValue(attrs[key] as string) as string;
            }
          }
          result.declaration.attributes = attrs;
        }
      }
      continue;
    }

    // !-- comment
    if (xmlString.charCodeAt(i) === 0x21 && xmlString.slice(i, i + 3) === "!--") {
      const end = xmlString.indexOf("-->", i + 3);
      if (end === -1) break;
      const comment = xmlString.slice(i + 3, end);
      i = end + 3;
      if (!ignoreComment) {
        if (trim) addField(peek(stack), "comment", comment.trim());
        else addField(peek(stack), "comment", comment);
      }
      continue;
    }

    // ![CDATA[
    if (xmlString.charCodeAt(i) === 0x21 && xmlString.slice(i, i + 8) === "![CDATA[") {
      const end = xmlString.indexOf("]]>", i + 8);
      if (end === -1) break;
      const cdata = xmlString.slice(i + 8, end);
      i = end + 3;
      if (!ignoreCdata) {
        if (trim) addField(peek(stack), "cdata", cdata.trim());
        else addField(peek(stack), "cdata", cdata);
      }
      continue;
    }

    // <!DOCTYPE
    if (xmlString.charCodeAt(i) === 0x21 && xmlString.slice(i, i + 9) === "!DOCTYPE") {
      const end = xmlString.indexOf(">", i + 9);
      if (end === -1) break;
      const doctype = xmlString.slice(i + 9, end).trim();
      i = end + 1;
      if (!ignoreDoctype) {
        addField(peek(stack), "doctype", doctype);
      }
      continue;
    }

    // </ closing tag
    if (xmlString.charCodeAt(i) === 0x2f /* / */) {
      const end = xmlString.indexOf(">", i + 1);
      if (end === -1) break;
      i = end + 1;
      stack.pop();
      continue;
    }

    // < opening tag
    const tagNameEnd = findTagNameEnd(xmlString, i);
    const tagName = xmlString.slice(i, tagNameEnd);
    let pos = tagNameEnd;

    const attributes = parseAttributesFromXml(xmlString, pos);
    pos = attributes.pos;

    if (attributes.attrs && nativeTypeAttributes) {
      for (const key in attributes.attrs) {
        attributes.attrs[key] = nativeTypeValue(attributes.attrs[key] as string) as string;
      }
    }

    const isSelfClosing = xmlString.charCodeAt(pos) === 0x2f /* / */;
    if (isSelfClosing) pos += 2;
    else pos++;

    const element: Element = {
      type: "element",
      name: tagName,
    };
    if (attributes.attrs) {
      element.attributes = attributes.attrs;
    }

    const parent = peek(stack);
    if (!parent.elements) {
      parent.elements = [];
    }
    parent.elements.push(element);

    if (!isSelfClosing) {
      if (deferSet !== undefined && deferSet.has(tagName)) {
        // Deferred container: capture inner XML verbatim instead of parsing
        // children. Scan to the matching close tag, counting same-name opens
        // so nested occurrences (if any) don't end the capture early.
        const closeTag = `</${tagName}>`;
        let depth = 1;
        let scan = pos;
        let closeIdx = -1;
        for (;;) {
          closeIdx = xmlString.indexOf(closeTag, scan);
          if (closeIdx === -1) break;
          let p = scan;
          for (;;) {
            const openIdx = xmlString.indexOf(`<${tagName}`, p);
            if (openIdx === -1 || openIdx >= closeIdx) break;
            // Boundary check so `<rowx>` doesn't count as `<row`.
            const after = xmlString.charCodeAt(openIdx + tagName.length + 1);
            if (
              after === 0x20 ||
              after === 0x09 ||
              after === 0x0a ||
              after === 0x0d ||
              after === 0x2f ||
              after === 0x3e
            ) {
              depth++;
            }
            p = openIdx + tagName.length + 1;
          }
          scan = closeIdx + closeTag.length;
          depth--;
          if (depth === 0) break;
        }
        if (closeIdx === -1) {
          element.raw = xmlString.slice(pos);
          i = len;
        } else {
          element.raw = xmlString.slice(pos, closeIdx);
          i = scan;
        }
        continue;
      }
      stack.push(element);
    }

    i = pos;
  }

  if (result.elements) {
    const temp = result.elements;
    delete result.elements;
    result.elements = temp;
    delete result.text;
  }

  return result;
}

function findTagNameEnd(str: string, start: number): number {
  let i = start;
  const len = str.length;
  while (i < len) {
    const ch = str.charCodeAt(i);
    if (ch === 0x20 || ch === 0x09 || ch === 0x0a || ch === 0x0d || ch === 0x2f || ch === 0x3e) {
      return i;
    }
    i++;
  }
  return i;
}

function parseAttributesFromXml(
  str: string,
  start: number,
): { attrs: Record<string, string> | undefined; pos: number } {
  // `attrs` is allocated lazily on the first attribute — tags without
  // attributes (the majority in data-heavy parts) must not pay a record
  // allocation, and the caller checks `attrs !== undefined` instead of
  // Object.keys() (which allocated a throwaway key array per tag).
  let attrs: Record<string, string> | undefined;
  let i = start;
  const len = str.length;

  while (i < len) {
    while (i < len && isWhitespace(str.charCodeAt(i))) i++;
    if (i >= len || str.charCodeAt(i) === 0x3e || str.charCodeAt(i) === 0x2f) {
      break;
    }

    const nameStart = i;
    while (i < len && str.charCodeAt(i) !== 0x3d) {
      if (str.charCodeAt(i) === 0x3e || str.charCodeAt(i) === 0x2f) break;
      i++;
    }
    const name = str.slice(nameStart, i);

    if (str.charCodeAt(i) !== 0x3d) break;
    i++;

    while (i < len && isWhitespace(str.charCodeAt(i))) i++;

    const quote = str.charCodeAt(i);
    if (quote !== 0x22 && quote !== 0x27) break;
    i++;
    const valueStart = i;
    while (i < len && str.charCodeAt(i) !== quote) i++;
    if (attrs === undefined) attrs = {};
    attrs[name] = unescapeXml(str.slice(valueStart, i));
    i++;
  }

  return { attrs, pos: i };
}

export function parseAttributes(str: string): Record<string, string> {
  const result: Record<string, string> = {};
  let i = 0;
  const len = str.length;

  while (i < len) {
    while (i < len && isWhitespace(str.charCodeAt(i))) i++;
    if (i >= len) break;

    const nameStart = i;
    while (i < len && str.charCodeAt(i) !== 0x3d) {
      if (isWhitespace(str.charCodeAt(i))) break;
      i++;
    }
    const name = str.slice(nameStart, i);

    while (i < len && isWhitespace(str.charCodeAt(i))) i++;
    if (i >= len || str.charCodeAt(i) !== 0x3d) break;
    i++;

    while (i < len && isWhitespace(str.charCodeAt(i))) i++;

    const quote = str.charCodeAt(i);
    if (quote !== 0x22 && quote !== 0x27) break;
    i++;
    const valueStart = i;
    while (i < len && str.charCodeAt(i) !== quote) i++;
    result[name] = unescapeXml(str.slice(valueStart, i));
    i++;
  }
  return result;
}

/**
 * Top of the parse stack. The stack is guaranteed non-empty — the result root
 * is pushed at init and push/pop stay balanced across well-formed input — so
 * this is a compile-time narrow (one non-null assertion) rather than a runtime
 * check: it must not add a throw path that changes how `parse` surfaces
 * malformed documents. Centralising the access keeps that single `!` off the
 * read sites, matching the "wrap indexed access behind a helper" pattern.
 */
function peek(stack: Element[]): Element {
  return stack[stack.length - 1]!;
}

function addField(parent: Element, type: string, value: string) {
  if (!parent.elements) {
    parent.elements = [];
  }
  // Merge adjacent text/cdata nodes: a CDATA section containing the literal
  // `]]>` is serialized as two adjacent CDATA sections and must reassemble
  // into a single node on parse. Adjacent text nodes likewise merge.
  if (type === "text" || type === "cdata") {
    const last = parent.elements[parent.elements.length - 1];
    if (last && last.type === type) {
      const key = type as "text" | "cdata";
      last[key] = (last[key] as string) + value;
      return;
    }
  }
  const element: Element = { type };
  (element as Record<string, unknown>)[type] = value;
  parent.elements.push(element);
}

/** True when the nearest ancestor with an explicit xml:space sets "preserve". */
function isPreserveContext(stack: Element[]): boolean {
  for (let i = stack.length - 1; i >= 0; i--) {
    const node = stack[i];
    if (!node) continue;
    const space = node.attributes?.["xml:space"];
    if (space !== undefined) return space === "preserve";
  }
  return false;
}

function isWhitespace(ch: number): boolean {
  return ch === 0x20 || ch === 0x09 || ch === 0x0a || ch === 0x0d;
}
