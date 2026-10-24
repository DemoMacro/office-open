import { assertValidXmlCharacters, isValidXmlCharacter } from "./characters";
import { XmlParseError } from "./errors";
import type { Element, ParseOptions } from "./types";

const ENTITY_MAP: Record<string, string> = {
  "&amp;": "&",
  "&lt;": "<",
  "&gt;": ">",
  "&quot;": '"',
  "&apos;": "'",
};
// Matches the five named entities plus numeric character references
// (&#65; decimal, &#x42; hex). The final alternative catches malformed and
// unsupported references so callers never receive their source text.
const ENTITY_PATTERN = /&(?:amp|lt|gt|quot|apos);|&#x[0-9a-fA-F]+;|&#[0-9]+;|&/g;
// "Contains a non-whitespace character" — same predicate as
// `text.trim().length > 0` without allocating the trimmed copy.
const HAS_CONTENT = /\S/;

export function unescapeXml(str: string): string {
  // Fast path: entities all start with '&', and OOXML parts overwhelmingly
  // contain none (a 63 MB worksheet measured zero occurrences). The regex
  // scan + replace setup per call showed up to ~12% of large-file parse
  // profiles, so gate it on the sentinel byte.
  if (str.indexOf("&") === -1) return str;
  return str.replace(ENTITY_PATTERN, (match) => {
    if (match === "&") throw new XmlParseError("Malformed XML: unsupported or malformed entity.");
    if (ENTITY_MAP[match] !== undefined) return ENTITY_MAP[match];
    // Numeric character reference: strip "&#" prefix and ";" suffix.
    const body = match.slice(2, -1);
    const code =
      body[0] === "x" || body[0] === "X" ? parseInt(body.slice(1), 16) : parseInt(body, 10);
    if (!Number.isSafeInteger(code) || code < 0 || !isValidXmlCharacter(code)) {
      throw new XmlParseError(`Malformed XML: invalid character reference ${match}.`);
    }
    return String.fromCodePoint(code);
  });
}

export function nativeTypeValue(value: string): string | number | boolean {
  if (value === "") return value;
  // Digit-only fast path — plain integers are the most common numeric shape
  // in OOXML (row/column indexes, sizes, ids). At most 15 digits is always
  // exact in float64, so the scan replaces the Number() + String(n)
  // round-trip (the String(n) side allocates) without re-checking losslessness.
  const neg = value.charCodeAt(0) === 0x2d; /* - */
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

interface XmlNamespaceLayer {
  prefixes: Map<string, string> | undefined;
  defaultCanonical: string | undefined;
}

function canonicalXmlName(
  name: string,
  prefixes: Map<string, string> | undefined,
  defaultCanonical: string | undefined,
): string {
  const colon = name.indexOf(":");
  if (colon > 0) {
    const prefix = name.slice(0, colon);
    if (prefix === "xml") return name;
    const canonical = prefixes?.get(prefix);
    if (canonical === undefined) return name;
    const local = name.slice(colon + 1);
    return canonical === "" ? local : canonical + ":" + local;
  }
  if (defaultCanonical !== undefined && defaultCanonical !== "")
    return defaultCanonical + ":" + name;
  return name;
}

function setXmlAttribute(
  attrs: Record<string, string> | undefined,
  name: string,
  value: string,
): Record<string, string> {
  if (attrs !== undefined && Object.hasOwn(attrs, name)) {
    throw new XmlParseError(`Malformed XML: duplicate attribute ${name}.`);
  }
  const result = attrs ?? {};
  result[name] = value;
  return result;
}

function claimCanonicalNamespace(
  namespaces: Map<string, string>,
  canonicalPrefix: string,
  namespace: string,
): void {
  const owner = namespaces.get(canonicalPrefix);
  if (owner !== undefined && owner !== namespace) {
    const label = canonicalPrefix === "" ? "default namespace" : `prefix ${canonicalPrefix}`;
    throw new XmlParseError(`Malformed XML: normalized ${label} has conflicting namespaces.`);
  }
  namespaces.set(canonicalPrefix, namespace);
}

function setNamespaceDeclaration(
  attrs: Record<string, string> | undefined,
  canonicalPrefix: string,
  namespace: string,
): Record<string, string> {
  const name = canonicalPrefix === "" ? "xmlns" : `xmlns:${canonicalPrefix}`;
  const existing = attrs?.[name];
  if (existing !== undefined) {
    if (existing !== namespace) {
      throw new XmlParseError(`Malformed XML: duplicate attribute ${name}.`);
    }
    return attrs ?? {};
  }
  return setXmlAttribute(attrs, name, namespace);
}

interface DeferredClose {
  closeIndex: number;
  nextIndex: number;
}

function openingTagEnd(source: string, from: number): number {
  let index = from;
  while (index < source.length) {
    const char = source.charCodeAt(index);
    if (char === 0x22 || char === 0x27) {
      const close = source.indexOf(char === 0x22 ? '"' : "'", index + 1);
      if (close === -1) return -1;
      index = close + 1;
      continue;
    }
    if (char === 0x3e) return index + 1;
    index++;
  }
  return -1;
}

function markupEnd(source: string, from: number, terminator: string): number {
  const end = source.indexOf(terminator, from);
  return end === -1 ? -1 : end + terminator.length;
}

/** Index of the `>` closing a DOCTYPE, including a quoted internal subset. */
function doctypeEnd(source: string, from: number): number {
  const bracket = source.indexOf("[", from);
  const simple = source.indexOf(">", from);
  if (bracket === -1 || (simple !== -1 && simple < bracket)) return simple;

  let cursor = bracket + 1;
  while (cursor < source.length) {
    if (source.startsWith("<!--", cursor)) {
      const end = source.indexOf("-->", cursor + 4);
      if (end === -1) return -1;
      cursor = end + 3;
      continue;
    }
    if (source.startsWith("<?", cursor)) {
      const end = source.indexOf("?>", cursor + 2);
      if (end === -1) return -1;
      cursor = end + 2;
      continue;
    }
    const quote = source.charCodeAt(cursor);
    if (quote === 0x22 || quote === 0x27) {
      const close = source.indexOf(quote === 0x22 ? '"' : "'", cursor + 1);
      if (close === -1) return -1;
      cursor = close + 1;
      continue;
    }
    if (source.charCodeAt(cursor) === 0x5d /* ] */) {
      return source.charCodeAt(cursor + 1) === 0x3e ? cursor + 1 : -1;
    }
    cursor++;
  }
  return -1;
}

function findDeferredClose(source: string, tagName: string, from: number): DeferredClose {
  let depth = 1;
  let cursor = from;
  while (cursor < source.length) {
    const openIndex = source.indexOf("<", cursor);
    if (openIndex === -1) break;

    if (source.startsWith("<![CDATA[", openIndex)) {
      cursor = markupEnd(source, openIndex + 9, "]]>");
      if (cursor === -1) throw new XmlParseError("Malformed XML: unclosed CDATA section.");
      continue;
    }
    if (source.startsWith("<!--", openIndex)) {
      cursor = markupEnd(source, openIndex + 4, "-->");
      if (cursor === -1) throw new XmlParseError("Malformed XML: unclosed comment.");
      continue;
    }
    if (source.startsWith("<?", openIndex)) {
      cursor = markupEnd(source, openIndex + 2, "?>");
      if (cursor === -1) throw new XmlParseError("Malformed XML: unclosed processing instruction.");
      continue;
    }
    if (source.startsWith("<!DOCTYPE", openIndex)) {
      const end = doctypeEnd(source, openIndex + 9);
      if (end === -1) throw new XmlParseError("Malformed XML: unclosed DOCTYPE declaration.");
      cursor = end + 1;
      continue;
    }

    const isClosing = source.charCodeAt(openIndex + 1) === 0x2f;
    const nameStart = openIndex + (isClosing ? 2 : 1);
    const nameEnd = findTagNameEnd(source, nameStart);
    const name = source.slice(nameStart, nameEnd);
    const tagEnd = isClosing ? nameEnd : openingTagEnd(source, nameEnd);
    if (tagEnd === -1) throw new XmlParseError(`Malformed XML: unclosed tag <${name}>.`);
    if (isClosing) {
      let afterName = nameEnd;
      while (afterName < source.length && /\s/.exec(source[afterName]!)) afterName++;
      if (source.charCodeAt(afterName) !== 0x3e) {
        throw new XmlParseError(`Malformed XML: invalid closing tag </${name}>.`);
      }
      if (name === tagName) {
        depth--;
        if (depth === 0) return { closeIndex: openIndex, nextIndex: afterName + 1 };
      }
      cursor = afterName + 1;
      continue;
    }

    const selfClosing = source.charCodeAt(tagEnd - 2) === 0x2f;
    if (name === tagName && !selfClosing) depth++;
    cursor = tagEnd;
  }
  throw new XmlParseError(`Malformed XML: unclosed element <${tagName}>.`);
}

export function parse(xmlString: string, options?: ParseOptions): Element {
  assertValidXmlCharacters(xmlString);
  const captureSpaces = options?.captureSpacesBetweenElements ?? false;
  const trim = options?.trim ?? false;
  const ignoreDeclaration = options?.ignoreDeclaration ?? false;
  const ignoreText = options?.ignoreText ?? false;
  const ignoreComment = options?.ignoreComment ?? false;
  const ignoreCdata = options?.ignoreCdata ?? false;
  const ignoreDoctype = options?.ignoreDoctype ?? false;
  const ignoreInstruction = options?.ignoreInstruction ?? false;
  const ignoreAttributes = options?.ignoreAttributes ?? false;
  const nativeTypeAttributes = options?.nativeTypeAttributes ?? false;
  // Lookup set for deferred elements (raw inner-XML capture). Undefined when
  // the option is absent so the common path pays one truthiness check.
  const deferSet =
    options?.deferElements !== undefined && options.deferElements.length > 0
      ? new Set(options.deferElements)
      : undefined;

  // Namespace normalization (see ParseOptions.normalizeNamespaces): in-scope
  // prefix → canonical-prefix bindings, copy-on-write per element that carries
  // xmlns declarations. `prefixes === undefined` means nothing is bound yet
  // (the common subtree shares the root layer object untouched).
  const nsTable =
    options?.normalizeNamespaces !== undefined
      ? new Map(Object.entries(options.normalizeNamespaces))
      : undefined;
  const nsRootLayer: XmlNamespaceLayer = { prefixes: undefined, defaultCanonical: undefined };
  const nsStack: XmlNamespaceLayer[] = [nsRootLayer];

  const result: Element = {};
  const stack: Element[] = [result];
  const openNames: string[] = [];

  let i = 0;
  const len = xmlString.length;

  while (i < len) {
    // Text node: read up to the next '<'. Pure-whitespace nodes (indentation)
    // are dropped below unless captureSpaces is on, but leading/trailing
    // spaces of nodes that have content are preserved.
    if (xmlString.charCodeAt(i) !== 0x3c /* < */) {
      const start = i;
      // Single-char indexOf runs as a native memchr scan — far faster than a
      // JS byte-by-byte loop (the outer check already guarantees i is not '<').
      const lt = xmlString.indexOf("<", i);
      i = lt === -1 ? len : lt;
      let text = unescapeXml(xmlString.slice(start, i));
      if (trim) text = text.trim();
      if (ignoreText) continue;
      if (text.length > 0) {
        // trim mode already guarantees non-whitespace past the length gate;
        // otherwise test natively instead of allocating a trimmed copy.
        if (captureSpaces || (trim ? true : HAS_CONTENT.test(text)) || isPreserveContext(stack)) {
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
      if (end === -1) throw new XmlParseError("Malformed XML: unclosed processing instruction.");
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
      } else if (!ignoreInstruction) {
        const parent = peek(stack);
        if (parent.elements === undefined) parent.elements = [];
        parent.elements.push({ type: "instruction", instruction: body });
      }
      continue;
    }

    // !-- comment
    if (xmlString.charCodeAt(i) === 0x21 && xmlString.slice(i, i + 3) === "!--") {
      const end = xmlString.indexOf("-->", i + 3);
      if (end === -1) throw new XmlParseError("Malformed XML: unclosed comment.");
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
      if (end === -1) throw new XmlParseError("Malformed XML: unclosed CDATA section.");
      const cdata = xmlString.slice(i + 8, end);
      i = end + 3;
      if (!ignoreCdata) {
        if (trim) addField(peek(stack), "cdata", cdata.trim());
        else addField(peek(stack), "cdata", cdata);
      }
      continue;
    }

    // <!DOCTYPE
    if (xmlString.charCodeAt(i) === 0x21 && xmlString.slice(i, i + 8) === "!DOCTYPE") {
      const end = doctypeEnd(xmlString, i + 9);
      if (end === -1) throw new XmlParseError("Malformed XML: unclosed DOCTYPE declaration.");
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
      if (end === -1) throw new XmlParseError("Malformed XML: unclosed closing tag.");
      const closingName = xmlString.slice(i + 1, end).trim();
      if (closingName.length === 0) throw new XmlParseError("Malformed XML: empty closing tag.");
      const closingLayer = nsStack[nsStack.length - 1]!;
      const closingFinalName = nsTable
        ? canonicalXmlName(closingName, closingLayer.prefixes, closingLayer.defaultCanonical)
        : closingName;
      const openName = openNames.pop();
      if (openName === undefined || openName !== closingFinalName) {
        throw new XmlParseError(
          `Malformed XML: expected closing tag </${openName ?? "root"}>, got </${closingFinalName}>.`,
        );
      }
      i = end + 1;
      stack.pop();
      if (nsTable !== undefined) nsStack.pop();
      continue;
    }

    // < opening tag
    const tagNameEnd = findTagNameEnd(xmlString, i);
    const tagName = xmlString.slice(i, tagNameEnd);
    let pos = tagNameEnd;

    // Namespace bindings for this element: inherit the parent layer; the
    // attribute scan clones it on the first xmlns declaration it meets.
    const nsParent = nsTable !== undefined ? nsStack[nsStack.length - 1]! : undefined;
    let nsPrefixes = nsParent?.prefixes;
    let nsDefault = nsParent?.defaultCanonical;
    let canonicalNamespaces: Map<string, string> | undefined;
    let nsDeclared = false;

    // Attribute scan, inlined from parseAttributesFromXml: called once per
    // opening tag, the `{ attrs, pos }` wrapper was one allocation per element.
    // `attrs` is allocated lazily on the first attribute — tags without
    // attributes (the majority in data-heavy parts) must not pay a record
    // allocation.
    let attrs: Record<string, string> | undefined;
    while (pos < len) {
      while (pos < len && isWhitespace(xmlString.charCodeAt(pos))) pos++;
      if (pos >= len || xmlString.charCodeAt(pos) === 0x3e || xmlString.charCodeAt(pos) === 0x2f) {
        break;
      }

      const nameStart = pos;
      while (pos < len && xmlString.charCodeAt(pos) !== 0x3d) {
        if (xmlString.charCodeAt(pos) === 0x3e || xmlString.charCodeAt(pos) === 0x2f) break;
        pos++;
      }
      // XML's Eq production allows whitespace around '='; trim it off the
      // name so `Target = "…"` resolves like `Target="…"`.
      const name = xmlString.slice(nameStart, pos).trim();

      while (pos < len && isWhitespace(xmlString.charCodeAt(pos))) pos++;
      const hasEquals = xmlString.charCodeAt(pos) === 0x3d;
      if (!hasEquals) {
        if (name.length > 0) {
          throw new XmlParseError(`Malformed XML: attribute ${name} is missing a value.`);
        }
        break;
      }
      pos++;

      while (pos < len && isWhitespace(xmlString.charCodeAt(pos))) pos++;

      const quote = xmlString.charCodeAt(pos);
      if (quote !== 0x22 && quote !== 0x27) {
        throw new XmlParseError(`Malformed XML: attribute ${name} is missing a quoted value.`);
      }
      pos++;
      const valueStart = pos;
      while (pos < len && xmlString.charCodeAt(pos) !== quote) pos++;
      const value = unescapeXml(xmlString.slice(valueStart, pos));
      pos++;

      if (nsTable !== undefined && (name === "xmlns" || name.startsWith("xmlns:"))) {
        const canonical = nsTable.get(value);
        const prefix = name === "xmlns" ? "" : name.slice(6);
        const canonicalPrefix = canonical ?? prefix;
        canonicalNamespaces ??= new Map();
        claimCanonicalNamespace(canonicalNamespaces, canonicalPrefix, value);
        if (!nsDeclared) {
          nsPrefixes = new Map(nsPrefixes ?? []);
          nsDeclared = true;
        }
        if (name === "xmlns") {
          nsDefault = canonicalPrefix;
        } else {
          nsPrefixes!.set(prefix, canonicalPrefix);
        }
        attrs = setNamespaceDeclaration(attrs, canonicalPrefix, value);
        continue;
      }
      attrs = setXmlAttribute(attrs, name, value);
    }
    if (pos >= len) throw new XmlParseError(`Malformed XML: unclosed element <${tagName}>.`);
    const terminal = xmlString.charCodeAt(pos);
    if (terminal !== 0x3e && terminal !== 0x2f) {
      throw new XmlParseError("Malformed XML: malformed opening tag.");
    }

    let rawValues: Record<string, string> | undefined;
    if (attrs && nativeTypeAttributes) {
      for (const key in attrs) {
        const raw = attrs[key] as string;
        const value = nativeTypeValue(raw);
        attrs[key] = value as string;
        if (typeof value === "boolean") {
          rawValues ??= {};
          rawValues[key] = raw;
        }
      }
    }

    // Namespace-normalize the element name and the non-declaration attribute
    // names onto this element's effective bindings (its own declarations
    // included — attributes see the element's in-scope bindings).
    let finalName = tagName;
    if (nsTable !== undefined) {
      finalName = canonicalXmlName(tagName, nsPrefixes, nsDefault);
      if (attrs !== undefined && nsPrefixes !== undefined) {
        for (const key of Object.keys(attrs)) {
          if (key === "xmlns" || key.startsWith("xmlns:")) continue;
          const kcolon = key.indexOf(":");
          if (kcolon <= 0) continue;
          const p = key.slice(0, kcolon);
          if (p === "xml") continue;
          const c = nsPrefixes.get(p);
          if (c !== undefined && c !== p) {
            const renamed = c === "" ? key.slice(kcolon + 1) : c + key.slice(kcolon);
            if (renamed !== key && Object.hasOwn(attrs, renamed)) {
              throw new XmlParseError(`Malformed XML: duplicate attribute ${renamed}.`);
            }
            const attrValue = attrs[key];
            if (attrValue !== undefined) attrs[renamed] = attrValue;
            delete attrs[key];
            if (rawValues !== undefined && Object.hasOwn(rawValues, key)) {
              const rawValue = rawValues[key]!;
              delete rawValues[key];
              rawValues[renamed] = rawValue;
            }
          }
        }
      }
    }

    const isSelfClosing = xmlString.charCodeAt(pos) === 0x2f; /* / */
    if (isSelfClosing) pos += 2;
    else pos++;

    // Constant shape: every element node carries all four fields (undefined
    // placeholders included), so attribute/children assignment writes an
    // existing slot instead of walking a hidden-class transition — element
    // access sites stay monomorphic instead of 2-3 shapes deep.
    const element: Element = {
      type: "element",
      name: finalName,
      attributes: ignoreAttributes ? undefined : attrs,
      elements: undefined,
    };
    if (rawValues) element.attributeRawValues = rawValues;

    const parent = peek(stack);
    if (!parent.elements) {
      parent.elements = [];
    }
    parent.elements.push(element);

    if (!isSelfClosing) {
      if (deferSet !== undefined && deferSet.has(tagName)) {
        // Deferred content can contain CDATA whose text names this element or
        // quoted attribute values that look like tags.
        const close = findDeferredClose(xmlString, tagName, pos);
        element.raw = xmlString.slice(pos, close.closeIndex);
        i = close.nextIndex;
        continue;
      }
      stack.push(element);
      openNames.push(finalName);
      if (nsTable !== undefined) {
        // Share the parent layer unless this element declared anything (a
        // cloned prefix map, or a default declaration that can shadow one).
        const layerChanged = nsDeclared || nsDefault !== nsParent!.defaultCanonical;
        nsStack.push(
          layerChanged
            ? {
                prefixes: nsPrefixes,
                defaultCanonical: nsDefault,
              }
            : nsParent!,
        );
      }
    }

    i = pos;
  }

  if (result.elements) {
    const temp = result.elements;
    delete result.elements;
    result.elements = temp;
    delete result.text;
  }

  if (openNames.length > 0) {
    throw new XmlParseError(
      `Malformed XML: unclosed element <${openNames[openNames.length - 1]!}>.`,
    );
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
