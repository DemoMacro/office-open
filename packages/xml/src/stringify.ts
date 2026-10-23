import { escapeXml } from "./escape";
import type { Element, StringifyOptions } from "./types";

// Non-global on purpose: a /g regex would carry stateful lastIndex across calls.
// Text content needs only the three markup delimiters — quotes are legal as-is.
const TEXT_SPECIALS = /[&<>]/;

export function stringify(js: Element, options?: StringifyOptions): string {
  const opts = normalizeOptions(options);
  const parts: string[] = [];

  if (js.declaration && !opts.ignoreDeclaration) {
    parts.push(writeDeclaration(js.declaration));
  }
  if (js.instruction !== undefined && !opts.ignoreInstruction) {
    assertInstruction(js.instruction);
    parts.push(`<?${js.instruction}?>`);
  }

  if (js.elements?.length) {
    parts.push(writeElements(js.elements, opts, 0, !parts.length));
  }

  return parts.join("");
}

interface NormalizedOptions {
  spaces: string;
  ignoreDeclaration: boolean;
  ignoreInstruction: boolean;
  ignoreAttributes: boolean;
  ignoreText: boolean;
  ignoreComment: boolean;
  ignoreCdata: boolean;
  ignoreDoctype: boolean;
  fullTagEmptyElement: boolean;
  indentText: boolean;
  indentCdata: boolean;
  attributeValueFn?: StringifyOptions["attributeValueFn"];
}

// Shared frozen defaults — the no-options call (the hot path) skips the
// per-call object construction entirely.
const DEFAULT_OPTIONS: NormalizedOptions = {
  spaces: "",
  ignoreDeclaration: false,
  ignoreInstruction: false,
  ignoreAttributes: false,
  ignoreText: false,
  ignoreComment: false,
  ignoreCdata: false,
  ignoreDoctype: false,
  fullTagEmptyElement: false,
  indentText: false,
  indentCdata: false,
};

function normalizeOptions(options?: StringifyOptions): NormalizedOptions {
  if (!options) return DEFAULT_OPTIONS;
  let spaces = "";
  if (options.spaces != null) {
    spaces = typeof options.spaces === "number" ? " ".repeat(options.spaces) : options.spaces;
  }
  return {
    spaces,
    ignoreDeclaration: options.ignoreDeclaration ?? false,
    ignoreInstruction: options.ignoreInstruction ?? false,
    ignoreAttributes: options.ignoreAttributes ?? false,
    ignoreText: options.ignoreText ?? false,
    ignoreComment: options.ignoreComment ?? false,
    ignoreCdata: options.ignoreCdata ?? false,
    ignoreDoctype: options.ignoreDoctype ?? false,
    fullTagEmptyElement: options.fullTagEmptyElement ?? false,
    indentText: options.indentText ?? false,
    indentCdata: options.indentCdata ?? false,
    attributeValueFn: options.attributeValueFn,
  };
}

function writeIndentation(spaces: string, depth: number, firstLine: boolean): string {
  if (!spaces) return "";
  return (!firstLine ? "\n" : "") + spaces.repeat(depth);
}

function writeDeclaration(declaration: NonNullable<Element["declaration"]>): string {
  const attrs = declaration.attributes;
  if (!attrs) return '<?xml version="1.0"?>';

  const version = attrs.version === undefined ? "1.0" : String(attrs.version);
  if (version !== "1.0" && version !== "1.1") {
    throw new Error(`Unsupported XML declaration version ${version}.`);
  }
  const parts: string[] = [`<?xml version="${version}"`];
  if (attrs.encoding) parts.push(` encoding="${attrs.encoding}"`);
  if (attrs.standalone) parts.push(` standalone="${attrs.standalone}"`);
  return parts.join("") + "?>";
}

function writeAttributes(
  attributes: Record<string, string | number | undefined>,
  elementName: string,
  element: Element,
  attributeValueFn?: StringifyOptions["attributeValueFn"],
): string {
  // Rope accumulation: attribute counts are small (1-3), a parts array + join
  // would cost more than V8's cons-string +=. for-in over plain data records
  // skips the Object.keys array allocation JSC penalizes (~1.3× on bun).
  let s = "";
  for (const key in attributes) {
    const value = attributes[key];
    if (value === null || value === undefined) continue;

    // attributeValueFn (xml-js hook) owns escaping when provided; otherwise
    // we escape all XML-special characters ourselves.
    const raw = String(value);
    const attr = attributeValueFn
      ? attributeValueFn(raw, key, elementName, element)
      : escapeXml(raw);
    s += ` ${key}="${attr}"`;
  }
  return s;
}

function assertInstruction(instruction: string): void {
  if (instruction.includes("?>")) {
    throw new Error("Malformed XML: processing instruction contains its closing delimiter.");
  }
}

function assertComment(comment: string): void {
  if (comment.includes("--")) {
    throw new Error("Malformed XML: comment contains the reserved double hyphen.");
  }
  if (comment.endsWith("-")) {
    throw new Error("Malformed XML: comment ends with a hyphen.");
  }
}

function writeElements(
  elements: Element[],
  opts: NormalizedOptions,
  depth: number,
  firstLine: boolean,
): string {
  // Rope accumulation — V8 cons-strings make += O(1) and skip the parts-array
  // allocation the join-based form pays per level. The element body is inlined
  // (no writeElement/writeElements mutual recursion): an ordered if-chain over
  // the node type benchmarks ~1.6× faster than a switch plus a helper call
  // per element under JSC, with V8 unchanged.
  let s = "";
  for (let i = 0; i < elements.length; i++) {
    const element = elements[i];
    if (!element) continue;
    const isFirst = firstLine && i === 0;
    const type = element.type;
    if (type === "element") {
      const name = element.name;
      if (!name) continue;
      if (opts.spaces) s += writeIndentation(opts.spaces, depth, isFirst);
      const attributes = element.attributes;
      const attrStr =
        attributes && !opts.ignoreAttributes
          ? writeAttributes(attributes, name, element, opts.attributeValueFn)
          : "";
      // Deferred content: re-emit the captured inner XML verbatim — children
      // were never parsed, and the bytes must survive a set/save round-trip.
      if (element.raw !== undefined) {
        validateRawXml(element.name!, element.raw);
        s += `<${name}${attrStr}>${element.raw}</${name}>`;
        continue;
      }
      const children = element.elements;
      // The xml:space dictionary probe is deliberately last: it only matters
      // for empty leaf elements, so child-bearing elements skip it.
      const withClosingTag =
        (children !== undefined && children.length > 0) ||
        opts.fullTagEmptyElement ||
        attributes?.["xml:space"] === "preserve";
      if (!withClosingTag) {
        s += `<${name}${attrStr}/>`;
        continue;
      }
      const open = `<${name}${attrStr}>`;
      if (children !== undefined && children.length > 0) {
        const inner = writeElements(children, opts, depth + 1, false);
        // The child-element scan is only needed to pretty-print the closing
        // tag — skip it entirely when not indenting.
        if (opts.spaces && children.some((e) => e.type === "element")) {
          s += open + inner + "\n" + opts.spaces.repeat(depth) + `</${name}>`;
        } else {
          s += open + inner + `</${name}>`;
        }
      } else {
        s += open + `</${name}>`;
      }
    } else if (type === "text") {
      if (opts.ignoreText) continue;
      if (opts.indentText && opts.spaces) s += writeIndentation(opts.spaces, depth, isFirst);
      // Text escaping inline — fast path: most text content contains no
      // markup delimiters, so a regex test (a native scan in V8/JSC, ~10× a
      // charCodeAt loop) returns the original string reference with zero
      // allocation. Chained replaces afterwards: specials are sparse, so each
      // pass after the first usually finds nothing and exits quickly.
      const text = element.text;
      if (text == null) continue;
      const str = String(text);
      s += TEXT_SPECIALS.test(str)
        ? str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
        : str;
    } else if (type === "instruction") {
      if (opts.ignoreInstruction) continue;
      if (opts.spaces) s += writeIndentation(opts.spaces, depth, isFirst);
      const instruction = element.instruction ?? "";
      assertInstruction(instruction);
      s += `<?${instruction}?>`;
    } else if (type === "cdata") {
      if (opts.ignoreCdata) continue;
      if (opts.indentCdata && opts.spaces) s += writeIndentation(opts.spaces, depth, isFirst);
      s += writeCdata(element.cdata);
    } else if (type === "comment") {
      if (opts.ignoreComment) continue;
      if (opts.spaces) s += writeIndentation(opts.spaces, depth, isFirst);
      s += writeComment(element.comment);
    } else if (type === "doctype") {
      if (opts.ignoreDoctype) continue;
      if (opts.spaces) s += writeIndentation(opts.spaces, depth, isFirst);
      s += writeDoctype(element.doctype);
    }
  }
  return s;
}

function writeCdata(cdata: string | undefined | null): string {
  if (cdata == null) return "";
  const escaped = cdata.replace(/\]\]>/g, "]]]]><![CDATA[>");
  return `<![CDATA[${escaped}]]>`;
}

function writeComment(comment: string | undefined | null): string {
  if (comment == null) return "";
  assertComment(comment);
  return `<!--${comment}-->`;
}

function validateRawXml(name: string, raw: string): void {
  const stack: string[] = [];
  let cursor = 0;
  while (cursor < raw.length) {
    const open = raw.indexOf("<", cursor);
    if (open === -1) break;
    if (raw.startsWith("<!--", open)) {
      const end = raw.indexOf("-->", open + 4);
      if (end === -1)
        throw new Error(`Malformed XML: raw content for <${name}> has an unclosed comment.`);
      cursor = end + 3;
      continue;
    }
    if (raw.startsWith("<![CDATA[", open)) {
      const end = raw.indexOf("]]>", open + 9);
      if (end === -1)
        throw new Error(`Malformed XML: raw content for <${name}> has an unclosed CDATA section.`);
      cursor = end + 3;
      continue;
    }
    if (raw.startsWith("<?", open)) {
      const end = raw.indexOf("?>", open + 2);
      if (end === -1)
        throw new Error(`Malformed XML: raw content for <${name}> has an unclosed instruction.`);
      cursor = end + 2;
      continue;
    }
    const isClosing = raw.charCodeAt(open + 1) === 0x2f;
    const nameStart = open + (isClosing ? 2 : 1);
    let nameEnd = nameStart;
    while (nameEnd < raw.length && !/\s|>|\//.exec(raw[nameEnd]!)) nameEnd++;
    const tag = raw.slice(nameStart, nameEnd);
    if (!tag) throw new Error(`Malformed XML: raw content for <${name}> has an empty tag name.`);
    let cursor2 = nameEnd;
    while (cursor2 < raw.length && raw.charCodeAt(cursor2) !== 0x3e) {
      const quote = raw.charCodeAt(cursor2);
      if (quote === 0x22 || quote === 0x27) {
        const close = raw.indexOf(quote === 0x22 ? '"' : "'", cursor2 + 1);
        if (close === -1)
          throw new Error(
            `Malformed XML: raw content for <${name}> has an unclosed attribute value.`,
          );
        cursor2 = close + 1;
        continue;
      }
      cursor2++;
    }
    if (cursor2 >= raw.length)
      throw new Error(`Malformed XML: raw content for <${name}> has an unclosed tag.`);
    let contentEnd = cursor2;
    while (contentEnd > 0 && isXmlWhitespace(raw.charCodeAt(contentEnd - 1))) contentEnd--;
    const selfClosing = raw.charCodeAt(contentEnd - 1) === 0x2f;
    cursor = cursor2 + 1;
    if (isClosing) {
      if (stack.pop() !== tag) {
        throw new Error(`Malformed XML: raw content for <${name}> has a mismatched closing tag.`);
      }
    } else if (!selfClosing) stack.push(tag);
  }
  if (stack.length > 0) {
    throw new Error(`Malformed XML: raw content for <${name}> has unclosed elements.`);
  }
}

function writeDoctype(doctype: string | undefined | null): string {
  if (doctype == null) return "";
  return `<!DOCTYPE ${doctype}>`;
}

function isXmlWhitespace(charCode: number): boolean {
  return charCode === 0x20 || charCode === 0x09 || charCode === 0x0a || charCode === 0x0d;
}

type NonNullable<T> = T extends null | undefined ? never : T;
