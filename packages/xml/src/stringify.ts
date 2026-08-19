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

  if (js.elements?.length) {
    parts.push(writeElements(js.elements, opts, 0, !parts.length));
  }

  return parts.join("");
}

function normalizeOptions(options?: StringifyOptions): {
  spaces: string;
  ignoreDeclaration: boolean;
  ignoreText: boolean;
  ignoreComment: boolean;
  ignoreCdata: boolean;
  ignoreDoctype: boolean;
  fullTagEmptyElement: boolean;
  indentText: boolean;
  indentCdata: boolean;
  attributeValueFn?: StringifyOptions["attributeValueFn"];
} {
  if (!options) {
    return {
      spaces: "",
      ignoreDeclaration: false,
      ignoreText: false,
      ignoreComment: false,
      ignoreCdata: false,
      ignoreDoctype: false,
      fullTagEmptyElement: false,
      indentText: false,
      indentCdata: false,
    };
  }
  let spaces = "";
  if (options.spaces != null) {
    spaces = typeof options.spaces === "number" ? " ".repeat(options.spaces) : options.spaces;
  }
  return {
    spaces,
    ignoreDeclaration: options.ignoreDeclaration ?? false,
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

  const parts: string[] = [`<?xml version="1.0"`];
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
  // would cost more than V8's cons-string +=.
  let s = "";
  for (const key of Object.keys(attributes)) {
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

function writeElement(
  element: Element,
  opts: ReturnType<typeof normalizeOptions>,
  depth: number,
): string {
  if (!element.name) return "";
  const name = element.name;
  const attrStr = element.attributes
    ? writeAttributes(element.attributes, name, element, opts.attributeValueFn)
    : "";
  // Deferred content: re-emit the captured inner XML verbatim — children were
  // never parsed, and the bytes must survive a set/save round-trip.
  if (element.raw !== undefined) {
    return `<${name}${attrStr}>${element.raw}</${name}>`;
  }
  const withClosingTag =
    (element.elements?.length ?? 0) > 0 ||
    element.attributes?.["xml:space"] === "preserve" ||
    opts.fullTagEmptyElement;

  if (!withClosingTag) {
    return `<${name}${attrStr}/>`;
  }

  const open = `<${name}${attrStr}>`;
  if (element.elements?.length) {
    const inner = writeElements(element.elements, opts, depth + 1, false);
    // The child-element scan is only needed to pretty-print the closing tag —
    // skip it entirely when not indenting.
    if (opts.spaces && element.elements.some((e) => e.type === "element")) {
      return open + inner + "\n" + opts.spaces.repeat(depth) + `</${name}>`;
    }
    return open + inner + `</${name}>`;
  }
  return open + `</${name}>`;
}

function writeElements(
  elements: Element[],
  opts: ReturnType<typeof normalizeOptions>,
  depth: number,
  firstLine: boolean,
): string {
  // Rope accumulation — V8 cons-strings make += O(1) and skip the parts-array
  // allocation the join-based form pays per level.
  let s = "";
  for (let i = 0; i < elements.length; i++) {
    const element = elements[i];
    if (!element) continue;
    const isFirst = firstLine && i === 0;
    switch (element.type) {
      case "element":
        if (opts.spaces) s += writeIndentation(opts.spaces, depth, isFirst);
        s += writeElement(element, opts, depth);
        break;
      case "text":
        if (opts.ignoreText) continue;
        if (opts.indentText && opts.spaces) s += writeIndentation(opts.spaces, depth, isFirst);
        s += writeText(element.text);
        break;
      case "cdata":
        if (opts.ignoreCdata) continue;
        if (opts.indentCdata && opts.spaces) s += writeIndentation(opts.spaces, depth, isFirst);
        s += writeCdata(element.cdata);
        break;
      case "comment":
        if (opts.ignoreComment) continue;
        if (opts.spaces) s += writeIndentation(opts.spaces, depth, isFirst);
        s += writeComment(element.comment);
        break;
      case "doctype":
        if (opts.ignoreDoctype) continue;
        if (opts.spaces) s += writeIndentation(opts.spaces, depth, isFirst);
        s += writeDoctype(element.doctype);
        break;
      default:
        break;
    }
  }
  return s;
}

function writeText(text: string | number | boolean | undefined | null): string {
  if (text == null) return "";
  const str = String(text);
  // Fast path: most text content contains no markup delimiters. A regex test
  // beats a charCodeAt loop by ~10× (V8 compiles it to a native scan) and
  // returns the original string reference — zero allocation for the common case.
  if (!TEXT_SPECIALS.test(str)) return str;
  // Chained native replaces: specials are sparse, so each pass after the first
  // usually finds nothing and exits quickly.
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function writeCdata(cdata: string | undefined | null): string {
  if (cdata == null) return "";
  const escaped = cdata.replace(/\]\]>/g, "]]]]><![CDATA[>");
  return `<![CDATA[${escaped}]]>`;
}

function writeComment(comment: string | undefined | null): string {
  if (comment == null) return "";
  return `<!--${comment}-->`;
}

function writeDoctype(doctype: string | undefined | null): string {
  if (doctype == null) return "";
  return `<!DOCTYPE ${doctype}>`;
}

type NonNullable<T> = T extends null | undefined ? never : T;
