import type { Element, Js2XmlOptions } from "./types";

export function js2xml(js: Element, options?: Js2XmlOptions): string {
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

/** Alias for js2xml — xml-js compatible export */
export function json2xml(json: Element, options?: Js2XmlOptions): string {
    return js2xml(json, options);
}

function normalizeOptions(options?: Js2XmlOptions): {
    spaces: string;
    ignoreDeclaration: boolean;
    ignoreText: boolean;
    ignoreComment: boolean;
    ignoreCdata: boolean;
    ignoreDoctype: boolean;
    fullTagEmptyElement: boolean;
    indentText: boolean;
    indentCdata: boolean;
    attributeValueFn?: Js2XmlOptions["attributeValueFn"];
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
    return (!firstLine && spaces ? "\n" : "") + spaces.repeat(depth);
}

function writeDeclaration(declaration: NonNullable<Element["declaration"]>): string {
    const attrs = declaration.attributes;
    if (!attrs) return '<?xml version="1.0"?>';

    let result = '<?xml version="1.0"';
    if (attrs.encoding) result += ` encoding="${attrs.encoding}"`;
    if (attrs.standalone) result += ` standalone="${attrs.standalone}"`;
    return result + "?>";
}

function writeAttributes(
    attributes: Record<string, string | number | undefined>,
    elementName: string,
    element: Element,
    attributeValueFn?: Js2XmlOptions["attributeValueFn"],
): string {
    const parts: string[] = [];
    for (const key of Object.keys(attributes)) {
        const value = attributes[key];
        if (value === null || value === undefined) continue;

        let attr = String(value).replace(/"/g, "&quot;");
        if (attributeValueFn) {
            attr = attributeValueFn(attr, key, elementName, element);
        }
        parts.push(` ${key}="${attr}"`);
    }
    return parts.join("");
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
    const withClosingTag =
        (element.elements?.length ?? 0) > 0 ||
        element.attributes?.["xml:space"] === "preserve" ||
        opts.fullTagEmptyElement;

    if (!withClosingTag) {
        return `<${name}${attrStr}/>`;
    }

    const parts: string[] = [];
    parts.push(`<${name}${attrStr}>`);
    const hasChildElements = element.elements?.some((e) => e.type === "element") ?? false;
    if (element.elements?.length) {
        parts.push(writeElements(element.elements, opts, depth + 1, false));
    }
    if (opts.spaces && hasChildElements) {
        parts.push("\n" + opts.spaces.repeat(depth));
    }
    parts.push(`</${name}>`);
    return parts.join("");
}

function writeElements(
    elements: Element[],
    opts: ReturnType<typeof normalizeOptions>,
    depth: number,
    firstLine: boolean,
): string {
    let result = "";
    for (let i = 0; i < elements.length; i++) {
        const element = elements[i];
        const isFirst = firstLine && i === 0;
        switch (element.type) {
            case "element":
                result += writeIndentation(opts.spaces, depth, isFirst);
                result += writeElement(element, opts, depth);
                break;
            case "text":
                if (opts.ignoreText) continue;
                if (opts.indentText) result += writeIndentation(opts.spaces, depth, isFirst);
                result += writeText(element.text);
                break;
            case "cdata":
                if (opts.ignoreCdata) continue;
                if (opts.indentCdata) result += writeIndentation(opts.spaces, depth, isFirst);
                result += writeCdata(element.cdata);
                break;
            case "comment":
                if (opts.ignoreComment) continue;
                result += writeIndentation(opts.spaces, depth, isFirst);
                result += writeComment(element.comment);
                break;
            case "doctype":
                if (opts.ignoreDoctype) continue;
                result += writeIndentation(opts.spaces, depth, isFirst);
                result += writeDoctype(element.doctype);
                break;
            default:
                break;
        }
    }
    return result;
}

function writeText(text: string | number | boolean | undefined | null): string {
    if (text == null) return "";
    const str = String(text);
    return str
        .replace(/&amp;/g, "&")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
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
