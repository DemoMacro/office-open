import type { Element, Xml2JsOptions } from "./types";

const ENTITY_MAP: Record<string, string> = {
    "&amp;": "&",
    "&lt;": "<",
    "&gt;": ">",
    "&quot;": '"',
    "&apos;": "'",
};
const ENTITY_PATTERN = /&(amp|lt|gt|quot|apos);/g;

function unescapeXml(str: string): string {
    return str.replace(ENTITY_PATTERN, (match) => ENTITY_MAP[match]);
}

function nativeTypeValue(value: string): string | number | boolean {
    const n = Number(value);
    if (!isNaN(n)) return n;
    const lower = value.toLowerCase();
    if (lower === "true") return true;
    if (lower === "false") return false;
    return value;
}

export function xml2js(xmlString: string, options?: Xml2JsOptions): Element {
    const captureSpaces = options?.captureSpacesBetweenElements ?? false;
    const trim = options?.trim ?? false;
    const ignoreDeclaration = options?.ignoreDeclaration ?? false;
    const ignoreText = options?.ignoreText ?? false;
    const ignoreComment = options?.ignoreComment ?? false;
    const ignoreCdata = options?.ignoreCdata ?? false;
    const ignoreDoctype = options?.ignoreDoctype ?? false;
    const nativeTypeAttributes = options?.nativeTypeAttributes ?? false;

    const result: Element = {};
    const stack: Element[] = [result];

    let i = 0;
    const len = xmlString.length;

    while (i < len) {
        if (!captureSpaces && isWhitespace(xmlString.charCodeAt(i))) {
            i++;
            continue;
        }

        if (xmlString.charCodeAt(i) !== 0x3c /* < */) {
            const start = i;
            while (i < len && xmlString.charCodeAt(i) !== 0x3c) i++;
            let text = unescapeXml(xmlString.slice(start, i));
            if (trim) text = text.trim();
            if (ignoreText) continue;
            if (text.length > 0) {
                if (captureSpaces || text.trim().length > 0) {
                    addField(stack[stack.length - 1], "text", text);
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
                    const attrs = parseAttributes(xmlMatch[1]);
                    if (nativeTypeAttributes) {
                        for (const key of Object.keys(attrs)) {
                            attrs[key] = nativeTypeValue(attrs[key]) as string;
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
                if (trim) addField(stack[stack.length - 1], "comment", comment.trim());
                else addField(stack[stack.length - 1], "comment", comment);
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
                if (trim) addField(stack[stack.length - 1], "cdata", cdata.trim());
                else addField(stack[stack.length - 1], "cdata", cdata);
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
                addField(stack[stack.length - 1], "doctype", doctype);
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

        if (nativeTypeAttributes) {
            for (const key of Object.keys(attributes.attrs)) {
                attributes.attrs[key] = nativeTypeValue(attributes.attrs[key]) as string;
            }
        }

        const isSelfClosing = xmlString.charCodeAt(pos) === 0x2f /* / */;
        if (isSelfClosing) pos += 2;
        else pos++;

        const element: Element = {
            type: "element",
            name: tagName,
        };
        if (Object.keys(attributes.attrs).length > 0) {
            element.attributes = attributes.attrs;
        }

        const parent = stack[stack.length - 1];
        if (!parent.elements) {
            parent.elements = [];
        }
        parent.elements.push(element);

        if (!isSelfClosing) {
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
        if (
            ch === 0x20 ||
            ch === 0x09 ||
            ch === 0x0a ||
            ch === 0x0d ||
            ch === 0x2f ||
            ch === 0x3e
        ) {
            return i;
        }
        i++;
    }
    return i;
}

function parseAttributesFromXml(
    str: string,
    start: number,
): { attrs: Record<string, string>; pos: number } {
    const attrs: Record<string, string> = {};
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
        attrs[name] = str.slice(valueStart, i);
        i++;
    }

    return { attrs, pos: i };
}

function parseAttributes(str: string): Record<string, string> {
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
        result[name] = str.slice(valueStart, i);
        i++;
    }

    return result;
}

function addField(parent: Element, type: string, value: string) {
    if (!parent.elements) {
        parent.elements = [];
    }
    const element: Element = { type };
    (element as Record<string, unknown>)[type] = value;
    parent.elements.push(element);
}

function isWhitespace(ch: number): boolean {
    return ch === 0x20 || ch === 0x09 || ch === 0x0a || ch === 0x0d;
}
