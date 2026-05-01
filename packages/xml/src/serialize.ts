import { escapeXml } from "./escape";

const DEFAULT_INDENT = "    ";

interface ResolvedElement {
    name: string;
    attributes: string[];
    content: (ResolvedElement | string)[];
    indent: string;
    depth: number;
    emptyArray: boolean;
}

export function xml(
    input: Record<string, any> | Record<string, any>[],
    options?:
        | boolean
        | string
        | {
              indent?: boolean | string;
              declaration?: boolean | { encoding?: string; standalone?: string };
          },
): string {
    const opts = normalizeOptions(options);
    const parts: string[] = [];

    if (opts.declaration) {
        const declOpts = opts.declaration === true ? {} : opts.declaration;
        const enc = declOpts.encoding || "UTF-8";
        const sa = declOpts.standalone;
        let decl = '<?xml version="1.0" encoding="' + enc + '"';
        if (sa) decl += ' standalone="' + sa + '"';
        decl += "?>";
        parts.push(decl);
        if (opts.indent) parts.push("\n");
    }

    const items = Array.isArray(input) ? input : [input];
    for (let i = 0; i < items.length; i++) {
        parts.push(formatElement(resolve(items[i], opts.indent, 0)));
        if (opts.indent && i < items.length - 1) parts.push("\n");
    }

    return parts.join("");
}

type XmlInputOptions = {
    indent?: boolean | string;
    declaration?: boolean | { encoding?: string; standalone?: string };
};

function normalizeOptions(options?: boolean | string | XmlInputOptions): {
    indent: string;
    declaration: XmlInputOptions["declaration"];
} {
    const opts =
        typeof options === "object" && !Array.isArray(options)
            ? options
            : { indent: options as boolean | string };
    let indent = "";
    if (opts.indent) {
        indent = opts.indent === true ? DEFAULT_INDENT : String(opts.indent);
    }
    return { indent, declaration: opts.declaration };
}

function resolve(data: Record<string, any>, indent: string, depth: number): ResolvedElement {
    const keys = Object.keys(data);
    const name = keys[0];
    const values = data[name];
    const attributes: string[] = [];
    const content: (ResolvedElement | string)[] = [];

    if (values == null) {
        return { name, attributes, content, indent, depth, emptyArray: false };
    }

    switch (typeof values) {
        case "object": {
            if (values._attr) {
                for (const key of Object.keys(values._attr)) {
                    attributes.push(`${key}="${escapeXml(String(values._attr[key]))}"`);
                }
            }
            if (values._cdata) {
                const escaped = String(values._cdata).replace(/\]\]>/g, "]]]]><![CDATA[>");
                content.push(`<![CDATA[${escaped}]]>`);
            }
            if (Array.isArray(values)) {
                if (values.length === 0) {
                    return { name, attributes, content, indent, depth, emptyArray: true };
                }
                for (const value of values) {
                    if (value && typeof value === "object" && "_attr" in value) {
                        for (const key of Object.keys(value._attr)) {
                            attributes.push(`${key}="${escapeXml(String(value._attr[key]))}"`);
                        }
                    } else if (value && typeof value === "object") {
                        content.push(resolve(value, indent, depth + 1));
                    } else if (value != null) {
                        content.push(escapeXml(String(value)));
                    }
                }
            }
            break;
        }
        default:
            content.push(escapeXml(String(values)));
    }

    return { name, attributes, content, indent, depth, emptyArray: false };
}

function formatElement(elem: ResolvedElement): string {
    const { name, attributes, content, indent, depth } = elem;
    const hasChildren = content.length > 0;
    const ind = indent ? indent.repeat(depth) : "";

    const attrStr = attributes.length ? " " + attributes.join(" ") : "";

    if (!hasChildren) {
        if (elem.emptyArray) {
            return `${ind}<${name}${attrStr}></${name}>`;
        }
        return `${ind}<${name}${attrStr}/>`;
    }

    // Check if content is purely text (single string, no nested elements)
    const textContent = content.length === 1 && typeof content[0] === "string" ? content[0] : null;
    if (textContent !== null && !indent) {
        return `<${name}${attrStr}>${textContent}</${name}>`;
    }
    if (textContent !== null) {
        return `${ind}<${name}${attrStr}>${textContent}</${name}>`;
    }

    const parts: string[] = [];
    parts.push(`${ind}<${name}${attrStr}>`);
    if (indent) parts.push("\n");
    for (const child of content) {
        if (typeof child === "string") {
            parts.push(`${indent.repeat(depth + 1)}${child}`);
        } else {
            parts.push(formatElement(child));
        }
        if (indent) parts.push("\n");
    }
    parts.push(`${ind}</${name}>`);
    return parts.join("");
}
