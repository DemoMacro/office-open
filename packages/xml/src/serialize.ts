import { escapeXml } from "./escape";

const DEFAULT_INDENT = "    ";

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
        const keys = Object.keys(items[i]);
        parts.push(formatElement(keys[0], items[i][keys[0]], opts.indent, 0));
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

/**
 * Single-pass XML formatter: directly converts IXmlableObject to string,
 * eliminating the intermediate ResolvedElement tree.
 */
function formatElement(name: string, values: any, indent: string, depth: number): string {
    const attrParts: string[] = [];
    const textParts: string[] = [];
    const elemParts: string[] = [];
    let emptyArray = false;

    if (values == null) {
        const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
        const ind = indent ? indent.repeat(depth) : "";
        return `${ind}<${name}${attrStr}/>`;
    }

    if (typeof values === "object") {
        if (values._attr) {
            for (const key of Object.keys(values._attr)) {
                attrParts.push(`${key}="${escapeXml(String(values._attr[key]))}"`);
            }
        }
        if (values._attributes) {
            for (const key of Object.keys(values._attributes)) {
                attrParts.push(`${key}="${escapeXml(String(values._attributes[key]))}"`);
            }
        }
        if (values._cdata) {
            const escaped = String(values._cdata).replace(/\]\]>/g, "]]]]><![CDATA[>");
            textParts.push(`<![CDATA[${escaped}]]>`);
        }
        if (Array.isArray(values)) {
            if (values.length === 0) {
                emptyArray = true;
            } else {
                for (const value of values) {
                    if (value && typeof value === "object" && "_attr" in value) {
                        for (const key of Object.keys(value._attr)) {
                            attrParts.push(`${key}="${escapeXml(String(value._attr[key]))}"`);
                        }
                    } else if (value && typeof value === "object" && "_attributes" in value) {
                        for (const key of Object.keys(value._attributes)) {
                            attrParts.push(`${key}="${escapeXml(String(value._attributes[key]))}"`);
                        }
                    } else if (value && typeof value === "object") {
                        const childKeys = Object.keys(value);
                        elemParts.push(
                            formatElement(
                                childKeys[0],
                                value[childKeys[0]],
                                indent,
                                depth + 1,
                            ),
                        );
                    } else if (value != null) {
                        textParts.push(escapeXml(String(value)));
                    }
                }
            }
        }
    } else {
        textParts.push(escapeXml(String(values)));
    }

    const ind = indent ? indent.repeat(depth) : "";
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    const totalParts = textParts.length + elemParts.length;

    if (totalParts === 0) {
        return emptyArray ? `${ind}<${name}${attrStr}></${name}>` : `${ind}<${name}${attrStr}/>`;
    }

    // Text-only optimization: single text child, no element children
    if (elemParts.length === 0 && textParts.length === 1) {
        return indent
            ? `${ind}<${name}${attrStr}>${textParts[0]}</${name}>`
            : `<${name}${attrStr}>${textParts[0]}</${name}>`;
    }

    // Mixed content
    const parts: string[] = [];
    parts.push(`${ind}<${name}${attrStr}>`);
    if (indent) parts.push("\n");
    const childIndent = indent ? indent.repeat(depth + 1) : "";
    for (const t of textParts) {
        parts.push(`${childIndent}${t}`);
        if (indent) parts.push("\n");
    }
    for (const e of elemParts) {
        parts.push(e);
        if (indent) parts.push("\n");
    }
    parts.push(`${ind}</${name}>`);
    return parts.join("");
}
