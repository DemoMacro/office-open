import { textOf } from "@office-open/xml";

import { readXmlFromZip } from "./archive";

export interface CoreProperties {
    title?: string;
    subject?: string;
    creator?: string;
    keywords?: string;
    description?: string;
    lastModifiedBy?: string;
    revision?: string;
    created?: string;
    modified?: string;
}

const FIELD_MAP: Array<{ name: string; key: keyof CoreProperties }> = [
    { name: "dc:title", key: "title" },
    { name: "dc:subject", key: "subject" },
    { name: "dc:creator", key: "creator" },
    { name: "dc:description", key: "description" },
    { name: "cp:keywords", key: "keywords" },
    { name: "cp:lastModifiedBy", key: "lastModifiedBy" },
];

export function parseCoreProperties(zip: Map<string, Uint8Array>): CoreProperties {
    const xml = readXmlFromZip(zip, "docProps/core.xml");
    if (!xml) return {};

    const props: CoreProperties = {};

    for (const field of FIELD_MAP) {
        const el = xml.elements?.find((e) => e.name === field.name);
        const value = textOf(el) || undefined;
        if (value) (props as Record<string, unknown>)[field.key] = value;
    }

    const revEl = xml.elements?.find((e) => e.name === "cp:revision");
    if (revEl) {
        const rev = textOf(revEl);
        if (rev) {
            const n = Number(rev);
            if (!isNaN(n)) props.revision = rev;
        }
    }

    return props;
}
