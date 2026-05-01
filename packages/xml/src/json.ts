import { xml2js } from "./parse";
import type { Xml2JsOptions } from "./types";

/** Convert XML string to JSON string — xml-js compatible export */
export function xml2json(xml: string, options?: Xml2JsOptions): string {
    return JSON.stringify(xml2js(xml, options));
}
