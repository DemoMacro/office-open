/**
 * Document attributes module for WordprocessingML documents.
 *
 * This module defines the XML namespace declarations used in OOXML documents.
 * These namespaces are required for proper document parsing and generation.
 *
 * Reference: http://officeopenxml.com/anatomyofOOXML.php
 *
 * @module
 */
import type { IXmlableObject } from "@file/xml-components";

/**
 * XML namespace URIs used in WordprocessingML documents.
 *
 * These namespaces define the various XML schemas that can be referenced
 * in a document, including WordprocessingML, DrawingML, VML, and others.
 */
/* CSpell:disable */
export const DocumentAttributeNamespaces = {
    aink: "http://schemas.microsoft.com/office/drawing/2016/ink",
    am3d: "http://schemas.microsoft.com/office/drawing/2017/model3d",
    cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    cx: "http://schemas.microsoft.com/office/drawing/2014/chartex",
    cx1: "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    cx2: "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    cx3: "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    cx4: "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    cx5: "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
    cx6: "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    cx7: "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    cx8: "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    dc: "http://purl.org/dc/elements/1.1/",
    dcmitype: "http://purl.org/dc/dcmitype/",
    dcterms: "http://purl.org/dc/terms/",
    m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
    mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
    o: "urn:schemas-microsoft-com:office:office",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    v: "urn:schemas-microsoft-com:vml",
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    w10: "urn:schemas-microsoft-com:office:word",
    w14: "http://schemas.microsoft.com/office/word/2010/wordml",
    w15: "http://schemas.microsoft.com/office/word/2012/wordml",
    w16: "http://schemas.microsoft.com/office/word/2018/wordml",
    w16cex: "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    w16cid: "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    w16sdtdh: "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    w16se: "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    wne: "http://schemas.microsoft.com/office/word/2006/wordml",
    wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    wpc: "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    wpi: "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    xsi: "http://www.w3.org/2001/XMLSchema-instance",
};
/* CSpell:enable */

/**
 * Type representing valid namespace keys.
 */
export type DocumentAttributeNamespace = keyof typeof DocumentAttributeNamespaces;

const attrsCache = new Map<string, IXmlableObject>();

/**
 * Builds XML namespace attributes for a WordprocessingML document.
 *
 * Results are cached by (namespaces + ignorable) key.
 *
 * @internal
 */
export function buildDocumentAttributes(
    ns: readonly DocumentAttributeNamespace[],
    Ignorable?: string,
): IXmlableObject {
    const cacheKey = `${ns.join(",")}:${Ignorable ?? ""}`;
    const cached = attrsCache.get(cacheKey);
    if (cached) return cached;

    const attrs: Record<string, string> = {};
    if (Ignorable !== undefined) {
        attrs["mc:Ignorable"] = Ignorable;
    }
    for (const n of ns) {
        attrs[`xmlns:${n}`] = DocumentAttributeNamespaces[n];
    }
    const result = { _attr: attrs };
    attrsCache.set(cacheKey, result);
    return result;
}
