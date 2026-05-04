import { findChild, children, attr, attrNum } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { SectionJson } from "./types";

export function parseSectionProperties(sectPr: Element): SectionJson["properties"] {
    const props: Record<string, unknown> = {};

    // Page size
    const pgSz = findChild(sectPr, "w:pgSz");
    if (pgSz) {
        const size: Record<string, unknown> = {};
        const w = attrNum(pgSz, "w:w");
        const h = attrNum(pgSz, "w:h");
        const orient = attr(pgSz, "w:orient");
        if (w !== undefined) size.width = w;
        if (h !== undefined) size.height = h;
        if (orient) size.orientation = orient;
        if (Object.keys(size).length > 0) {
            if (!props.page) props.page = {};
            (props.page as Record<string, unknown>).size = size;
        }
    }

    // Page margins
    const pgMar = findChild(sectPr, "w:pgMar");
    if (pgMar) {
        const margin: Record<string, unknown> = {};
        const top = attrNum(pgMar, "w:top");
        const right = attrNum(pgMar, "w:right");
        const bottom = attrNum(pgMar, "w:bottom");
        const left = attrNum(pgMar, "w:left");
        const gutter = attrNum(pgMar, "w:gutter");
        const header = attrNum(pgMar, "w:header");
        const footer = attrNum(pgMar, "w:footer");

        if (top !== undefined) margin.top = top;
        if (right !== undefined) margin.right = right;
        if (bottom !== undefined) margin.bottom = bottom;
        if (left !== undefined) margin.left = left;
        if (gutter !== undefined) margin.gutter = gutter;
        if (header !== undefined) margin.header = header;
        if (footer !== undefined) margin.footer = footer;

        if (Object.keys(margin).length > 0) {
            if (!props.page) props.page = {};
            (props.page as Record<string, unknown>).margin = margin;
        }
    }

    // Section type
    const type = findChild(sectPr, "w:type");
    if (type) {
        const val = attr(type, "w:val");
        if (val) props.type = val;
    }

    // Title page
    const titlePg = findChild(sectPr, "w:titlePg");
    if (titlePg) {
        const val = attr(titlePg, "w:val");
        if (val === undefined || (val !== "0" && val !== "false")) {
            props.titlePage = true;
        }
    }

    // Page number start
    const pgNumType = findChild(sectPr, "w:pgNumType");
    if (pgNumType) {
        const start = attrNum(pgNumType, "w:start");
        if (start !== undefined) {
            if (!props.page) props.page = {};
            if (!(props.page as Record<string, unknown>).pageNumbers) {
                (props.page as Record<string, unknown>).pageNumbers = {};
            }
            ((props.page as Record<string, unknown>).pageNumbers as Record<string, unknown>).start =
                start;
        }
    }

    // Columns
    const cols = findChild(sectPr, "w:cols");
    if (cols) {
        const num = attrNum(cols, "w:num");
        const space = attrNum(cols, "w:space");
        if ((num !== undefined && num > 1) || space !== undefined) {
            const col: Record<string, unknown> = {};
            if (num !== undefined) col.count = num;
            if (space !== undefined) col.space = space;
            props.column = col;
        }
    }

    // Vertical align
    const vAlign = findChild(sectPr, "w:vAlign");
    if (vAlign) {
        const val = attr(vAlign, "w:val");
        if (val) props.verticalAlign = val;
    }

    // Line numbers
    const lnNumType = findChild(sectPr, "w:lnNumType");
    if (lnNumType) {
        const countBy = attrNum(lnNumType, "w:countBy");
        const restart = attr(lnNumType, "w:restart");
        const start = attrNum(lnNumType, "w:start");
        if (countBy !== undefined || restart || start !== undefined) {
            const lineNumbers: Record<string, unknown> = {};
            if (countBy !== undefined) lineNumbers.countBy = countBy;
            if (restart) lineNumbers.restart = restart;
            if (start !== undefined) lineNumbers.start = start;
            props.lineNumbers = lineNumbers;
        }
    }

    // Header/footer references
    const headerRefs = children(sectPr, "w:headerReference");
    const footerRefs = children(sectPr, "w:footerReference");

    for (const ref of headerRefs) {
        const type = attr(ref, "w:type");
        const rId = attr(ref, "r:id");
        if (type && rId) {
            if (!props.headerRefs) props.headerRefs = {};
            (props.headerRefs as Record<string, string>)[type] = rId;
        }
    }

    for (const ref of footerRefs) {
        const type = attr(ref, "w:type");
        const rId = attr(ref, "r:id");
        if (type && rId) {
            if (!props.footerRefs) props.footerRefs = {};
            (props.footerRefs as Record<string, string>)[type] = rId;
        }
    }

    // Borders
    const pgBorders = findChild(sectPr, "w:pgBorders");
    if (pgBorders) {
        const borders: Record<string, unknown> = {};
        for (const border of pgBorders.elements ?? []) {
            if (border.name?.startsWith("w:")) {
                const borderName = border.name?.replace("w:", "") ?? "";
                const val = attr(border, "w:val");
                const sz = attrNum(border, "w:sz");
                const color = attr(border, "w:color");
                if (val && val !== "none" && val !== "nil") {
                    const borderDef: Record<string, unknown> = { style: val };
                    if (sz !== undefined) borderDef.size = sz;
                    if (color) borderDef.color = color;
                    borders[borderName] = borderDef;
                }
            }
        }
        if (Object.keys(borders).length > 0) {
            if (!props.page) props.page = {};
            (props.page as Record<string, unknown>).borders = borders;
        }
    }

    return Object.keys(props).length > 0 ? props : undefined;
}
