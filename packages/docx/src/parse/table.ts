import { findChild, children, attr, attrNum } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { DocxParseContext } from "./context";
import { parseParagraph } from "./paragraph";
import type { TableJson, TableRowJson, TableCellJson } from "./types";

export function parseTable(tbl: Element, ctx: DocxParseContext): TableJson {
    const result: TableJson = {
        $type: "table",
        rows: [],
    };

    const tblPr = findChild(tbl, "w:tblPr");
    if (tblPr) {
        parseTableProperties(tblPr, result);
    }

    // Parse rows
    const rows = children(tbl, "w:tr");
    for (const tr of rows) {
        result.rows.push(parseTableRow(tr, ctx));
    }

    return result;
}

function parseTableProperties(tblPr: Element, out: TableJson): void {
    // Table width
    const tblW = findChild(tblPr, "w:tblW");
    if (tblW) {
        const size = attrNum(tblW, "w:w");
        const type = attr(tblW, "w:type");
        if (size !== undefined) {
            out.width = { size, type: type ?? "auto" };
        }
    }

    // Table style
    const tblStyle = findChild(tblPr, "w:tblStyle");
    if (tblStyle) {
        const val = attr(tblStyle, "w:val");
        if (val) out.style = val;
    }

    // Table alignment
    const jc = findChild(tblPr, "w:jc");
    if (jc) {
        const val = attr(jc, "w:val");
        if (val) out.alignment = val;
    }

    // Column widths from tblGrid
    // Note: tblGrid is a sibling of tblPr, not a child
}

function parseTableRow(tr: Element, ctx: DocxParseContext): TableRowJson {
    const cells: TableCellJson[] = [];

    for (const child of tr.elements ?? []) {
        if (child.name === "w:tc") {
            cells.push(parseTableCell(child, ctx));
        }
    }

    const result: TableRowJson = { cells };
    return result;
}

function parseTableCell(tc: Element, ctx: DocxParseContext): TableCellJson {
    const result: TableCellJson = { children: [] };
    const tcPr = findChild(tc, "w:tcPr");

    if (tcPr) {
        // Column span (gridSpan)
        const gridSpan = findChild(tcPr, "w:gridSpan");
        if (gridSpan) {
            const val = attrNum(gridSpan, "w:val");
            if (val !== undefined && val > 1) result.columnSpan = val;
        }

        // Row span (vMerge)
        const vMerge = findChild(tcPr, "w:vMerge");
        if (vMerge) {
            const val = attr(vMerge, "w:val");
            if (val === "restart") {
                // Start of vertical merge — determine rowSpan by counting subsequent vMerge cells
                // For MVP, just note the merge restart
            }
        }

        // Cell width
        const tcW = findChild(tcPr, "w:tcW");
        if (tcW) {
            const size = attrNum(tcW, "w:w");
            const type = attr(tcW, "w:type");
            if (size !== undefined) {
                result.width = { size, type: type ?? "auto" };
            }
        }

        // Shading
        const shd = findChild(tcPr, "w:shd");
        if (shd) {
            const fill = attr(shd, "w:fill");
            if (fill && fill !== "auto") {
                result.shading = { fill };
            }
        }

        // Vertical alignment
        const vAlign = findChild(tcPr, "w:vAlign");
        if (vAlign) {
            const val = attr(vAlign, "w:val");
            if (val) result.verticalAlign = val;
        }
    }

    // Parse cell paragraphs
    for (const child of tc.elements ?? []) {
        if (child.name === "w:p") {
            result.children.push(parseParagraph(child, ctx));
        }
    }

    return result;
}
