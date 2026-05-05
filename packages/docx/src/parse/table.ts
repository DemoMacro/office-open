import { findChild, children, attr, attrNum, colorAttr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { DocxParseContext } from "./context";
import { parseParagraph } from "./paragraph";
import { parseSdtContent } from "./sdt";
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

    // Column widths from tblGrid (sibling of tblPr)
    const tblGrid = children(tbl, "w:tblGrid")[0];
    if (tblGrid) {
        const gridCols = children(tblGrid, "w:gridCol");
        if (gridCols.length > 0) {
            result.columnWidths = gridCols.map((col) => attrNum(col, "w:w") ?? 0);
        }
    }

    // Parse rows
    const rows = children(tbl, "w:tr");
    for (const tr of rows) {
        result.rows.push(parseTableRow(tr, ctx));
    }

    // Post-process: calculate actual rowSpan from vMerge
    calculateRowSpans(result);

    return result;
}

export function parseTableProperties(tblPr: Element, out: TableJson): void {
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

    // Table borders
    const tblBorders = findChild(tblPr, "w:tblBorders");
    if (tblBorders) {
        const borders: Record<string, unknown> = {};
        for (const child of tblBorders.elements ?? []) {
            if (child.name && child.name.startsWith("w:")) {
                const borderName = child.name.replace("w:", "");
                const val = String(attr(child, "w:val") ?? "");
                const sz = attrNum(child, "w:sz");
                const color = colorAttr(child, "w:color");
                const space = attrNum(child, "w:space");
                if (val && val !== "none" && val !== "nil") {
                    const borderDef: Record<string, unknown> = { style: val };
                    if (sz !== undefined) borderDef.size = sz;
                    if (color) borderDef.color = color;
                    if (space !== undefined) borderDef.space = space;
                    borders[borderName] = borderDef;
                }
            }
        }
        if (Object.keys(borders).length > 0) out.borders = borders;
    }

    // Table cell margins
    const tblCellMar = findChild(tblPr, "w:tblCellMar");
    if (tblCellMar) {
        const margins: Record<string, unknown> = {};
        for (const child of tblCellMar.elements ?? []) {
            if (child.name && child.name.startsWith("w:")) {
                const name = child.name.replace("w:", "");
                const w = attrNum(child, "w:w");
                const type = attr(child, "w:type");
                if (w !== undefined) {
                    margins[name] = { w, type: type ?? "dxa" };
                }
            }
        }
        if (Object.keys(margins).length > 0) out.cellMargins = margins;
    }

    // Table indentation
    const tblInd = findChild(tblPr, "w:tblInd");
    if (tblInd) {
        const w = attrNum(tblInd, "w:w");
        const type = attr(tblInd, "w:type");
        if (w !== undefined) {
            out.indentation = { w, type: type ?? "dxa" };
        }
    }

    // Table layout
    const tblLayout = findChild(tblPr, "w:tblLayout");
    if (tblLayout) {
        const val = attr(tblLayout, "w:type");
        if (val) out.layout = val;
    }
}

export function parseTableRow(tr: Element, ctx: DocxParseContext): TableRowJson {
    const result: TableRowJson = { cells: [] };

    const trPr = findChild(tr, "w:trPr");
    if (trPr) {
        // Row height
        const trHeight = findChild(trPr, "w:trHeight");
        if (trHeight) {
            const val = attrNum(trHeight, "w:val");
            const rule = attr(trHeight, "w:hRule");
            if (val !== undefined) {
                result.height = { value: val, rule: rule ?? undefined };
            }
        }

        // Table header row repeat
        const tblHeader = findChild(trPr, "w:tblHeader");
        if (tblHeader) {
            result.isHeader = true;
        }
    }

    for (const child of tr.elements ?? []) {
        if (child.name === "w:tc") {
            result.cells.push(parseTableCell(child, ctx));
        }
    }

    return result;
}

export function parseTableCell(tc: Element, ctx: DocxParseContext): TableCellJson {
    const result: TableCellJson = { children: [] };
    const tcPr = findChild(tc, "w:tcPr");

    if (tcPr) {
        // Column span (gridSpan)
        const gridSpan = findChild(tcPr, "w:gridSpan");
        if (gridSpan) {
            const val = attrNum(gridSpan, "w:val");
            if (val !== undefined && val > 1) result.columnSpan = val;
        }

        // Row span (vMerge) — mark restart, actual span calculated later
        const vMerge = findChild(tcPr, "w:vMerge");
        if (vMerge) {
            const val = attr(vMerge, "w:val");
            if (val === "restart") {
                result.rowSpan = 1; // Will be updated by calculateRowSpans
            } else {
                // Continuation merge — rowSpan will be inherited
                result.rowSpan = 0;
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
            const val = attr(shd, "w:val");
            if (fill && fill !== "auto") {
                result.shading = { fill, ...(val && val !== "clear" && { type: val }) };
            }
        }

        // Vertical alignment
        const vAlign = findChild(tcPr, "w:vAlign");
        if (vAlign) {
            const val = attr(vAlign, "w:val");
            if (val) result.verticalAlign = val;
        }

        // No wrap
        const noWrap = findChild(tcPr, "w:noWrap");
        if (noWrap) {
            result.noWrap = true;
        }

        // Text direction
        const textDirection = findChild(tcPr, "w:textDirection");
        if (textDirection) {
            const val = attr(textDirection, "w:val");
            if (val) result.textDirection = val;
        }

        // Cell margins
        const tcMar = findChild(tcPr, "w:tcMar");
        if (tcMar) {
            const margins: Record<string, unknown> = {};
            for (const child of tcMar.elements ?? []) {
                if (child.name && child.name.startsWith("w:")) {
                    const name = child.name.replace("w:", "");
                    const w = attrNum(child, "w:w");
                    const type = attr(child, "w:type");
                    if (w !== undefined) {
                        margins[name] = { w, type: type ?? "dxa" };
                    }
                }
            }
            if (Object.keys(margins).length > 0) result.margins = margins;
        }

        // Cell borders
        const tcBorders = findChild(tcPr, "w:tcBorders");
        if (tcBorders) {
            const borders: Record<string, unknown> = {};
            for (const child of tcBorders.elements ?? []) {
                if (child.name && child.name.startsWith("w:")) {
                    const borderName = child.name.replace("w:", "");
                    const val = String(attr(child, "w:val") ?? "");
                    const sz = attrNum(child, "w:sz");
                    const color = colorAttr(child, "w:color");
                    const space = attrNum(child, "w:space");
                    if (val && val !== "none" && val !== "nil") {
                        const borderDef: Record<string, unknown> = { style: val };
                        if (sz !== undefined) borderDef.size = sz;
                        if (color) borderDef.color = color;
                        if (space !== undefined) borderDef.space = space;
                        borders[borderName] = borderDef;
                    }
                }
            }
            if (Object.keys(borders).length > 0) result.borders = borders;
        }
    }

    // Parse cell content
    for (const child of tc.elements ?? []) {
        if (child.name === "w:p") {
            result.children.push(parseParagraph(child, ctx));
        } else if (child.name === "w:tcPr") {
            // Already processed above
        } else if (child.name === "w:sdt") {
            // SDT wrapping cell content
            const sdt = parseSdtContent(child, ctx);
            if (sdt) result.children.push(sdt as unknown as (typeof result.children)[number]);
        } else if (child.name === "w:tbl") {
            // Nested table — preserve as raw (rare)
            result.children.push({ $raw: true, element: child });
        } else {
            result.children.push({ $raw: true, element: child });
        }
    }

    return result;
}

/** Post-process: calculate actual rowSpan values from vMerge patterns */
export function calculateRowSpans(table: TableJson): void {
    // Track merge spans by column index
    const mergeCounts: number[] = [];

    for (const row of table.rows) {
        let colIdx = 0;
        for (const cell of row.cells) {
            if (cell.rowSpan === 1) {
                // Start of merge — we'll count as we go
                mergeCounts[colIdx] = 1;
            } else if (cell.rowSpan === 0 || cell.rowSpan === undefined) {
                // Continuation — increment parent merge
                if (mergeCounts[colIdx] !== undefined) {
                    mergeCounts[colIdx]++;
                }
            }

            // Advance by column span
            const span = cell.columnSpan ?? 1;
            for (let i = 1; i < span; i++) {
                colIdx++;
                mergeCounts[colIdx] = 0; // Spanned columns don't get their own merge
            }
            colIdx++;
        }
    }

    // Now set the final rowSpan values
    for (const row of table.rows) {
        let colIdx = 0;
        for (const cell of row.cells) {
            if (cell.rowSpan === 1 && mergeCounts[colIdx] !== undefined) {
                cell.rowSpan = mergeCounts[colIdx];
            } else if (cell.rowSpan === 0) {
                cell.rowSpan = undefined; // Remove placeholder
            }

            const span = cell.columnSpan ?? 1;
            colIdx += span;
        }
    }
}
