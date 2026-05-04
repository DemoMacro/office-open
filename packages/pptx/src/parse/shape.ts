import { uint8ToBase64, getImageType } from "@office-open/core";
import { findChild, attr, attrNum, findDeep } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { PptxParseContext } from "./context";
import { resolveSlideMediaPath } from "./context";
import { parseTransform, parseGeometry, parseFill, parseOutline, parseEffects } from "./drawingml";
import { parseTextBody } from "./text-body";
import type {
    ShapeJson,
    PictureJson,
    ConnectorShapeJson,
    TableJson,
    TableCellJson,
    TableRowJson,
} from "./types";

export function parseShape(
    sp: Element,
    ctx: PptxParseContext,
    _slideRels: Map<string, { target: string; type: string }>,
): ShapeJson {
    const result: ShapeJson = { $type: "shape" };

    // NVSpPr (non-visual shape properties)
    const nvSpPr = findChild(sp, "p:nvSpPr");
    if (nvSpPr) {
        const cNvPr = findChild(nvSpPr, "p:cNvPr");
        if (cNvPr) {
            const id = attrNum(cNvPr, "id");
            const name = attr(cNvPr, "name");
            if (id !== undefined) result.id = id;
            if (name) result.name = name;
        }

        const nvPr = findChild(nvSpPr, "p:nvPr");
        if (nvPr) {
            const ph = findChild(nvPr, "p:ph");
            if (ph) {
                const type = attr(ph, "type") ?? attr(ph, "idx");
                // Map placeholder types
                if (type === "title" || type === "0") result.placeholder = "title";
                else if (type === "body" || type === "1") result.placeholder = "body";
                else if (type === "subTitle" || type === "2") result.placeholder = "subTitle";
                else if (type === "sldNum" || type === "3") result.placeholder = "sldNum";
                else if (type === "dt" || type === "4") result.placeholder = "dt";
                else if (type === "ftr" || type === "5") result.placeholder = "ftr";
                else if (type === "hdr" || type === "6") result.placeholder = "hdr";
                else if (type) result.placeholder = type as ShapeJson["placeholder"];

                const idx = attrNum(ph, "idx");
                if (idx !== undefined) result.placeholderIndex = idx;
            }
        }
    }

    // SpPr (shape properties)
    const spPr = findChild(sp, "p:spPr");
    if (spPr) {
        const transform = parseTransform(spPr);
        Object.assign(result, transform);

        result.geometry = parseGeometry(spPr);
        result.fill = parseFill(spPr);
        result.outline = parseOutline(spPr);
        result.effects = parseEffects(spPr);
    }

    // Text body
    const txBody = findChild(sp, "p:txBody");
    if (txBody) {
        const textBody = parseTextBody(txBody, ctx);
        if (textBody.paragraphs) result.paragraphs = textBody.paragraphs;
        if (textBody.vertical) result.textVertical = textBody.vertical;
        if (textBody.anchor) result.textAnchor = textBody.anchor;
        if (textBody.autoFit) result.textAutoFit = textBody.autoFit;
    }

    return result;
}

export function parsePicture(
    pic: Element,
    ctx: PptxParseContext,
    slideRels: Map<string, { target: string; type: string }>,
): PictureJson | undefined {
    const result: PictureJson = { $type: "picture" };

    // NVPicPr (non-visual picture properties)
    const nvPicPr = findChild(pic, "p:nvPicPr");
    if (nvPicPr) {
        const cNvPr = findChild(nvPicPr, "p:cNvPr");
        if (cNvPr) {
            const id = attrNum(cNvPr, "id");
            const name = attr(cNvPr, "name");
            if (id !== undefined) result.id = id;
            if (name) result.name = name;
        }
    }

    // BlipFill (image data)
    const blipFill = findChild(pic, "p:blipFill");
    if (blipFill) {
        const blip = findChild(blipFill, "a:blip");
        if (blip) {
            const embedId = attr(blip, "r:embed");

            if (embedId) {
                const rel = slideRels.get(embedId);
                if (rel) {
                    const mediaPath = resolveSlideMediaPath(rel.target);
                    const data = ctx.zip.get(mediaPath);
                    if (data) {
                        result.data = uint8ToBase64(data);
                        result.type = getImageType(mediaPath);
                    }
                }
            }
        }
    }

    if (!result.data) return undefined;

    // Shape properties (transform, outline, etc.)
    const spPr = findChild(pic, "p:spPr");
    if (spPr) {
        const transform = parseTransform(spPr);
        Object.assign(result, transform);
        result.outline = parseOutline(spPr);
        result.effects = parseEffects(spPr);
    }

    return result;
}

export function parseConnector(cxnSp: Element, _ctx: PptxParseContext): ConnectorShapeJson {
    const result: ConnectorShapeJson = { $type: "connectorShape" };

    // NV connector properties
    const nvCxnSpPr = findChild(cxnSp, "p:nvCxnSpPr");
    if (nvCxnSpPr) {
        const cNvPr = findChild(nvCxnSpPr, "p:cNvPr");
        if (cNvPr) {
            const id = attrNum(cNvPr, "id");
            const name = attr(cNvPr, "name");
            if (id !== undefined) result.id = id;
            if (name) result.name = name;
        }
    }

    // Shape properties
    const spPr = findChild(cxnSp, "p:spPr");
    if (spPr) {
        const transform = parseTransform(spPr);
        Object.assign(result, transform);
        result.outline = parseOutline(spPr);

        // Connector style
        const prstGeom = findChild(spPr, "a:prstGeom");
        if (prstGeom) {
            const type = attr(prstGeom, "prst");
            if (type) result.style = type;
        }
    }

    return result;
}

export function parsePptxTable(
    graphicFrame: Element,
    ctx: PptxParseContext,
): TableJson | undefined {
    const result: TableJson = { $type: "table", rows: [] };

    // Non-visual properties
    // xfrm is in the spPr child
    const spPr = findChild(graphicFrame, "p:xfrm") ?? findChild(graphicFrame, "a:xfrm");
    if (spPr) {
        const off = findChild(spPr, "a:off");
        if (off) {
            result.x = attrNum(off, "x");
            result.y = attrNum(off, "y");
        }
        const ext = findChild(spPr, "a:ext");
        if (ext) {
            result.width = attrNum(ext, "cx");
            result.height = attrNum(ext, "cy");
        }
    }

    // Find the actual table element
    const tbl = findDeep(graphicFrame, "a:tbl")[0];
    if (!tbl) return undefined;

    // Table properties (a:tblPr)
    const tblPr = findChild(tbl, "a:tblPr");
    if (tblPr) {
        const bandRow = attr(tblPr, "bandRow");
        const bandCol = attr(tblPr, "bandCol");
        const firstRow = attr(tblPr, "firstRow");
        const firstCol = attr(tblPr, "firstCol");
        const lastRow = attr(tblPr, "lastRow");
        const lastCol = attr(tblPr, "lastCol");
        if (bandRow || bandCol || firstRow || firstCol || lastRow || lastCol) {
            result.tableStyle = {
                ...(bandRow === "1" && { bandRow: true }),
                ...(bandCol === "1" && { bandCol: true }),
                ...(firstRow === "1" && { firstRow: true }),
                ...(firstCol === "1" && { firstCol: true }),
                ...(lastRow === "1" && { lastRow: true }),
                ...(lastCol === "1" && { lastCol: true }),
            };
        }
    }

    // Parse rows
    for (const tr of tbl.elements ?? []) {
        if (tr.name !== "a:tr") continue;

        const row: TableRowJson = { cells: [] };
        const trH = attrNum(tr, "h");
        if (trH !== undefined) row.height = trH;

        for (const tc of tr.elements ?? []) {
            if (tc.name !== "a:tc") continue;

            const cell: TableCellJson = {};

            // Column span
            const gridSpan = attrNum(tc, "gridSpan");
            if (gridSpan !== undefined && gridSpan > 1) cell.columnSpan = gridSpan;

            // Row span
            const rowSpan = attrNum(tc, "rowSpan");
            if (rowSpan !== undefined && rowSpan > 1) cell.rowSpan = rowSpan;

            // Cell properties
            const tcPr = findChild(tc, "a:tcPr");
            if (tcPr) {
                // Fill
                const solidFill = findChild(tcPr, "a:solidFill");
                if (solidFill) {
                    const srgbClr = findChild(solidFill, "a:srgbClr");
                    if (srgbClr) {
                        const color = attr(srgbClr, "val");
                        if (color) cell.fill = { type: "solid", color };
                    }
                }

                // Vertical align
                const anchor = attr(tcPr, "anchor");
                if (anchor) cell.verticalAlign = anchor;

                // Margins
                const marL = attrNum(tcPr, "marL");
                const marR = attrNum(tcPr, "marR");
                const marT = attrNum(tcPr, "marT");
                const marB = attrNum(tcPr, "marB");
                if (
                    marL !== undefined ||
                    marR !== undefined ||
                    marT !== undefined ||
                    marB !== undefined
                ) {
                    const margins: Record<string, number> = {};
                    if (marL !== undefined) margins.left = marL;
                    if (marR !== undefined) margins.right = marR;
                    if (marT !== undefined) margins.top = marT;
                    if (marB !== undefined) margins.bottom = marB;
                    cell.margins = margins;
                }

                // Width
                const tcW = attrNum(tcPr, "w");
                if (tcW !== undefined) cell.width = tcW;

                // Cell borders
                const tcBorders = findChild(tcPr, "a:tcBorders");
                if (tcBorders) {
                    const borders: Record<string, unknown> = {};
                    for (const border of tcBorders.elements ?? []) {
                        if (border.name?.startsWith("a:")) {
                            const borderName = border.name.replace("a:", "");
                            const w = attrNum(border, "w");
                            if (w !== undefined) {
                                const borderDef: Record<string, unknown> = { width: w };
                                // Extract color from a:solidFill/a:srgbClr/@val
                                const solidFill = findChild(border, "a:solidFill");
                                if (solidFill) {
                                    const srgbClr = findChild(solidFill, "a:srgbClr");
                                    if (srgbClr) {
                                        const color = attr(srgbClr, "val");
                                        if (color) borderDef.color = color;
                                    }
                                }
                                // Extract dashStyle from a:prstDash/@val
                                const prstDash = findChild(border, "a:prstDash");
                                if (prstDash) {
                                    const dashVal = attr(prstDash, "val");
                                    if (dashVal) borderDef.dashStyle = dashVal;
                                }
                                borders[borderName] = borderDef;
                            }
                        }
                    }
                    if (Object.keys(borders).length > 0) cell.borders = borders;
                }
            }

            // Text body
            const txBody = findChild(tc, "a:txBody");
            if (txBody) {
                const textBody = parseTextBody(txBody, ctx);
                if (textBody.paragraphs) cell.paragraphs = textBody.paragraphs;
            }

            row.cells.push(cell);
        }

        result.rows.push(row);
    }

    return result;
}
