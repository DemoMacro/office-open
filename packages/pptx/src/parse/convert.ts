import { Picture } from "@file/picture/picture";
import { GroupShape } from "@file/shape/group-shape";
import { ConnectorShape } from "@file/shape/line-shape";
import { Paragraph } from "@file/shape/paragraph/paragraph";
import { Run } from "@file/shape/paragraph/run";
import type { TextCapitalization } from "@file/shape/paragraph/run-properties";
import { Shape } from "@file/shape/shape";
import { Table } from "@file/table/table";
import type { BaseXmlComponent } from "@file/xml-components";
import { RawPassthrough } from "@office-open/core";
import type { FillOptions } from "@office-open/core/drawingml";

import { isRaw } from "./types";
import type {
    SlideChildJson,
    ShapeJson,
    PictureJson,
    ConnectorShapeJson,
    GroupShapeJson,
    TableJson,
    ParagraphJson,
    RunJson,
    ParsedOutlineOptions,
} from "./types";

// ── Constants ──

const EMU_PER_PIXEL = 9525;

// ── Public API ──

/**
 * Convert parsed slide children to constructor-ready BaseXmlComponent[].
 */
export function toSlideChildren(children: SlideChildJson[]): BaseXmlComponent[] {
    return children.map(convertSlideChild);
}

// ── Internal converters ──

function convertSlideChild(child: SlideChildJson): BaseXmlComponent {
    if (isRaw(child)) return new RawPassthrough(child.element);

    switch (child.$type) {
        case "shape":
            return convertShape(child);
        case "picture":
            return convertPicture(child);
        case "connectorShape":
            return convertConnectorShape(child);
        case "table":
            return convertTable(child);
        case "groupShape":
            return convertGroupShape(child);
        default:
            return new RawPassthrough((child as Record<string, unknown>).element as any);
    }
}

function convertShape(json: ShapeJson): Shape {
    const { paragraphs, fill, outline, x, y, width, height, rotation, ...rest } = json as any;
    return new Shape({
        ...rest,
        x: x != null ? Math.round(x / EMU_PER_PIXEL) : undefined,
        y: y != null ? Math.round(y / EMU_PER_PIXEL) : undefined,
        width: width != null ? Math.round(width / EMU_PER_PIXEL) : undefined,
        height: height != null ? Math.round(height / EMU_PER_PIXEL) : undefined,
        rotation: rotation != null ? rotation * 60000 : undefined,
        fill: convertFill(fill),
        outline: convertOutline(outline),
        paragraphs: paragraphs?.map(convertParagraph),
    });
}

function convertPicture(json: PictureJson): Picture {
    const { data, type, x, y, width, height, ...rest } = json as any;
    return new Picture({
        ...rest,
        data: base64ToUint8Array(data),
        type: type as any,
        x: x != null ? Math.round(x / EMU_PER_PIXEL) : undefined,
        y: y != null ? Math.round(y / EMU_PER_PIXEL) : undefined,
        width: width != null ? Math.round(width / EMU_PER_PIXEL) : undefined,
        height: height != null ? Math.round(height / EMU_PER_PIXEL) : undefined,
    });
}

function convertConnectorShape(json: ConnectorShapeJson): ConnectorShape {
    const { outline, x, y, width, height, ...rest } = json as any;
    // Parser outputs x/y/width/height (EMU), constructor expects x1/y1/x2/y2 (pixel)
    const x1 = x != null ? Math.round(x / EMU_PER_PIXEL) : 0;
    const y1 = y != null ? Math.round(y / EMU_PER_PIXEL) : 0;
    const x2 = x != null && width != null ? Math.round((x + width) / EMU_PER_PIXEL) : x1 + 100;
    const y2 = y != null && height != null ? Math.round((y + height) / EMU_PER_PIXEL) : y1 + 100;
    return new ConnectorShape({
        ...rest,
        x1,
        y1,
        x2,
        y2,
        outline: convertOutline(outline),
    });
}

function convertTable(json: TableJson): Table {
    const { rows, ...rest } = json as any;
    return new Table({
        ...rest,
        rows: rows.map(convertTableRow),
    });
}

function convertGroupShape(json: GroupShapeJson): GroupShape {
    const { children, x, y, width, height, rotation, ...rest } = json as any;
    return new GroupShape({
        ...rest,
        x: x != null ? Math.round(x / EMU_PER_PIXEL) : undefined,
        y: y != null ? Math.round(y / EMU_PER_PIXEL) : undefined,
        width: width != null ? Math.round(width / EMU_PER_PIXEL) : undefined,
        height: height != null ? Math.round(height / EMU_PER_PIXEL) : undefined,
        // Group rotation is in degrees (rot / 60000), constructor expects rot value
        rotation: rotation != null ? rotation * 60000 : undefined,
        children: children ? toSlideChildren(children) : [],
    });
}

function convertTableRow(row: any): any {
    return {
        cells: (row.cells ?? []).map(convertTableCell),
        height: row.height,
    };
}

function convertTableCell(cell: any): any {
    return {
        paragraphs: cell.paragraphs?.map(convertParagraph),
        columnSpan: cell.columnSpan,
        rowSpan: cell.rowSpan,
        fill: convertFill(cell.fill),
        borders: convertCellBorders(cell.borders),
        verticalAlign: cell.verticalAlign,
        width: cell.width,
    };
}

function convertParagraph(json: ParagraphJson): Paragraph {
    const { children, text, ...rest } = json as any;
    const runs = children?.map(convertRun) ?? (text ? [new Run({ text })] : undefined);
    return new Paragraph({
        ...rest,
        children: runs,
    });
}

function convertRun(json: RunJson): Run {
    const { fill, underline, strike, capitalization, fontSize, ...rest } = json as any;
    return new Run({
        ...rest,
        fontSize: fontSize != null ? Math.round(fontSize / 100) : undefined,
        fill: convertFill(fill),
        underline: mapUnderline(underline),
        strike: mapStrike(strike),
        capitalization: mapCapitalization(capitalization),
    });
}

// ── Helpers ──

const BORDER_KEY_MAP: Record<string, string> = {
    lnT: "top",
    lnB: "bottom",
    lnL: "left",
    lnR: "right",
    lnH: "insideHorizontal",
    lnV: "insideVertical",
};

function convertCellBorders(borders: Record<string, unknown> | undefined): any {
    if (!borders) return undefined;
    const result: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(borders)) {
        const mappedKey = BORDER_KEY_MAP[key] ?? key;
        result[mappedKey] = value;
    }
    return Object.keys(result).length > 0 ? result : undefined;
}

function convertFill(fill: FillOptions | undefined): any {
    if (!fill || typeof fill === "string") return fill;
    return fill;
}

function convertOutline(outline: ParsedOutlineOptions | undefined): any {
    if (!outline) return undefined;
    const result: Record<string, unknown> = {};
    if (outline.width != null) result.width = outline.width;
    if (outline.color != null) result.color = outline.color;
    if (outline.dashStyle != null) result.dashStyle = mapDash(outline.dashStyle);
    if (outline.cap != null) result.cap = outline.cap;
    if (outline.compound != null) result.compound = outline.compound;
    if (outline.headEnd) result.headEnd = outline.headEnd;
    if (outline.tailEnd) result.tailEnd = outline.tailEnd;
    if (outline.join != null) result.join = outline.join;
    return Object.keys(result).length > 0 ? result : undefined;
}

function base64ToUint8Array(base64: string): Uint8Array {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
}

const UNDERLINE_MAP: Record<string, string> = {
    sng: "SINGLE",
    dbl: "DOUBLE",
    heavy: "HEAVY",
    dotted: "DOTTED",
    dash: "DASH",
    dashDot: "DASH_DOT",
    dashDotDot: "DASH_DOT_DOT",
    sysDash: "SYS_DASH",
    sysDotDash: "SYS_DOT_DASH",
};

const STRIKE_MAP: Record<string, string> = {
    noStrike: "NONE",
    sngStrike: "SINGLE",
    dblStrike: "DOUBLE",
    heavyStrike: "DOUBLE",
};

const CAPITALIZATION_MAP: Record<string, keyof typeof TextCapitalization> = {
    none: "NONE",
    small: "SMALL",
    all: "ALL",
};

const DASH_MAP: Record<string, string> = {
    solid: "SOLID",
    dash: "DASH",
    dashDot: "DASH_DOT",
    dashDotDot: "LG_DASH_DOT",
    dot: "DOT",
    lgDash: "LARGE_DASH",
    sysDash: "SYS_DASH",
    sysDotDash: "SYS_DOT_DASH",
};

function mapUnderline(value: string | undefined): any {
    if (!value) return undefined;
    return UNDERLINE_MAP[value] ?? value;
}

function mapStrike(value: string | undefined): any {
    if (!value) return undefined;
    return STRIKE_MAP[value] ?? value;
}

function mapCapitalization(value: string | undefined): any {
    if (!value) return undefined;
    return CAPITALIZATION_MAP[value] ?? value;
}

function mapDash(value: string | undefined): string | undefined {
    if (!value) return undefined;
    return DASH_MAP[value] ?? value;
}
