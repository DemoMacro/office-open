import type { RawElement, MixedChildren } from "@office-open/core";
import { isRaw } from "@office-open/core";
import type { FillOptions } from "@office-open/core/drawingml";
import type { Element } from "@office-open/xml";

export type { RawElement, MixedChildren };
export { isRaw };

export interface PptxDocumentJson {
    slides: SlideJson[];
    slideWidth?: number;
    slideHeight?: number;
    title?: string;
    creator?: string;
    subject?: string;
    keywords?: string;
    description?: string;
    lastModifiedBy?: string;
    $parts?: Record<string, Element>;
}

export interface SlideJson {
    children: SlideChildJson[];
    background?: string | Record<string, unknown>;
    notes?: string | ParagraphJson[];
    transition?: Record<string, unknown>;
}

export type SlideChildJson =
    | ShapeJson
    | PictureJson
    | ConnectorShapeJson
    | TableJson
    | GroupShapeJson
    | RawElement;

export interface ShapeJson {
    $type?: "shape";
    id?: number;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    geometry?: string;
    fill?: FillOptions;
    outline?: ParsedOutlineOptions;
    effects?: ParsedEffectsOptions;
    flipH?: boolean;
    flipV?: boolean;
    rotation?: number;
    text?: string;
    paragraphs?: ParagraphJson[];
    textVertical?: string;
    textAnchor?: string;
    textAutoFit?: string;
    placeholder?: string;
    placeholderIndex?: number;
}

export interface PictureJson {
    $type?: "picture";
    id?: number;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    data?: string;
    type?: string;
    outline?: ParsedOutlineOptions;
    effects?: ParsedEffectsOptions;
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
}

export interface ConnectorShapeJson {
    $type?: "connectorShape";
    id?: number;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    outline?: ParsedOutlineOptions;
    style?: string;
    beginX?: number;
    beginY?: number;
    endX?: number;
    endY?: number;
}

export interface GroupShapeJson {
    $type?: "groupShape";
    id?: number;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
    children?: SlideChildJson[];
}

export interface TableJson {
    $type?: "table";
    id?: number;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    rows: TableRowJson[];
    tableStyle?: Record<string, unknown>;
}

export interface TableRowJson {
    cells: TableCellJson[];
    height?: number;
}

export interface TableCellJson {
    paragraphs?: ParagraphJson[];
    columnSpan?: number;
    rowSpan?: number;
    fill?: FillOptions;
    borders?: Record<string, unknown>;
    verticalAlign?: string;
    margins?: Record<string, unknown>;
    width?: number;
}

export interface ParagraphJson {
    text?: string;
    children?: RunJson[];
    alignment?: string;
    indent?: { level?: number; left?: number };
    spacing?: { before?: number; after?: number; line?: number };
    bullet?: { type?: string; level?: number };
    numbering?: { type?: string; level?: number };
}

export interface RunJson {
    text?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: string;
    font?: string;
    lang?: string;
    fill?: FillOptions;
    hyperlink?: { url: string; tooltip?: string };
    strike?: string;
    baseline?: number;
    spacing?: number;
    capitalization?: string;
    shadow?: boolean;
    outline?: boolean;
    rightToLeft?: boolean;
    noProof?: boolean;
}

// Parser-specific typed interfaces (differ from constructor OutlineOptions/IEffectsOptions)

export interface ParsedOutlineOptions {
    width?: number;
    cap?: string;
    compound?: string;
    alignment?: string;
    color?: string;
    dashStyle?: string;
    round?: boolean;
    bevel?: boolean;
    headEnd?: { type?: string; width?: number; length?: number };
    tailEnd?: { type?: string; width?: number; length?: number };
    join?: string;
}

export interface ParsedEffectsOptions {
    outerShadow?: ParsedShadowOptions;
    innerShadow?: ParsedShadowOptions;
    glow?: ParsedGlowOptions;
    reflection?: ParsedReflectionOptions;
}

export interface ParsedShadowOptions {
    blurRadius?: number;
    distance?: number;
    direction?: number;
    alignment?: string;
    rotateWithShape?: boolean;
    color?: string;
}

export interface ParsedGlowOptions {
    radius?: number;
    color?: string;
}

export interface ParsedReflectionOptions {
    blurRadius?: number;
    distance?: number;
    direction?: number;
    fadeDirection?: number;
    startAlpha?: number;
    endAlpha?: number;
    rotateWithShape?: boolean;
}
