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
}

export interface SlideJson {
    children: SlideChildJson[];
    background?: string | Record<string, unknown>;
    notes?: string;
    transition?: Record<string, unknown>;
    [key: string]: unknown;
}

export type SlideChildJson = ShapeJson | PictureJson | ConnectorShapeJson | TableJson;

export interface ShapeJson {
    $type: "shape";
    id?: number;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    geometry?: string;
    fill?: Record<string, unknown>;
    outline?: Record<string, unknown>;
    effects?: Record<string, unknown>;
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
    [key: string]: unknown;
}

export interface PictureJson {
    $type: "picture";
    id?: number;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    data?: string;
    type?: string;
    outline?: Record<string, unknown>;
    effects?: Record<string, unknown>;
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
    [key: string]: unknown;
}

export interface ConnectorShapeJson {
    $type: "connectorShape";
    id?: number;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    outline?: Record<string, unknown>;
    style?: string;
    beginX?: number;
    beginY?: number;
    endX?: number;
    endY?: number;
    [key: string]: unknown;
}

export interface TableJson {
    $type: "table";
    id?: number;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    rows: TableRowJson[];
    [key: string]: unknown;
}

export interface TableRowJson {
    cells: TableCellJson[];
    height?: number;
    [key: string]: unknown;
}

export interface TableCellJson {
    paragraphs?: ParagraphJson[];
    columnSpan?: number;
    rowSpan?: number;
    fill?: Record<string, unknown>;
    borders?: Record<string, unknown>;
    verticalAlign?: string;
    margins?: Record<string, unknown>;
    width?: number;
    [key: string]: unknown;
}

export interface ParagraphJson {
    text?: string;
    children?: RunJson[];
    alignment?: string;
    indent?: { level?: number; left?: number };
    spacing?: { before?: number; after?: number; line?: number };
    bullet?: { type?: string; level?: number };
    numbering?: { type?: string; level?: number };
    [key: string]: unknown;
}

export interface RunJson {
    text?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: string;
    font?: string;
    lang?: string;
    fill?: Record<string, unknown>;
    hyperlink?: { url: string; tooltip?: string };
    strike?: string;
    baseline?: number;
    spacing?: number;
    capitalization?: string;
    shadow?: boolean;
    outline?: boolean;
    rightToLeft?: boolean;
    noProof?: boolean;
    [key: string]: unknown;
}
