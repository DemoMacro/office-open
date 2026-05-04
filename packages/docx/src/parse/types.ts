export type FileChildJson =
    | ParagraphJson
    | TableJson
    | ImageRunJson
    | ExternalHyperlinkJson
    | PageBreakJson;

export type ParagraphChildJson =
    | TextRunJson
    | ImageRunJson
    | ExternalHyperlinkJson
    | PageBreakJson
    | TabJson
    | BookmarkJson;

export interface ParagraphJson {
    $type: "paragraph";
    children?: ParagraphChildJson[];
    text?: string;
    alignment?: string;
    spacing?: {
        before?: number;
        after?: number;
        line?: number;
        lineRule?: string;
    };
    indent?: {
        left?: number;
        right?: number;
        firstLine?: number;
        hanging?: number;
    };
    heading?: string;
    style?: string;
    bullet?: { level: number };
    numbering?: { reference: string; level: number; instance?: number };
    border?: {
        top?: Record<string, unknown>;
        bottom?: Record<string, unknown>;
        left?: Record<string, unknown>;
        right?: Record<string, unknown>;
    };
    shading?: { fill?: string; type?: string };
    thematicBreak?: boolean;
    pageBreakBefore?: boolean;
    keepNext?: boolean;
    keepLines?: boolean;
    widowControl?: boolean;
    outlineLevel?: number;
    run?: Record<string, unknown>;
    [key: string]: unknown;
}

export interface TextRunJson {
    $type: "textRun";
    text?: string;
    bold?: boolean;
    italics?: boolean;
    underline?: { type?: string; color?: string };
    strike?: boolean;
    doubleStrike?: boolean;
    smallCaps?: boolean;
    allCaps?: boolean;
    subScript?: boolean;
    superScript?: boolean;
    size?: number;
    sizeComplexScript?: boolean | number;
    color?: string;
    font?: string;
    highlight?: string;
    characterSpacing?: number;
    rightToLeft?: boolean;
    shading?: { fill?: string };
    kern?: number;
    position?: string;
    emboss?: boolean;
    imprint?: boolean;
    shadow?: boolean;
    outline?: boolean;
    vanish?: boolean;
    noProof?: boolean;
    effect?: string;
    math?: boolean;
    [key: string]: unknown;
}

export interface ImageRunJson {
    $type: "imageRun";
    type: string;
    data: string;
    transformation?: {
        width?: number;
        height?: number;
    };
    floating?: Record<string, unknown>;
    altText?: { name?: string; description?: string };
    [key: string]: unknown;
}

export interface ExternalHyperlinkJson {
    $type: "externalHyperlink";
    link: string;
    children?: ParagraphChildJson[];
    tooltip?: string;
    [key: string]: unknown;
}

export interface PageBreakJson {
    $type: "pageBreak";
}

export interface TabJson {
    $type: "tab";
}

export interface BookmarkJson {
    $type: "bookmark";
    [key: string]: unknown;
}

export interface TableJson {
    $type: "table";
    rows: TableRowJson[];
    width?: { size?: number; type?: string };
    columnWidths?: number[];
    style?: string;
    alignment?: string;
    borders?: Record<string, unknown>;
    [key: string]: unknown;
}

export interface TableRowJson {
    cells: TableCellJson[];
    [key: string]: unknown;
}

export interface TableCellJson {
    children: ParagraphJson[];
    width?: { size?: number; type?: string };
    shading?: { fill?: string };
    verticalAlign?: string;
    borders?: Record<string, unknown>;
    columnSpan?: number;
    rowSpan?: number;
    [key: string]: unknown;
}

export interface SectionJson {
    properties?: Record<string, unknown>;
    children: FileChildJson[];
    headers?: {
        default?: { children: ParagraphJson[] };
        first?: { children: ParagraphJson[] };
        even?: { children: ParagraphJson[] };
    };
    footers?: {
        default?: { children: ParagraphJson[] };
        first?: { children: ParagraphJson[] };
        even?: { children: ParagraphJson[] };
    };
}

export interface DocxDocumentJson {
    sections: SectionJson[];
    title?: string;
    subject?: string;
    creator?: string;
    keywords?: string;
    description?: string;
    lastModifiedBy?: string;
    revision?: string;
    [key: string]: unknown;
}
