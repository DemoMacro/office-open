import type { RawElement, MixedChildren } from "@office-open/core";
import { isRaw } from "@office-open/core";
import type { Element } from "@office-open/xml";

export type { RawElement, MixedChildren };
export { isRaw };

export type FileChildJson =
    | ParagraphJson
    | TableJson
    | ImageRunJson
    | ExternalHyperlinkJson
    | PageBreakJson
    | SdtJson
    | MathJson
    | RawElement;

export type ParagraphChildJson =
    | TextRunJson
    | ImageRunJson
    | ExternalHyperlinkJson
    | PageBreakJson
    | LineBreakJson
    | ColumnBreakJson
    | TabJson
    | BookmarkJson
    | SdtRunJson
    | MathJson
    | FieldJson
    | RawElement;

export interface ParagraphJson {
    $type?: "paragraph";
    children?: MixedChildren<ParagraphChildJson>;
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
    border?: Record<string, unknown>;
    shading?: { fill?: string; color?: string; type?: string };
    thematicBreak?: boolean;
    pageBreakBefore?: boolean;
    keepNext?: boolean;
    keepLines?: boolean;
    widowControl?: boolean;
    outlineLevel?: number;
    tabStops?: Array<{ position: number; type?: string; leader?: string }>;
    suppressLineNumbers?: boolean;
    contextualSpacing?: boolean;
    mirrorIndents?: boolean;
    run?: Record<string, unknown>;
}

export interface TextRunJson {
    $type?: "textRun";
    text?: string;
    deletedText?: boolean;
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
    font?:
        | string
        | { ascii?: string; hAnsi?: string; eastAsia?: string; cs?: string; hint?: string };
    highlight?: string;
    characterSpacing?: number;
    rightToLeft?: boolean;
    shading?: { fill?: string; color?: string };
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
    boldCs?: boolean;
    italicCs?: boolean;
    lang?: string;
}

export interface ImageRunJson {
    $type?: "imageRun";
    type: string;
    data: string;
    transformation?: {
        width?: number;
        height?: number;
    };
    floating?: Record<string, unknown>;
    altText?: { name?: string; description?: string };
}

export interface ExternalHyperlinkJson {
    $type?: "externalHyperlink";
    link: string;
    children?: MixedChildren<ParagraphChildJson>;
    tooltip?: string;
}

export interface PageBreakJson {
    $type?: "pageBreak";
}

export interface LineBreakJson {
    $type?: "lineBreak";
}

export interface ColumnBreakJson {
    $type?: "columnBreak";
}

export interface TabJson {
    $type?: "tab";
}

export interface BookmarkJson {
    $type?: "bookmark";
    name?: string;
    element?: Element;
}

export interface SdtJson {
    $type?: "sdt";
    sdtPr?: Element;
    content?: MixedChildren<FileChildJson>;
    element?: Element;
}

export interface SdtRunJson {
    $type?: "sdtRun";
    sdtPr?: Element;
    content?: MixedChildren<ParagraphChildJson>;
    element?: Element;
}

export interface MathJson {
    $type?: "math";
    element: Element;
}

export interface FieldJson {
    $type?: "field";
    instruction?: string;
    locked?: boolean;
    dirty?: boolean;
    children?: MixedChildren<ParagraphChildJson>;
}

export interface TableJson {
    $type?: "table";
    rows: TableRowJson[];
    width?: { size?: number; type?: string };
    columnWidths?: number[];
    style?: string;
    alignment?: string;
    borders?: Record<string, unknown>;
    cellMargins?: Record<string, unknown>;
    indentation?: Record<string, unknown>;
    layout?: string;
}

export interface TableRowJson {
    cells: TableCellJson[];
    height?: { value: number; rule?: string };
    header?: boolean;
}

export interface TableCellJson {
    children: MixedChildren<ParagraphJson>;
    width?: { size?: number; type?: string };
    shading?: { fill?: string };
    verticalAlign?: string;
    borders?: Record<string, unknown>;
    columnSpan?: number;
    rowSpan?: number;
    noWrap?: boolean;
    textDirection?: string;
    margins?: Record<string, unknown>;
}

export interface SectionJson {
    properties?: Record<string, unknown>;
    children: FileChildJson[];
    headers?: Record<string, { children: FileChildJson[] }>;
    footers?: Record<string, { children: FileChildJson[] }>;
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
    numbering?: import("./numbering").ParsedNumberingConfig[];
    $parts?: Record<string, Element>;
    footnotes?: FootnoteEntry[];
    endnotes?: FootnoteEntry[];
    comments?: CommentEntry[];
}

export interface FootnoteEntry {
    id: string;
    type?: "separator" | "continuationSeparator" | "normal";
    children?: MixedChildren<FileChildJson>;
}

export interface CommentEntry {
    id: string;
    author?: string;
    date?: string;
    initials?: string;
    children?: MixedChildren<FileChildJson>;
}
