import type { IPropertiesOptions } from "@file/core-properties/properties";
import { Footer, Header } from "@file/header";
import { PageBreak, ColumnBreak } from "@file/paragraph/formatting";
import { ExternalHyperlink } from "@file/paragraph/links";
import { Paragraph } from "@file/paragraph/paragraph";
import { Run } from "@file/paragraph/run";
import { Tab } from "@file/paragraph/run/empty-children";
import { ImageRun } from "@file/paragraph/run/image-run";
import { SimpleField } from "@file/paragraph/run/simple-field";
import { Table } from "@file/table/table";
import { TableCell } from "@file/table/table-cell/table-cell";
import { TableRow } from "@file/table/table-row/table-row";
import type { BaseXmlComponent } from "@file/xml-components";
import { RawPassthrough } from "@office-open/core";

import { isRaw } from "./types";
import type {
    DocxDocumentJson,
    SectionJson,
    FileChildJson,
    ParagraphChildJson,
    ParagraphJson,
    TextRunJson,
    ImageRunJson,
    ExternalHyperlinkJson,
    BookmarkJson,
    SdtJson,
    SdtRunJson,
    FieldJson,
    TableJson,
    TableRowJson,
    TableCellJson,
} from "./types";

// ── Public API ──

/**
 * Convert parsed section children to constructor-ready FileChild[].
 */
export function toSectionChildren(children: FileChildJson[]): BaseXmlComponent[] {
    return children.map(convertFileChild);
}

/**
 * Convert parsed paragraph children to constructor-ready ParagraphChild[].
 */
export function toParagraphChildren(children: ParagraphChildJson[]): BaseXmlComponent[] {
    return children.map(convertParagraphChild);
}

/**
 * Convert parsed DocxDocumentJson to constructor-ready Document options.
 * Handles numbering, headers/footers, and all section properties.
 */
export function toDocumentOptions(json: DocxDocumentJson): IPropertiesOptions {
    return {
        ...(json.title && { title: json.title }),
        ...(json.creator && { creator: json.creator }),
        ...(json.subject && { subject: json.subject }),
        ...(json.description && { description: json.description }),
        ...(json.keywords && { keywords: json.keywords }),
        ...(json.lastModifiedBy && { lastModifiedBy: json.lastModifiedBy }),
        ...(json.revision && { revision: json.revision }),
        ...(json.numbering &&
            json.numbering.length > 0 && { numbering: { config: json.numbering } }),
        sections: json.sections.map(convertSection),
    } as unknown as IPropertiesOptions;
}

// ── Section converter ──

function convertSection(section: SectionJson): Record<string, unknown> {
    const props = { ...section.properties };
    delete props.headerRefs;
    delete props.footerRefs;

    const result: Record<string, unknown> = {
        properties: props,
        children: toSectionChildren(section.children),
    };

    if (section.headers) {
        const headers: Record<string, unknown> = {};
        for (const [type, content] of Object.entries(section.headers)) {
            headers[type] = new Header({
                children: toSectionChildren(content.children) as any,
            });
        }
        result.headers = headers;
    }

    if (section.footers) {
        const footers: Record<string, unknown> = {};
        for (const [type, content] of Object.entries(section.footers)) {
            footers[type] = new Footer({
                children: toSectionChildren(content.children) as any,
            });
        }
        result.footers = footers;
    }

    return result;
}

// ── File-level converters ──

function convertFileChild(child: FileChildJson): BaseXmlComponent {
    if (isRaw(child)) return new RawPassthrough(child.element);

    switch (child.$type) {
        case "paragraph":
            return convertParagraph(child);
        case "table":
            return convertTable(child);
        case "imageRun":
            return convertImageRun(child);
        case "externalHyperlink":
            return convertExternalHyperlink(child);
        case "pageBreak":
            return new PageBreak();
        case "sdt":
            return convertSdt(child);
        case "math":
            return new RawPassthrough(child.element);
        default:
            return new RawPassthrough((child as Record<string, unknown>).element as any);
    }
}

// ── Paragraph-level converters ──

function convertParagraphChild(child: ParagraphChildJson): BaseXmlComponent {
    if (isRaw(child)) return new RawPassthrough(child.element);

    switch (child.$type) {
        case "textRun":
            return convertRun(child);
        case "imageRun":
            return convertImageRun(child);
        case "externalHyperlink":
            return convertExternalHyperlink(child);
        case "pageBreak":
            return new PageBreak();
        case "lineBreak":
            return new ColumnBreak();
        case "columnBreak":
            return new ColumnBreak();
        case "tab":
            return new Tab();
        case "bookmark":
            return new RawPassthrough((child as BookmarkJson).element!);
        case "sdtRun":
            return convertSdtRun(child);
        case "math":
            return new RawPassthrough((child as any).element);
        case "field":
            return convertField(child);
        default:
            return new RawPassthrough((child as Record<string, unknown>).element as any);
    }
}

function convertParagraph(json: ParagraphJson): Paragraph {
    const { children, text, ...rest } = json as any;
    const runs = children ? toParagraphChildren(children) : text ? [new Run({ text })] : undefined;
    return new Paragraph({
        ...rest,
        children: runs,
    });
}

function convertRun(json: TextRunJson): Run {
    const { underline, strike, doubleStrike, size, sizeCs, ...rest } = json as any;
    return new Run({
        ...rest,
        underline: convertUnderline(underline),
        strike: strike ? true : undefined,
        doubleStrike: doubleStrike ? true : undefined,
        size: size,
        sizeComplexScript: sizeCs,
    });
}

const EMU_PER_PIXEL = 9525;

function convertImageRun(json: ImageRunJson): ImageRun {
    const { data, type, transformation, altText } = json as any;
    const convertedTransform = transformation
        ? {
              width: Math.round(transformation.width / EMU_PER_PIXEL),
              height: Math.round(transformation.height / EMU_PER_PIXEL),
              ...(transformation.flip && { flip: transformation.flip }),
              ...(transformation.offset && { offset: transformation.offset }),
          }
        : { width: 100, height: 100 };
    const imageData =
        typeof data === "string" ? Uint8Array.from(atob(data), (c) => c.charCodeAt(0)) : data;
    return new ImageRun({
        data: imageData,
        type: type as "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf",
        transformation: convertedTransform,
        altText,
    });
}

function convertExternalHyperlink(json: ExternalHyperlinkJson): ExternalHyperlink {
    const { children, link, tooltip } = json as any;
    return new ExternalHyperlink({
        link,
        tooltip,
        children: children ? (toParagraphChildren(children) as any) : [new Run({ text: link })],
    });
}

function convertSdt(json: SdtJson): BaseXmlComponent {
    return new RawPassthrough(json.element!);
}

function convertSdtRun(json: SdtRunJson): BaseXmlComponent {
    return new RawPassthrough(json.element!);
}

function convertField(json: FieldJson): SimpleField {
    const { instruction, children } = json as any;
    const cachedText =
        children
            ?.map((c: ParagraphChildJson) => {
                if ((c as Record<string, unknown>).$type === "textRun" && (c as TextRunJson).text)
                    return (c as TextRunJson).text;
                return undefined;
            })
            .filter(Boolean)
            .join("") ?? undefined;
    return new SimpleField(instruction ?? "", cachedText);
}

function convertTable(json: TableJson): Table {
    const { rows, ...rest } = json as any;
    return new Table({
        ...rest,
        rows: rows.map(convertTableRow),
    });
}

function convertTableRow(row: TableRowJson): TableRow {
    const { cells, height, ...rest } = row as any;
    return new TableRow({
        children: cells.map(convertTableCell),
        ...(height && { height }),
        ...rest,
    });
}

function convertTableCell(cell: TableCellJson): TableCell {
    const { children, ...rest } = cell as any;
    return new TableCell({
        ...rest,
        children: children ? toSectionChildren(children) : [new Paragraph({ text: "" })],
    });
}

// ── Helpers ──

function convertUnderline(underline: Record<string, unknown> | string | undefined): any {
    if (!underline) return undefined;
    if (typeof underline === "string") return { type: underline };
    return underline;
}
