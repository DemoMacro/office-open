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

// ── Helpers ──

// Type assertion for cross-boundary conversion (parse JSON → constructor options).
// The two type systems are intentionally separate; this is the single bridging point.
function bridge<T>(value: unknown): T {
    return value as T;
}

// ── Public API ──

export function toSectionChildren(children: FileChildJson[]): BaseXmlComponent[] {
    return children.map(convertFileChild);
}

export function toParagraphChildren(children: ParagraphChildJson[]): BaseXmlComponent[] {
    return children.map(convertParagraphChild);
}

export function toDocumentOptions(json: DocxDocumentJson): IPropertiesOptions {
    return bridge<IPropertiesOptions>({
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
    });
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
                children: toSectionChildren(content.children) as Paragraph[],
            });
        }
        result.headers = headers;
    }

    if (section.footers) {
        const footers: Record<string, unknown> = {};
        for (const [type, content] of Object.entries(section.footers)) {
            footers[type] = new Footer({
                children: toSectionChildren(content.children) as Paragraph[],
            });
        }
        result.footers = footers;
    }

    return result;
}

// ── File-level converters ──

export function convertFileChild(child: FileChildJson): BaseXmlComponent {
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
            return new RawPassthrough(bridge((child as Record<string, unknown>).element));
    }
}

// ── Paragraph-level converters ──

export function convertParagraphChild(child: ParagraphChildJson): BaseXmlComponent {
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
            return new Run({ children: [new Tab()] });
        case "bookmark":
            return new RawPassthrough(bridge((child as BookmarkJson).element));
        case "sdtRun":
            return convertSdtRun(child);
        case "math": {
            const el = (child as { element?: unknown }).element;
            return new RawPassthrough(bridge(el));
        }
        case "field":
            return convertField(child);
        default:
            return new RawPassthrough(bridge((child as Record<string, unknown>).element));
    }
}

export function convertParagraph(json: ParagraphJson): Paragraph {
    const { children, ...rest } = json;
    const runs = children ? toParagraphChildren(children) : undefined;
    return new Paragraph(bridge({ ...rest, children: runs }));
}

export function convertRun(json: TextRunJson): Run {
    const { underline, strike, doubleStrike, ...rest } = json;
    return new Run(
        bridge({
            ...rest,
            underline: convertUnderline(underline),
            strike: strike ? true : undefined,
            doubleStrike: doubleStrike ? true : undefined,
        }),
    );
}

const EMU_PER_PIXEL = 9525;

export function convertImageRun(json: ImageRunJson): ImageRun {
    const { data, type, transformation, altText } = json;
    const width = transformation?.width ?? 100;
    const height = transformation?.height ?? 100;
    const imageData =
        typeof data === "string" ? Uint8Array.from(atob(data), (c) => c.charCodeAt(0)) : data;
    return new ImageRun(
        bridge({
            data: imageData,
            type,
            transformation: {
                width: Math.round(width / EMU_PER_PIXEL),
                height: Math.round(height / EMU_PER_PIXEL),
            },
            altText,
        }),
    );
}

export function convertExternalHyperlink(json: ExternalHyperlinkJson): ExternalHyperlink {
    const { children, link, tooltip } = json;
    return new ExternalHyperlink(
        bridge({
            link,
            tooltip,
            children: children ? toParagraphChildren(children) : [new Run({ text: link })],
        }),
    );
}

export function convertSdt(json: SdtJson): BaseXmlComponent {
    return new RawPassthrough(json.element!);
}

export function convertSdtRun(json: SdtRunJson): BaseXmlComponent {
    return new RawPassthrough(json.element!);
}

export function convertField(json: FieldJson): SimpleField {
    const { instruction, children } = json;
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

export function convertTable(json: TableJson): Table {
    const { rows, ...rest } = json;
    return new Table(bridge({ ...rest, rows: rows.map(convertTableRow) }));
}

export function convertTableRow(row: TableRowJson): TableRow {
    const { cells, height, ...rest } = row;
    return new TableRow(
        bridge({
            ...rest,
            children: cells.map(convertTableCell),
            ...(height && { height }),
        }),
    );
}

export function convertTableCell(cell: TableCellJson): TableCell {
    const { children, ...rest } = cell;
    return new TableCell(
        bridge({
            ...rest,
            children: children ? toSectionChildren(children) : [new Paragraph({ text: "" })],
        }),
    );
}

export function convertUnderline(
    underline: { type?: string; color?: string } | string | undefined,
): { type?: string; color?: string } | undefined {
    if (!underline) return undefined;
    if (typeof underline === "string") return { type: underline };
    return underline;
}
