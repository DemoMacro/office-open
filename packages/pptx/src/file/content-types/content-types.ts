/**
 * Content Types module for PPTX packages.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

const PPTX_MAIN =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const PPTX_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const PPTX_SLIDE_MASTER =
    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
const PPTX_SLIDE_LAYOUT =
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const PPTX_THEME = "application/vnd.openxmlformats-officedocument.theme+xml";
const PPTX_NOTES = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const PPTX_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const PPTX_PRES_PROPS =
    "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";
const PPTX_VIEW_PROPS =
    "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";
const PPTX_TABLE_STYLES =
    "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";

type EntryType = "Default" | "Override";

interface ContentEntry {
    readonly type: EntryType;
    readonly contentType: string;
    readonly key: string;
}

const STATIC_ENTRIES: readonly ContentEntry[] = [
    {
        type: "Default",
        contentType: "application/vnd.openxmlformats-package.relationships+xml",
        key: "rels",
    },
    { type: "Default", contentType: "application/xml", key: "xml" },
    { type: "Default", contentType: "image/png", key: "png" },
    { type: "Default", contentType: "image/jpeg", key: "jpeg" },
    { type: "Default", contentType: "video/mp4", key: "mp4" },
    { type: "Override", contentType: PPTX_MAIN, key: "/ppt/presentation.xml" },
    {
        type: "Override",
        contentType: "application/vnd.openxmlformats-package.core-properties+xml",
        key: "/docProps/core.xml",
    },
    {
        type: "Override",
        contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        key: "/docProps/app.xml",
    },
    { type: "Override", contentType: PPTX_THEME, key: "/ppt/theme/theme1.xml" },
    { type: "Override", contentType: PPTX_SLIDE_MASTER, key: "/ppt/slideMasters/slideMaster1.xml" },
    { type: "Override", contentType: PPTX_SLIDE_LAYOUT, key: "/ppt/slideLayouts/slideLayout1.xml" },
    { type: "Override", contentType: PPTX_PRES_PROPS, key: "/ppt/presProps.xml" },
    { type: "Override", contentType: PPTX_VIEW_PROPS, key: "/ppt/viewProps.xml" },
    { type: "Override", contentType: PPTX_TABLE_STYLES, key: "/ppt/tableStyles.xml" },
];

const STATIC_CHILDREN: IXmlableObject[] = [
    { _attr: { xmlns: "http://schemas.openxmlformats.org/package/2006/content-types" } },
    ...STATIC_ENTRIES.map((e) => {
        if (e.type === "Default") {
            return { Default: { _attr: { ContentType: e.contentType, Extension: e.key } } };
        }
        return { Override: { _attr: { ContentType: e.contentType, PartName: e.key } } };
    }),
];

export class ContentTypes extends BaseXmlComponent {
    private readonly dynamicEntries: ContentEntry[] = [];

    public constructor() {
        super("Types");
    }

    public addSlide(index: number): void {
        this.dynamicEntries.push({
            type: "Override",
            contentType: PPTX_SLIDE,
            key: `/ppt/slides/slide${index}.xml`,
        });
    }

    public addNotesSlide(index: number): void {
        this.dynamicEntries.push({
            type: "Override",
            contentType: PPTX_NOTES,
            key: `/ppt/notesSlides/notesSlide${index}.xml`,
        });
    }

    public addChart(index: number): void {
        this.dynamicEntries.push({
            type: "Override",
            contentType: PPTX_CHART,
            key: `/ppt/charts/chart${index}.xml`,
        });
    }

    public addDiagramData(index: number): void {
        this.dynamicEntries.push({
            type: "Override",
            contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
            key: `/ppt/diagrams/data${index}.xml`,
        });
    }

    public addDiagramLayout(index: number): void {
        this.dynamicEntries.push({
            type: "Override",
            contentType:
                "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
            key: `/ppt/diagrams/layout${index}.xml`,
        });
    }

    public addDiagramStyle(index: number): void {
        this.dynamicEntries.push({
            type: "Override",
            contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
            key: `/ppt/diagrams/quickStyle${index}.xml`,
        });
    }

    public addDiagramColors(index: number): void {
        this.dynamicEntries.push({
            type: "Override",
            contentType:
                "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
            key: `/ppt/diagrams/colors${index}.xml`,
        });
    }

    public addDiagramDrawing(index: number): void {
        this.dynamicEntries.push({
            type: "Override",
            contentType: "application/vnd.ms-office.drawingml.diagramDrawing+xml",
            key: `/ppt/diagrams/drawing${index}.xml`,
        });
    }

    public override prepForXml(_context: IContext): IXmlableObject {
        const children = [...STATIC_CHILDREN];
        for (const e of this.dynamicEntries) {
            if (e.type === "Default") {
                children.push({
                    Default: { _attr: { ContentType: e.contentType, Extension: e.key } },
                });
            } else {
                children.push({
                    Override: { _attr: { ContentType: e.contentType, PartName: e.key } },
                });
            }
        }
        return { Types: children };
    }
}
