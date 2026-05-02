/**
 * Content Types module for PPTX packages.
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import { ContentTypeAttributes } from "./content-types-attributes";
import { createDefault } from "./default/default";
import { createOverride } from "./override/override";

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

export class ContentTypes extends XmlComponent {
    public constructor() {
        super("Types");

        this.root.push(
            new ContentTypeAttributes({
                xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
            }),
        );

        // Default content types by extension
        this.root.push(
            createDefault("application/vnd.openxmlformats-package.relationships+xml", "rels"),
        );
        this.root.push(createDefault("application/xml", "xml"));

        // Static overrides
        this.root.push(createOverride(PPTX_MAIN, "/ppt/presentation.xml"));
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-package.core-properties+xml",
                "/docProps/core.xml",
            ),
        );
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                "/docProps/app.xml",
            ),
        );
        this.root.push(createOverride(PPTX_THEME, "/ppt/theme/theme1.xml"));
        this.root.push(createOverride(PPTX_SLIDE_MASTER, "/ppt/slideMasters/slideMaster1.xml"));
        this.root.push(createOverride(PPTX_SLIDE_LAYOUT, "/ppt/slideLayouts/slideLayout1.xml"));
        this.root.push(createOverride(PPTX_PRES_PROPS, "/ppt/presProps.xml"));
        this.root.push(createOverride(PPTX_VIEW_PROPS, "/ppt/viewProps.xml"));
        this.root.push(createOverride(PPTX_TABLE_STYLES, "/ppt/tableStyles.xml"));
    }

    public addSlide(index: number): void {
        this.root.push(createOverride(PPTX_SLIDE, `/ppt/slides/slide${index}.xml`));
    }

    public addNotesSlide(index: number): void {
        this.root.push(createOverride(PPTX_NOTES, `/ppt/notesSlides/notesSlide${index}.xml`));
    }

    public addChart(index: number): void {
        this.root.push(createOverride(PPTX_CHART, `/ppt/charts/chart${index}.xml`));
    }

    public addDiagramData(index: number): void {
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
                `/ppt/diagrams/data${index}.xml`,
            ),
        );
    }

    public addDiagramLayout(index: number): void {
        this.root.push(
            createOverride(
                "http://schemas.microsoft.com/office/2007/relationships/diagramLayout",
                `/ppt/diagrams/layout${index}.xml`,
            ),
        );
    }

    public addDiagramStyle(index: number): void {
        this.root.push(
            createOverride(
                "http://schemas.microsoft.com/office/2007/relationships/diagramStyle",
                `/ppt/diagrams/quickStyle${index}.xml`,
            ),
        );
    }

    public addDiagramColors(index: number): void {
        this.root.push(
            createOverride(
                "http://schemas.microsoft.com/office/2007/relationships/diagramColors",
                `/ppt/diagrams/colors${index}.xml`,
            ),
        );
    }
}
