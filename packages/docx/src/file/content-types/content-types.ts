/**
 * Content Types module for Open Packaging Conventions.
 *
 * This module provides the [Content_Types].xml part which defines
 * the content types for all parts in the DOCX package.
 *
 * Reference: http://officeopenxml.com/anatomyofOOXML.php
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import { ContentTypeAttributes } from "./content-types-attributes";
import { createDefault } from "./default/default";
import { createOverride } from "./override/override";

/**
 * Represents the Content Types part of an OPC package.
 *
 * ContentTypes maps file extensions and specific paths to their
 * MIME content types, enabling applications to process each part correctly.
 *
 * Reference: http://officeopenxml.com/anatomyofOOXML.php
 *
 * @example
 * ```typescript
 * const contentTypes = new ContentTypes();
 * contentTypes.addHeader(1); // Add header1.xml
 * contentTypes.addFooter(1); // Add footer1.xml
 * ```
 */
export class ContentTypes extends XmlComponent {
    public constructor() {
        super("Types");

        this.root.push(
            new ContentTypeAttributes({
                xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
            }),
        );

        this.root.push(createDefault("image/png", "png"));
        this.root.push(createDefault("image/jpeg", "jpeg"));
        this.root.push(createDefault("image/jpeg", "jpg"));
        this.root.push(createDefault("image/bmp", "bmp"));
        this.root.push(createDefault("image/gif", "gif"));
        this.root.push(createDefault("image/tiff", "tif"));
        this.root.push(createDefault("image/tiff", "tiff"));
        this.root.push(createDefault("image/x-emf", "emf"));
        this.root.push(createDefault("image/x-wmf", "wmf"));
        this.root.push(createDefault("image/x-icon", "ico"));
        this.root.push(createDefault("image/svg+xml", "svg"));
        this.root.push(
            createDefault("application/vnd.openxmlformats-package.relationships+xml", "rels"),
        );
        this.root.push(createDefault("application/xml", "xml"));
        this.root.push(
            createDefault("application/vnd.openxmlformats-officedocument.obfuscatedFont", "odttf"),
        );

        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "/word/document.xml",
            ),
        );

        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
                "/word/styles.xml",
            ),
        );
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-package.core-properties+xml",
                "/docProps/core.xml",
            ),
        );
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.custom-properties+xml",
                "/docProps/custom.xml",
            ),
        );
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                "/docProps/app.xml",
            ),
        );
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
                "/word/numbering.xml",
            ),
        );
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
                "/word/footnotes.xml",
            ),
        );
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
                "/word/endnotes.xml",
            ),
        );
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
                "/word/settings.xml",
            ),
        );
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
                "/word/comments.xml",
            ),
        );
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
                "/word/fontTable.xml",
            ),
        );
    }

    /**
     * Registers a bibliography part in the content types.
     */
    public addBibliography(): void {
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.bibliography+xml",
                "/word/bibliography.xml",
            ),
        );
    }

    /**
     * Registers a footer part in the content types.
     *
     * @param index - Footer index number (e.g., 1 for footer1.xml)
     */
    public addFooter(index: number): void {
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
                `/word/footer${index}.xml`,
            ),
        );
    }

    /**
     * Registers a header part in the content types.
     *
     * @param index - Header index number (e.g., 1 for header1.xml)
     */
    public addHeader(index: number): void {
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
                `/word/header${index}.xml`,
            ),
        );
    }

    /**
     * Registers a chart part in the content types.
     *
     * @param index - Chart index number (e.g., 1 for charts/chart1.xml)
     */
    public addChart(index: number): void {
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                `/word/charts/chart${index}.xml`,
            ),
        );
    }

    /**
     * Registers a diagram data part in the content types.
     *
     * @param index - Diagram data index number
     */
    public addDiagramData(index: number): void {
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
                `/word/diagrams/data${index}.xml`,
            ),
        );
    }

    /**
     * Registers a diagram layout part in the content types.
     *
     * @param index - Diagram layout index number
     */
    public addDiagramLayout(index: number): void {
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
                `/word/diagrams/layout${index}.xml`,
            ),
        );
    }

    /**
     * Registers a diagram style part in the content types.
     *
     * @param index - Diagram style index number
     */
    public addDiagramStyle(index: number): void {
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
                `/word/diagrams/quickStyle${index}.xml`,
            ),
        );
    }

    /**
     * Registers a diagram colors part in the content types.
     *
     * @param index - Diagram colors index number
     */
    public addDiagramColors(index: number): void {
        this.root.push(
            createOverride(
                "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
                `/word/diagrams/colors${index}.xml`,
            ),
        );
    }

    /**
     * Registers an alternative format chunk part in the content types.
     *
     * @param path - Part path (e.g., "word/afchunks/afchunk1.html")
     * @param contentType - MIME content type
     * @param extension - File extension for the default entry
     */
    public addAltChunk(path: string, contentType: string, _extension: string): void {
        this.root.push(createOverride(contentType, path));
    }
}
