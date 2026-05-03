/**
 * SmartArtRun module for embedding SmartArt diagrams in WordprocessingML documents.
 *
 * SmartArt data is stored in `word/diagrams/data{n}.xml`.
 * Layout, style, and colors reference Word's built-in definitions.
 *
 * @module
 */
import { Drawing } from "@file/drawing";
import type { IFloating } from "@file/drawing";
import type { DocPropertiesOptions } from "@file/drawing/doc-properties/doc-properties";
import type { IMediaTransformation } from "@file/media";
import { createTransformation } from "@file/media";
import type { ISmartArtData } from "@file/smartart/smartart-collection";
import { createDataModel } from "@file/smartart/tree-to-model";
import type { IContext, IXmlableObject } from "@file/xml-components";

import { Run } from "../run";

/**
 * A tree node for SmartArt data.
 */
export interface ISmartArtNode {
    readonly text: string;
    readonly children?: readonly ISmartArtNode[];
}

/**
 * Options for creating a SmartArtRun.
 *
 * @publicApi
 */
export interface ISmartArtOptions {
    /** Tree-shaped data for the diagram */
    readonly data: {
        readonly nodes: readonly ISmartArtNode[];
    };
    /** Display dimensions */
    readonly transformation: IMediaTransformation;
    /** Floating positioning */
    readonly floating?: IFloating;
    /** Alternative text for accessibility */
    readonly altText?: DocPropertiesOptions;
    /** Layout ID (e.g. "default", "process1", "hierarchy1") */
    readonly layout?: string;
    /** Quick style ID (e.g. "simple1", "moderate1") */
    readonly style?: string;
    /** Color transform ID (e.g. "accent1_2", "colorful1") */
    readonly color?: string;
}

/**
 * Represents an embedded SmartArt diagram in a WordprocessingML document.
 *
 * @publicApi
 *
 * @example
 * ```typescript
 * new SmartArtRun({
 *   data: {
 *     nodes: [
 *       { text: "Main", children: [
 *         { text: "Sub 1" },
 *         { text: "Sub 2" },
 *       ]},
 *     ],
 *   },
 *   transformation: { width: 500, height: 300 },
 * });
 * ```
 */
export class SmartArtRun extends Run {
    private readonly smartArtOptions: ISmartArtOptions;
    private readonly smartArtKey: string;

    public constructor(options: ISmartArtOptions) {
        super({});
        this.smartArtOptions = options;

        // Generate a unique key
        const hash = this.hashSmartArtData(options);
        this.smartArtKey = `smartart_${hash}`;

        // Create media data for the drawing system
        const mediaData = {
            smartArtKey: this.smartArtKey,
            transformation: createTransformation(options.transformation),
            type: "smartart" as const,
        };

        const drawing = new Drawing(mediaData, {
            docProperties: options.altText,
            floating: options.floating,
        });

        this.root.push(drawing);
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const layoutId = this.smartArtOptions.layout ?? "default";
        const styleId = this.smartArtOptions.style ?? "simple1";
        const colorId = this.smartArtOptions.color ?? "accent1_2";

        const dataModel = createDataModel(this.smartArtOptions.data.nodes, layoutId, styleId, colorId);

        const smartArtData: ISmartArtData = {
            dataModel,
            key: this.smartArtKey,
            layout: layoutId,
            style: styleId,
            color: colorId,
        };

        context.file.SmartArts.addSmartArt(this.smartArtKey, smartArtData);

        return super.prepForXml(context);
    }

    private hashSmartArtData(options: ISmartArtOptions): number {
        const data = JSON.stringify(options.data);
        let hash = 0;
        for (let i = 0; i < data.length; i++) {
            const char = data.charCodeAt(i);
            hash = ((hash << 5) - hash + char) | 0;
        }
        return Math.abs(hash);
    }
}
