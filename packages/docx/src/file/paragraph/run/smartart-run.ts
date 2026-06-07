/**
 * SmartArtRun module for embedding SmartArt diagrams in WordprocessingML documents.
 *
 * SmartArt data is stored in `word/diagrams/data{n}.xml`.
 * Layout, style, and colors reference Word's built-in definitions.
 *
 * @module
 */
import { Drawing } from "@file/drawing";
import type { Floating } from "@file/drawing";
import type { DocPropertiesOptions } from "@file/drawing/doc-properties/doc-properties";
import type { MediaTransformation } from "@file/media";
import { createTransformation } from "@file/media";
import type { SmartArtData } from "@file/smartart/smartart-collection";
import { createDataModel } from "@file/smartart/tree-to-model";
import type { Context } from "@file/xml-components";

import { Run } from "../run";

/**
 * A tree node for SmartArt data.
 */
export interface SmartArtNode {
  text: string;
  children?: SmartArtNode[];
}

/**
 * Options for creating a SmartArtRun.
 *
 * @publicApi
 */
export interface SmartArtOptions {
  /** Tree-shaped data for the diagram */
  data: {
    nodes: SmartArtNode[];
  };
  /** Display dimensions */
  transformation: MediaTransformation;
  /** Floating positioning */
  floating?: Floating;
  /** Alternative text for accessibility */
  altText?: DocPropertiesOptions;
  /** Layout ID (e.g. "default", "process1", "hierarchy1") */
  layout?: string;
  /** Quick style ID (e.g. "simple1", "moderate1") */
  style?: string;
  /** Color transform ID (e.g. "accent1_2", "colorful1") */
  color?: string;
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
  private smartArtOptions: SmartArtOptions;
  private smartArtKey: string;

  public constructor(options: SmartArtOptions) {
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

    this.extraChildren.push(drawing);
  }

  protected override registerResources(context: Context): void {
    const layoutId = this.smartArtOptions.layout ?? "default";
    const styleId = this.smartArtOptions.style ?? "simple1";
    const colorId = this.smartArtOptions.color ?? "accent1_2";

    const dataModel = createDataModel(this.smartArtOptions.data.nodes, layoutId, styleId, colorId);

    const smartArtData: SmartArtData = {
      dataModel,
      key: this.smartArtKey,
      layout: layoutId,
      style: styleId,
      color: colorId,
    };

    context.file.smartArts.addSmartArt(this.smartArtKey, smartArtData);
  }

  private hashSmartArtData(options: SmartArtOptions): number {
    const data = JSON.stringify(options.data);
    let hash = 0;
    for (let i = 0; i < data.length; i++) {
      const char = data.charCodeAt(i);
      hash = ((hash << 5) - hash + char) | 0;
    }
    return Math.abs(hash);
  }
}
