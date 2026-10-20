import type { DataType } from "../util/data-type";

/** Structured failure raised when a chart style cannot survive format conversion. */
export class ChartStyleCompatibilityError extends Error {
  constructor(
    message: string,
    readonly part: string,
    readonly path: string,
    readonly name: string,
    readonly reason: string,
  ) {
    super(message);
    this.name = "ChartStyleCompatibilityError";
  }
}

/** Attributes of a style-properties child, keyed by expanded XML name. */
export interface StylePropertiesOptions {
  name: string;
  attributes: Record<string, string>;
}

/** A named style definition retained beside a chart subdocument. */
export interface ChartStyleResource {
  name: string;
  displayName?: string;
  attributes: Record<string, string>;
  properties: StylePropertiesOptions[];
}

/** Base fields for a named ODF drawing style resource. */
export interface DrawingStyleResource {
  name: string;
  displayName?: string;
  attributes: Record<string, string>;
}

/** An ODF fill image and its embedded media bytes, when available. */
export interface FillImageResource extends DrawingStyleResource {
  data?: DataType;
}

/** Named style resources retained beside canonical chart projections. */
export interface ChartStyleResources {
  styles: ChartStyleResource[];
  gradients: DrawingStyleResource[];
  hatches: DrawingStyleResource[];
  fillImages: FillImageResource[];
  strokeDashes: DrawingStyleResource[];
}
