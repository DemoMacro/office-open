import type { DataType } from "../util/data-type";

/** Structured failure raised when ODF-only chart styling enters an OOXML writer. */
export class OdfChartStyleOverlayError extends Error {
  constructor(
    message: string,
    readonly part: string,
    readonly path: string,
    readonly name: string,
    readonly reason: string,
  ) {
    super(message);
    this.name = "OdfChartStyleOverlayError";
  }
}

/** Attributes of an ODF style-properties child, keyed by expanded XML name. */
export interface OdfStylePropertiesOptions {
  name: string;
  attributes: Record<string, string>;
}

/** A named ODF style definition retained for an ODF chart subdocument. */
export interface OdfChartStyleResource {
  name: string;
  displayName?: string;
  attributes: Record<string, string>;
  properties: OdfStylePropertiesOptions[];
}

/** Base fields for a named ODF drawing style resource. */
export interface OdfDrawingStyleResource {
  name: string;
  displayName?: string;
  attributes: Record<string, string>;
}

/** An ODF fill image and its embedded media bytes, when available. */
export interface OdfFillImageResource extends OdfDrawingStyleResource {
  data?: DataType;
}

/** ODF style resources retained in source order for an ODF chart subdocument. */
export interface OdfChartStyleResources {
  styles: OdfChartStyleResource[];
  gradients: OdfDrawingStyleResource[];
  hatches: OdfDrawingStyleResource[];
  fillImages: OdfFillImageResource[];
  strokeDashes: OdfDrawingStyleResource[];
}

export type OdfChartFillOptions =
  | { type: "none" }
  | { type: "solid"; color: string }
  | { type: "gradient"; name: string }
  | { type: "hatch"; name: string }
  | {
      type: "bitmap";
      name: string;
      repeat?: "no-repeat" | "repeat" | "stretch";
      width?: string;
      height?: string;
      refPoint?: string;
      refPointX?: string;
      refPointY?: string;
      tileRepeatOffset?: string;
      opacityName?: string;
    };

export type OdfChartOutlineOptions =
  | { type: "none" }
  | { type: "solid"; color?: string }
  | { type: "dash"; name: string };

/** ODF-only chart shape style fields retained alongside canonical projections. */
export interface OdfChartStyleOptions {
  fill?: OdfChartFillOptions;
  outline?: OdfChartOutlineOptions;
}
