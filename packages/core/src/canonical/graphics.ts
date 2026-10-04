import type { ChartSpaceOptions } from "../chart";

/** Image graphic anchored to a spreadsheet cell. */
export interface CellImageGraphicOptions {
  /** Image graphic discriminant. */
  type: "image";
  /** Image package URI, without the leading slash. */
  href: string;
  /** Image bytes encoded with base64. */
  data: string;
  /** Accessible image name. */
  name?: string;
  /** Accessible image description. */
  description?: string;
  /** Drawing identifier. */
  id?: string;
  /** Package-local graphic style reference. */
  styleName?: string;
  /** Text style reference for labels attached to the frame. */
  textStyleName?: string;
  /** Relative frame width from the package style system. */
  relativeWidth?: string;
  /** Relative frame height from the package style system. */
  relativeHeight?: string;
  /** Horizontal offset from the cell in EMU. */
  x?: number;
  /** Vertical offset from the cell in EMU. */
  y?: number;
  /** Graphic width in EMU. */
  width?: number;
  /** Graphic height in EMU. */
  height?: number;
  /** Painting order; higher numbers render above lower numbers. */
  zIndex?: number;
  /** End anchor cell address, such as `D4`, for spreadsheet-anchored graphics. */
  endCellAddress?: string;
  /** Horizontal offset of the end anchor in EMU. */
  endX?: number;
  /** Vertical offset of the end anchor in EMU. */
  endY?: number;
}

/** Embedded object graphic anchored to a spreadsheet cell. */
export interface CellObjectGraphicOptions {
  /** Object graphic discriminant. */
  type: "object";
  /** Object package URI, without a leading `./` or slash. */
  href: string;
  /** Embedded chart subdocument. */
  chart?: ChartSpaceOptions;
  /** Accessible object name. */
  name?: string;
  /** Accessible object description. */
  description?: string;
  /** Drawing identifier. */
  id?: string;
  /** Package-local graphic style reference. */
  styleName?: string;
  /** Text style reference for labels attached to the frame. */
  textStyleName?: string;
  /** Relative frame width from the package style system. */
  relativeWidth?: string;
  /** Relative frame height from the package style system. */
  relativeHeight?: string;
  /** Horizontal offset from the cell in EMU. */
  x?: number;
  /** Vertical offset from the cell in EMU. */
  y?: number;
  /** Graphic width in EMU. */
  width?: number;
  /** Graphic height in EMU. */
  height?: number;
  /** Painting order; higher numbers render above lower numbers. */
  zIndex?: number;
  /** End anchor cell address, such as `D4`, for spreadsheet-anchored graphics. */
  endCellAddress?: string;
  /** Horizontal offset of the end anchor in EMU. */
  endX?: number;
  /** Vertical offset of the end anchor in EMU. */
  endY?: number;
}

export type CellGraphicOptions = CellImageGraphicOptions | CellObjectGraphicOptions;

/** A placeholder covering the non-origin cells of a merged range. */
export interface CoveredCellOptions {
  /** Covered-cell reference, such as `B2`. */
  reference: string;
  /** Declared source value type retained for round-trip. */
  valueType?: "float" | "string" | "date" | "boolean" | "percentage" | "currency" | "time";
  /** Placeholder text retained from the covered cell. */
  text?: string;
}
