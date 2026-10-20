import type {
  BlackWhiteMode,
  NonVisualContentPartPropertiesOptions,
  NonVisualDrawingPropertiesOptions,
  Transform2DOptions,
} from "@office-open/core/drawing";

export type { BlackWhiteMode };

/** PowerPoint 2010 p14:nvContentPartPr extension properties. */
export interface ContentPartNonVisualProperties {
  drawingProperties: NonVisualDrawingPropertiesOptions & { id: number | string };
  contentPartProperties?: NonVisualContentPartPropertiesOptions;
  properties?: { extensions?: string };
}

/** p:contentPart and its typed PowerPoint 2010 extension surface. */
export interface ContentPartOptions {
  rId: string;
  blackWhiteMode?: BlackWhiteMode;
  nonVisualProperties?: ContentPartNonVisualProperties;
  transform?: Transform2DOptions;
}
