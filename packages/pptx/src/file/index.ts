export { File, type IPresentationOptions, type ISlideOptions } from "./file";
export { Presentation } from "./presentation/presentation";
export type { IPresentationOptions as IPresentationXmlOptions } from "./presentation/presentation";
export { Shape, type IShapeOptions } from "./shape/shape";
export { TextBody, type ITextBodyOptions } from "./shape/text-body";
export { Paragraph, type IParagraphOptions } from "./shape/paragraph/paragraph";
export { Run, type IRunOptions } from "./shape/paragraph/run";
export { Text } from "./shape/paragraph/text";
export { RunProperties, type IRunPropertiesOptions } from "./shape/paragraph/run-properties";
export {
    ParagraphProperties,
    type IParagraphPropertiesOptions,
} from "./shape/paragraph/paragraph-properties";
export { EndParagraphRunProperties } from "./shape/paragraph/end-paragraph-run";
export { SolidFill } from "./drawingml/solid-fill";
export { NoFill } from "./drawingml/no-fill";
export {
    GradientFill,
    type GradientFillOptions,
    type GradientStopOptions,
} from "./drawingml/gradient-fill";
export { Outline, type OutlineOptions } from "./drawingml/outline";
export { BlipFill } from "./drawingml/blip-fill";
export { Transform2D, type ITransform2DOptions } from "./drawingml/transform-2d";
export { PresetGeometry } from "./drawingml/preset-geometry";
export {
    ShapeProperties,
    type IShapePropertiesOptions,
    type ShapeFill,
} from "./drawingml/shape-properties";
export { NonVisualDrawingProperties } from "./drawingml/non-visual-drawing-props";
export { NonVisualShapeProperties } from "./drawingml/non-visual-shape-props";
export { NonVisualPictureProperties } from "./drawingml/non-visual-picture-props";
export { GroupShapeProperties } from "./drawingml/group-shape-properties";
export { GroupTransform2D, type IGroupTransform2DOptions } from "./drawingml/group-transform-2d";
export { GroupShape, type IGroupShapeOptions } from "./shape/group-shape";
export { Media } from "./media/media";
export { createTransformation, type IMediaTransformation } from "./media/media";
export type { IMediaData, IMediaDataTransformation } from "./media/data";
export { CoreProperties, type ICorePropertiesOptions } from "./core-properties/properties";
export { AppProperties } from "./app-properties/app-properties";
export { ContentTypes } from "./content-types/content-types";
export { Relationships } from "./relationships/relationships";
export {
    createRelationship,
    type RelationshipType,
} from "./relationships/relationship/relationship";
export { Slide } from "./slide/slide";
export { ShapeTree } from "./shape-tree/shape-tree";
export { DefaultTheme } from "./theme/theme";
export { DefaultSlideMaster } from "./slide-master/slide-master";
export { DefaultSlideLayout } from "./slide-layout/slide-layout";
export { DefaultNotesMaster } from "./notes-master/notes-master";
export { PresentationWrapper, type IViewWrapper } from "./presentation/presentation-wrapper";
export { Picture, type IPictureOptions } from "./picture/picture";
export { Background, type IBackgroundOptions } from "./background/background";
export { TableFrame, type ITableFrameOptions } from "./table/table-frame";
export { Table, type ITableOptions } from "./table/table";
export { TableRow, type ITableRowOptions } from "./table/table-row";
export { TableCell, type ITableCellOptions } from "./table/table-cell";
export { TableProperties } from "./table/table-properties";
export { TableCellProperties, type ICellBorderOptions } from "./table/table-cell-properties";
export { ChartFrame, type IChartFrameOptions } from "./chart/chart-frame";
export { ChartCollection, type IChartData } from "./chart/chart-collection";
export { ChartSpace, type IChartSpaceOptions, type IChartSeriesData } from "./chart/chart-space";
export type { ChartType } from "./chart/chart-types/create-chart-type";
