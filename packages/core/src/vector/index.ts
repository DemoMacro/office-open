/**
 * Vector Markup Language (VML) — the transitional vector-shape vocabulary
 * (v:*) shared by docx textboxes/pictures-in-object, xlsx comment anchors and
 * pptx legacy drawings. Deprecated in favor of DrawingML by ISO/IEC 29500-4,
 * but still required to read and write real-world Office documents.
 *
 * @module
 */
export * from "./style";
export * from "./attributes";
export * from "./shapes";
export * from "./office-shape-defaults";
export * from "./shape-elements/fill";
export * from "./shape-elements/stroke";
export * from "./shape-elements/shadow";
export * from "./shape-elements/textbox";
export * from "./shape-elements/imagedata";
export * from "./shape-elements/path";
export * from "./shape-elements/textpath";
export * from "./shape-elements/formulas";
export * from "./shape-elements/handles";
export * from "./shape-elements/office-elements";
export * from "./shape-elements/word-elements";
export * from "./shape-elements/client-data";
export * from "./shape-elements/presentation-elements";
