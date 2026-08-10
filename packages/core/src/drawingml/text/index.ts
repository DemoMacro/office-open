/**
 * Shared DrawingML text descriptors — runs, paragraphs, fields, list styles,
 * and text body properties. Promoted from per-format implementations so
 * DOCX/PPTX/XLSX share one text model.
 *
 * @module
 */

export * from "./types";
export * from "./body-properties";
export * from "./run-properties";
export * from "./run";
export * from "./paragraph";
export * from "./text-body";
export * from "./list-style";
