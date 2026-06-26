/**
 * Media module exports.
 *
 * @module
 */
export * from "./media";
export * from "./data";
// Shared, content-deduplicated media collection — one implementation across all
// format packages (docx/pptx/xlsx), re-exported here so call sites read `Media`
// from their own package.
export { Media } from "@office-open/core";
