// Public API — only exports used by consumers (docx, pptx patchers + their tests)

// Namespace configs
export { DOCX_NS, PPTX_NS } from "./xml-namespace";
export type { XmlNamespaceConfig } from "./xml-namespace";

// Core replacer (the main entry point for patch operations)
export { createReplacer } from "./xml-replacer";
export type { ReplacerConfig } from "./xml-replacer";

// Lower-level building blocks (used by patchers + specs)
export { createRunRenderer } from "./run-renderer";
export type { RenderedParagraphNode } from "./run-renderer";

export { createTraverser } from "./xml-traverser";

export { createTokenReplacer } from "./paragraph-token-replacer";

export { createSplitInject, TokenNotFoundError } from "./paragraph-split-inject";

export {
  toJson,
  createTextElementContents,
  patchSpaceAttribute,
  getFirstLevelElements,
} from "./xml-patch-utils";

export { appendContentType } from "./content-types-manager";

export { getNextRelationshipIndex, appendRelationship } from "./relationship-manager";
