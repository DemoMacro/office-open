// Public patch runtime boundary shared by OOXML format packages.

// Namespace configs
export { DOCX_NS, PPTX_NS } from "./xml-namespace";
export type { XmlNamespaceOptions } from "./xml-namespace";

// Canonical patch command runtime (the shared editing model)
export { applyPatchCommands } from "./commands";
export type {
  PartialDeep,
  PatchAction,
  PatchAddress,
  PatchCollectionItem,
  PatchCommand,
  PatchInsertCommand,
  PatchField,
  PatchMoveCommand,
  PatchNestedValue,
  PatchNodeAddress,
  PatchPathStep,
  PatchPosition,
  PatchPropertyPath,
  PatchRemoveCommand,
  PatchReplaceCommand,
  PatchRouteStep,
  PatchTextAddress,
  PatchTextSelection,
  PatchTextReplaceCommand,
  PatchTextRemoveCommand,
  PatchUpdateValue,
  PatchUpdateCommand,
  RuntimePatchCommand,
} from "./commands";

// Core replacer (lower-level XML helper retained for package tooling)
export { createReplacer } from "./xml-replacer";
export type { ReplacerOptions } from "./xml-replacer";

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
  nextNumericId,
} from "./xml-patch-utils";

export { appendContentType, appendOverride, removeOverride } from "./content-types-manager";

export { getNextRelationshipIndex, appendRelationship } from "./relationship-manager";

// Core-properties metadata override (used by patch coreProperties layers)
export { applyCorePropertiesOverride } from "./core-properties-override";

// Shared patch option types (extended by each format's patch options)
export type { PatchOptions, PlaceholderDelimiters } from "./types";
export type { PatchDiagnostic, PatchPart, PatchPath, PatchReason } from "./diagnostics";
export { PatchError } from "./diagnostics";
export { PatchContext, assertPatchDelimiters, unsupportedPatchFormat } from "./context";
