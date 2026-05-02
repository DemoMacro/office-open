/**
 * Convenience utility functions — re-exports from core + docx-specific generators.
 *
 * @module
 */
// Re-export from core
export { convertMillimetersToTwip, convertInchesToTwip } from "@office-open/core";
export { uniqueNumericIdCreator, uniqueId, hashedId, uniqueUuid } from "@office-open/core";
export type { UniqueNumericIdCreator } from "@office-open/core";

// Docx-specific generators
import { uniqueNumericIdCreator } from "@office-open/core";

export const abstractNumUniqueNumericIdGen = (): (() => number) => uniqueNumericIdCreator();

export const concreteNumUniqueNumericIdGen = (): (() => number) => uniqueNumericIdCreator(1);

export const docPropertiesUniqueNumericIdGen = (): (() => number) => uniqueNumericIdCreator();

export const bookmarkUniqueNumericIdGen = (): (() => number) => uniqueNumericIdCreator();
