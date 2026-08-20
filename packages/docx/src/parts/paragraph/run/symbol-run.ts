/**
 * SymbolRun types for WordprocessingML documents.
 *
 * This module provides support for inserting symbol characters into documents.
 *
 * @module
 */
import type { RunOptions } from "./run";

/**
 * Options for creating a SymbolRun.
 *
 * @see {@link SymbolRun}
 */
export type SymbolRunOptions = {
  /** The Unicode character code for the symbol */
  char: string;
  /** The font to use for the symbol (e.g., "Wingdings", "Symbol") */
  symbolFont?: string;
  /** Symbol element vocabulary; omitted emits the standard w:sym element. */
  kind?: "standard" | "office2016";
} & RunOptions;
